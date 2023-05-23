using System.Text;
using Azure.AI.OpenAI;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using Microsoft.Extensions.Configuration;
using NAudio.Wave;
using YoutubeExplode;

namespace AgentMetrics
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // Create a configuration builder
            var builder = new ConfigurationBuilder()
                .SetBasePath(Environment.CurrentDirectory)
                .AddUserSecrets<Program>();
            
            // Build the configuration
            var configuration = builder.Build();
              
                
            // Get the temporary files folder
            var tempFolder = Path.GetTempPath();
            var audioFilePath = await DownloadAudioFileFromUriAsync(new Uri(args[0]), tempFolder);

            // Get the Azure Speech API key and region
            var subscriptionKey = configuration["AgentMetrics:SpeechSubscriptionKey"];
            var serviceRegion = configuration["AgentMetrics:SpeechServiceRegion"];

            var transcription = await TranscribeAudioFileWithAzureAsync(audioFilePath, subscriptionKey, serviceRegion);

            var openAIKey = configuration["AgentMetrics:OpenAIKey"];

            

            var annotatedTranscription = await CreateAnnotateTranscriptionAsync(transcription, openAIKey);

            var calculatedMetrics = await CalculateMetricsAsync(annotatedTranscription, openAIKey);
        }

        private static async Task<string> CalculateMetricsAsync(string annotatedTranscription, string? openAIKey)
        {

            var engine = "text-davinci-003";
            var prompt = @"
                    Given a call center transcript, rate the agent’s active listening skills. Active listening is focusing on a speaker, understanding their message, and responding thoughtfully. It involves being present, showing interest, noticing non-verbal cues, asking open-ended questions, paraphrasing and reflecting back, listening to understand and not to respond, and withholding judgment and advice. For example:

                    Customer: Hi, I have a problem with my internet connection. Agent: I’m sorry to hear that. That must be frustrating. (paraphrasing and empathizing) Customer: Yes, it is. I need the internet for my work. Agent: I understand. Can you tell me more about the problem? (asking open-ended questions) Customer: It started two weeks ago and it happens almost every day. Agent: I see. So it’s not constant but frequent. (reflecting back) Customer: Exactly. Agent: Okay, thank you. I’m going to run some tests on your line. Please stay on the line. (listening to understand and respond)

                    The possible ratings are:

                    Excellent: The agent uses all or most of the techniques consistently and effectively.
                    Good: The agent uses some of the techniques frequently and appropriately.
                    Fair: The agent uses a few of the techniques occasionally or inconsistently.
                    Poor: The agent uses none or very few of the techniques.
                    The output should be a JSON object with one property named “activeListening” with one of the ratings. For example:

                \{'activeListening': 'Excellent'\}

                Now do for this transcript:
                                
                ";
            prompt += annotatedTranscription;

            Console.WriteLine($"Prompt length: {prompt.Length}");

            var openAIClient = new OpenAIClient(openAIKey);
            var result = await openAIClient.GetCompletionsAsync(engine, new CompletionsOptions {Prompts = {prompt}, MaxTokens = (4097 - prompt.Length) < 0 ? 4097 : 4097 - prompt.Length, Temperature = 0.0f});

            var calculatedMetrics = result.Value.Choices[0].Text;

            Console.WriteLine($"CALCULATED TRANSCRIPTION: \n {calculatedMetrics}");
            return calculatedMetrics;
        }

        private static async Task<string> CreateAnnotateTranscriptionAsync(string transcription, string? openAIKey)
        {
            var engine = "text-davinci-003";
            var prompt = @"Transcribe the following conversation into a JSON array. The JSON array should contain items with each entry being a JSON object with one property, property name being 'agent' or 'customer' and value as the text of the conversation.";
            prompt = @$"
                {prompt}
                
                {transcription}";

            Console.WriteLine($"Prompt length: {prompt.Length}");

            var openAIClient = new OpenAIClient(openAIKey);
            var result = await openAIClient.GetCompletionsAsync(engine, new CompletionsOptions {Prompts = {prompt}, MaxTokens = (4097 - prompt.Length) < 0 ? 4097 : 4097 - prompt.Length, Temperature = 0.0f});

            var annotatedTranscription = result.Value.Choices[0].Text;

            Console.WriteLine($"ANNOTATED TRANSCRIPTION: \n {annotatedTranscription}");
            return annotatedTranscription;
        }

        private static async Task<string> TranscribeAudioFileWithAzureAsync(string audioFilePath, string? subscriptionKey, string? serviceRegion)
        {
            // Create an instance of the Azure Speech API
            var speechConfig = SpeechConfig.FromSubscription(subscriptionKey, serviceRegion);
            speechConfig.RequestWordLevelTimestamps();
            using var audioConfig = AudioConfig.FromWavFileInput(audioFilePath);
            using var speechRecognizer = new SpeechRecognizer(speechConfig, audioConfig);
            var stopRecognition = new TaskCompletionSource<int>();

            var transcription = new StringBuilder();

            speechRecognizer.Recognized += (s, e) =>
            {
                if (e.Result.Reason == ResultReason.RecognizedSpeech)
                {
                    //transcription.Append($"O:{e.Result.OffsetInTicks/10000} D:{e.Result.Duration.TotalMilliseconds} T:{e.Result.Text}");
                    transcription.Append($"{e.Result.Text}");
                    transcription.Append("\n");
                }
                else if (e.Result.Reason == ResultReason.NoMatch)
                {
                    Console.WriteLine($"NOMATCH: Speech could not be recognized.");
                }
            };

            speechRecognizer.SpeechEndDetected += (s, e) =>
            {                
                Console.WriteLine($"Speech ended.");                
            };

            speechRecognizer.Canceled += (s, e) =>
            {
                Console.WriteLine($"CANCELED: Reason={e.Reason}");

                if (e.Reason == CancellationReason.Error)
                {
                    Console.WriteLine($"CANCELED: ErrorCode={e.ErrorCode}");
                    Console.WriteLine($"CANCELED: ErrorDetails={e.ErrorDetails}");
                    Console.WriteLine($"CANCELED: Did you update the subscription info?");
                }
                stopRecognition.TrySetResult(0);
            };

            speechRecognizer.SessionStopped += (s, e) =>
            {
                Console.WriteLine("\n    Session stopped event.");
                Console.WriteLine("\nStop recognition.");
                stopRecognition.TrySetResult(0);
            };

            await speechRecognizer.StartContinuousRecognitionAsync().ConfigureAwait(false);

            Task.WaitAny(new[] { stopRecognition.Task });

            await speechRecognizer.StopContinuousRecognitionAsync().ConfigureAwait(false);

            var fullText = transcription.ToString();
            Console.WriteLine($"FULL TEXT: \n {fullText}");
            return fullText;
        }

        private static async Task<string> DownloadAudioFileFromUriAsync(Uri videoUri, string downloadFolder)
        {
            var youTube = new YoutubeClient();
            var manifest = await youTube.Videos.Streams.GetManifestAsync(videoUri.ToString());
            var streamInfo = manifest.GetAudioOnlyStreams().OrderByDescending(x => x.Bitrate).FirstOrDefault();

            if (streamInfo == default)
            {
                throw new Exception("No suitable audio stream found for this video.");
            }

            var extension = streamInfo.Container.Name;

            var fileName = $"{downloadFolder}{Guid.NewGuid()}.{extension}";

            await youTube.Videos.Streams.DownloadAsync(streamInfo, fileName);

            switch(extension)
            {
                case "mp4":
                    using(var reader = new MediaFoundationReader(fileName))
                    {
                        WaveFileWriter.CreateWaveFile(fileName.Replace(".mp4", ".wav"), reader);
                        return fileName.Replace(".mp4", ".wav");
                    }
                case "m4a":
                    extension = "m4a";
                    break;
                default:
                    throw new Exception($"Unknown audio container format: {extension}");
            }

            Console.WriteLine($"Video downloaded to: {fileName}");

            return fileName;
        }
    }
}

