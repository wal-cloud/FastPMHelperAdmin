using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace FastPMHelperAddin.Services
{
    public class LLMService
    {
        private readonly string _apiKey;
        private readonly HttpClient _httpClient;
        private const string GeminiEndpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent";

        public LLMService(string apiKey)
        {
            _apiKey = apiKey;
            _httpClient = new HttpClient();
        }

        public async Task<LLMExtractionResult> GetExtractionAsync(string emailBody,
            string sender, string subject)
        {
            // Extract only current and prior email (last 2 messages in thread)
            emailBody = ExtractRecentMessages(emailBody, 2);

            var prompt = BuildExtractionPrompt(emailBody, sender, subject);
            var responseText = await CallGeminiAsync(prompt);
            return ParseExtractionResponse(responseText);
        }

        public async Task<LLMDeltaResult> GetDeltaAsync(string emailBody,
            string currentContext, string currentBallHolder)
        {
            // Extract only current and prior email (last 2 messages in thread)
            emailBody = ExtractRecentMessages(emailBody, 2);

            var prompt = BuildDeltaPrompt(emailBody, currentContext, currentBallHolder);
            var responseText = await CallGeminiAsync(prompt);
            return ParseDeltaResponse(responseText);
        }

        private string BuildExtractionPrompt(string emailBody, string sender, string subject)
        {
            return $@"You are analyzing an email to extract action items for project management.

EMAIL SUBJECT: {subject}
FROM: {sender}
BODY:
{emailBody}

Extract the following information and respond in JSON format:
{{
  ""title"": ""A concise action title (50 chars max)"",
  ""ballHolder"": ""Person responsible - MUST be first and last name only (e.g., 'John Smith', not titles, emails, or extra text)"",
  ""description"": ""2-3 sentence summary of what needs to be done""
}}

Focus on:
- Action items, requests, or commitments
- Who is responsible (extract ONLY first and last name, no titles like 'Mr.', 'Dr.', no email addresses)
- Key deadlines or urgency

Respond ONLY with valid JSON, no additional text.";
        }

        private string BuildDeltaPrompt(string emailBody, string currentContext,
            string currentBallHolder)
        {
            return $@"You are analyzing a new email in an existing action thread.

CURRENT ACTION CONTEXT:
{currentContext}

CURRENT BALL HOLDER: {currentBallHolder}

NEW EMAIL BODY:
{emailBody}

Determine what changed and respond in JSON format:
{{
  ""newBallHolder"": ""Updated responsible person - MUST be first and last name only (e.g., 'John Smith', not titles, emails, or extra text). If no change, keep current."",
  ""updateSummary"": ""2-3 sentence summary of what changed or progressed"",
  ""isComplete"": false
}}

Look for:
- Status updates or progress
- Responsibility transfers (extract ONLY first and last name, no titles like 'Mr.', 'Dr.', no email addresses)
- New information or blockers
- Completion signals

Respond ONLY with valid JSON, no additional text.";
        }

        private async Task<string> CallGeminiAsync(string prompt)
        {
            try
            {
                var requestBody = new
                {
                    contents = new[]
                    {
                        new
                        {
                            parts = new[] { new { text = prompt } }
                        }
                    },
                    generationConfig = new
                    {
                        temperature = 0.3,
                        maxOutputTokens = 4096
                    }
                };

                var json = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var url = $"{GeminiEndpoint}?key={_apiKey}";
                var response = await _httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();

                var responseJson = await response.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"Gemini API Response: {responseJson}");

                var parsed = JObject.Parse(responseJson);

                // Extract text from Gemini response structure
                var text = parsed["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString();
                System.Diagnostics.Debug.WriteLine($"Extracted LLM text: {text}");

                return text ?? string.Empty;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LLM API error: {ex.Message}");
                throw;
            }
        }

        private string ExtractRecentMessages(string emailBody, int messageCount)
        {
            if (string.IsNullOrWhiteSpace(emailBody))
                return emailBody;

            // Common email thread delimiters
            string[] delimiters = new[]
            {
                "-----Original Message-----",
                "________________________________",
                "\nFrom:",
                "\nOn "  // Matches "On [date], [person] wrote:"
            };

            // Find all delimiter positions
            var positions = new List<int>();
            foreach (var delimiter in delimiters)
            {
                int index = 0;
                while ((index = emailBody.IndexOf(delimiter, index, StringComparison.OrdinalIgnoreCase)) != -1)
                {
                    positions.Add(index);
                    index += delimiter.Length;
                }
            }

            // If no delimiters found, return the whole body (truncated if needed)
            if (positions.Count == 0)
            {
                return emailBody.Length > 4000
                    ? emailBody.Substring(0, 4000) + "\n\n[Truncated...]"
                    : emailBody;
            }

            // Sort positions to get them in order
            positions.Sort();

            // Take the first N-1 delimiters (to get N messages)
            // The most recent message is before the first delimiter
            int cutoffPosition;
            if (positions.Count >= messageCount)
            {
                // Get the Nth delimiter position
                cutoffPosition = positions[messageCount - 1];
            }
            else
            {
                // Not enough delimiters, take the whole body
                cutoffPosition = emailBody.Length;
            }

            // Extract from start to cutoff
            string extracted = emailBody.Substring(0, cutoffPosition).Trim();

            // Ensure it's not too long (max 4000 chars for safety)
            if (extracted.Length > 4000)
            {
                extracted = extracted.Substring(0, 4000) + "\n\n[Truncated...]";
            }

            System.Diagnostics.Debug.WriteLine($"Extracted {messageCount} recent messages: {extracted.Length} chars (from original {emailBody.Length} chars)");

            return extracted;
        }

        private LLMExtractionResult ParseExtractionResponse(string responseText)
        {
            try
            {
                // Remove markdown code blocks if present
                responseText = responseText.Trim();
                if (responseText.StartsWith("```json"))
                    responseText = responseText.Substring(7);
                if (responseText.StartsWith("```"))
                    responseText = responseText.Substring(3);
                if (responseText.EndsWith("```"))
                    responseText = responseText.Substring(0, responseText.Length - 3);

                var json = JObject.Parse(responseText.Trim());
                return new LLMExtractionResult
                {
                    Title = json["title"]?.ToString() ?? "New Action",
                    BallHolder = json["ballHolder"]?.ToString() ?? "Unknown",
                    Description = json["description"]?.ToString() ?? ""
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LLM parse error: {ex.Message}");
                return new LLMExtractionResult
                {
                    Title = "Parse Error",
                    BallHolder = "Unknown",
                    Description = responseText
                };
            }
        }

        private LLMDeltaResult ParseDeltaResponse(string responseText)
        {
            try
            {
                responseText = responseText.Trim();
                if (responseText.StartsWith("```json"))
                    responseText = responseText.Substring(7);
                if (responseText.StartsWith("```"))
                    responseText = responseText.Substring(3);
                if (responseText.EndsWith("```"))
                    responseText = responseText.Substring(0, responseText.Length - 3);

                var json = JObject.Parse(responseText.Trim());
                return new LLMDeltaResult
                {
                    NewBallHolder = json["newBallHolder"]?.ToString() ?? "Unknown",
                    UpdateSummary = json["updateSummary"]?.ToString() ?? "",
                    IsComplete = json["isComplete"]?.ToObject<bool>() ?? false
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LLM parse error: {ex.Message}");
                return new LLMDeltaResult
                {
                    NewBallHolder = "Unknown",
                    UpdateSummary = responseText,
                    IsComplete = false
                };
            }
        }
    }

    public class LLMExtractionResult
    {
        public string Title { get; set; }
        public string BallHolder { get; set; }
        public string Description { get; set; }
    }

    public class LLMDeltaResult
    {
        public string NewBallHolder { get; set; }
        public string UpdateSummary { get; set; }
        public bool IsComplete { get; set; }
    }
}
