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
            // Extract and split the last 2 messages
            var (latestEmail, previousEmail) = SplitRecentMessages(emailBody, 2);

            var prompt = BuildExtractionPrompt(latestEmail, previousEmail, sender, subject);
            var responseText = await CallGeminiAsync(prompt);
            return ParseExtractionResponse(responseText);
        }

        /// <summary>
        /// Extract multiple distinct actions from an email
        /// </summary>
        public async Task<List<LLMExtractionResult>> GetMultipleExtractionsAsync(string emailBody,
            string sender, string subject)
        {
            // Extract and split the last 2 messages
            var (latestEmail, previousEmail) = SplitRecentMessages(emailBody, 2);

            var prompt = BuildMultipleExtractionsPrompt(latestEmail, previousEmail, sender, subject);
            var responseText = await CallGeminiAsync(prompt);
            return ParseMultipleExtractionResponse(responseText);
        }

        public async Task<LLMDeltaResult> GetDeltaAsync(string emailBody,
            string currentContext, string currentBallHolder)
        {
            // Extract and split the last 2 messages
            var (latestEmail, previousEmail) = SplitRecentMessages(emailBody, 2);

            var prompt = BuildDeltaPrompt(latestEmail, previousEmail, currentContext, currentBallHolder);
            var responseText = await CallGeminiAsync(prompt);
            return ParseDeltaResponse(responseText);
        }

        public async Task<string> GetClosureSummaryAsync(string emailBody, string actionContext)
        {
            // Extract and split the last 2 messages
            var (latestEmail, previousEmail) = SplitRecentMessages(emailBody, 2);

            var prompt = BuildClosurePrompt(latestEmail, previousEmail, actionContext);
            var responseText = await CallGeminiAsync(prompt);
            return ParseClosureSummary(responseText);
        }

        private string BuildExtractionPrompt(string latestEmail, string previousEmail, string sender, string subject)
        {
            string contextSection = string.IsNullOrWhiteSpace(previousEmail)
                ? ""
                : $@"
PREVIOUS EMAIL (for context - can be used to clarify what the action is about):
{previousEmail}
---END OF PREVIOUS EMAIL---
";

            return $@"You are analyzing an email to extract action items for project management.

EMAIL SUBJECT: {subject}
FROM: {sender}

{contextSection}
LATEST EMAIL (*** PRIMARY SOURCE - action intent and ballHolder come from here ***):
{latestEmail}
---END OF LATEST EMAIL---

CRITICAL INSTRUCTIONS - EMAIL PRIORITY:
- The LATEST EMAIL is the PRIMARY and MOST IMPORTANT source
- ballHolder MUST come from the LATEST EMAIL (who is being asked to do something NOW)
- The action intent/request comes from the LATEST EMAIL (what is being asked NOW)
- HOWEVER, you CAN use the PREVIOUS EMAIL to fill in missing context about WHAT the action is about

WEIGHTING RULES:
1. BallHolder: 100% from LATEST EMAIL (who is responsible based on the most recent message)
2. Action Intent: 100% from LATEST EMAIL (what is being requested - review, provide, approve, etc.)
3. Subject Matter/Context: Can use PREVIOUS EMAIL if LATEST is vague (what specifically needs to be reviewed/provided)

CRITICAL - SINGLE ACTION REPRESENTING ALL ITEMS:
- You MUST return EXACTLY ONE action object (not an array)
- If the email contains MULTIPLE action items, create ONE action that represents ALL of them together
- The title and description should summarize ALL action items found in the email

TITLE STRATEGY:
- If 1 action: Use specific title (e.g., ""Review structural drawings"")
- If 2-3 actions: List them briefly (e.g., ""Review drawings & provide calculations"")
- If 4+ actions: Use summary title (e.g., ""Kick-off meeting actions"", ""Project coordination tasks"", ""Outstanding action items"")

DESCRIPTION STRATEGY:
- If 1-3 actions: List each action item as bullet points or numbered list
- If 4+ actions: Summarize the overall scope and list key items

BALLHOLDER STRATEGY:
- If all actions have SAME person: Use that person's name
- If actions have DIFFERENT people: Use the primary coordinator or ""Multiple parties""

DO NOT return a JSON array - return a single JSON object only.

INCORRECT (DO NOT DO THIS):
[
  {{""title"": ""Action 1"", ...}},
  {{""title"": ""Action 2"", ...}}
]

CORRECT EXAMPLES:

Example 1 (Single action):
{{
  ""title"": ""Review structural drawings"",
  ""ballHolder"": ""John Smith"",
  ""description"": ""Review the attached structural drawings and provide feedback by Friday.""
}}

Example 2 (2-3 actions):
{{
  ""title"": ""Review drawings & confirm specifications"",
  ""ballHolder"": ""John Smith"",
  ""description"": ""1) Review structural drawings, 2) Confirm painting specifications for CS flange, 3) Provide feedback on design approach.""
}}

Example 3 (4+ actions):
{{
  ""title"": ""Kick-off meeting actions"",
  ""ballHolder"": ""Multiple parties"",
  ""description"": ""Following kick-off meeting, multiple action items assigned: UGL to review DMW sequence and confirm hardness testing approach. DNM to confirm painting requirements, fastener options, procedure qualifications, and provide drawings in ACAD format.""
}}

Extract the following information and respond in JSON format:
{{
  ""title"": ""Concise title (50 chars max) summarizing ALL actions - use strategies above based on number of items"",
  ""ballHolder"": ""Person responsible - MUST be first and last name only (e.g., 'John Smith'), or 'Multiple parties' if different people. NO titles, emails, or extra text."",
  ""description"": ""Summary of ALL action items - list individually if 1-3 items, summarize scope if 4+ items""
}}

Focus on:
- Identify ALL action items in the LATEST EMAIL
- Summarize them appropriately based on quantity (list if few, summarize if many)
- Extract ballHolder(s) from LATEST EMAIL - use 'Multiple parties' if different people for different items
- Use PREVIOUS EMAIL for context if LATEST is vague
- Include any deadlines mentioned

Respond with EXACTLY ONE JSON object (not an array), no additional text.";
        }

        private string BuildMultipleExtractionsPrompt(string latestEmail, string previousEmail, string sender, string subject)
        {
            string contextSection = string.IsNullOrWhiteSpace(previousEmail)
                ? ""
                : $@"
PREVIOUS EMAIL (for context - can be used to clarify what the actions are about):
{previousEmail}
---END OF PREVIOUS EMAIL---
";

            return $@"You are analyzing an email to extract ALL distinct action items for project management.

EMAIL SUBJECT: {subject}
FROM: {sender}

{contextSection}
LATEST EMAIL (*** PRIMARY SOURCE - action intents and ballHolders come from here ***):
{latestEmail}
---END OF LATEST EMAIL---

CRITICAL INSTRUCTIONS - MULTIPLE ACTIONS:
- The LATEST EMAIL is the PRIMARY and MOST IMPORTANT source
- Extract ALL distinct, separate action items from the LATEST EMAIL
- Each action MUST have a clear, specific deliverable or task
- DO NOT split a single action into multiple sub-tasks
- Each action MAY have a DIFFERENT ballHolder (person responsible)

WEIGHTING RULES (same as single action):
1. BallHolder: 100% from LATEST EMAIL (who is responsible for EACH action)
2. Action Intent: 100% from LATEST EMAIL (what is being requested for EACH action)
3. Subject Matter/Context: Can use PREVIOUS EMAIL if LATEST is vague

WHAT COUNTS AS SEPARATE ACTIONS:
✓ Different deliverables (""Review drawings"" vs ""Provide calculations"")
✓ Different responsible parties (""John review X"" vs ""Mary approve Y"")
✓ Different subjects (""Confirm painting spec"" vs ""Update PO with fasteners"")

WHAT IS NOT SEPARATE ACTIONS:
✗ Sub-steps of same task (""Review, markup, and return drawings"" = 1 action)
✗ Sequential phases (""Design then fabricate"" = 1 action if same person)
✗ Context clarifications (""Review structural drawings for tower"" = 1 action)

EXAMPLE EMAIL:
""Hi Team,
1. UGL - Review DMW sequence and confirm
2. DNM - Provide drawings in ACAD format
3. DNM - Confirm painting requirements for CS flange""

CORRECT EXTRACTION (3 separate actions):
[
  {{
    ""title"": ""Review DMW sequence"",
    ""ballHolder"": ""UGL Team"",
    ""description"": ""Review and confirm the sequence for DMW Hydrotest, P&P, and PWHT.""
  }},
  {{
    ""title"": ""Provide drawings in ACAD format"",
    ""ballHolder"": ""DNM Team"",
    ""description"": ""Provide all relevant drawings in ACAD format for project documentation.""
  }},
  {{
    ""title"": ""Confirm CS flange painting requirements"",
    ""ballHolder"": ""DNM Team"",
    ""description"": ""Confirm the specific painting requirements for the CS flange.""
  }}
]

Extract ALL distinct actions and respond with a JSON ARRAY:
[
  {{
    ""title"": ""Concise action title (50 chars max)"",
    ""ballHolder"": ""Person/Team responsible - first and last name or team name (e.g., 'John Smith' or 'UGL Team')"",
    ""description"": ""2-3 sentence summary of this specific action""
  }},
  ... (repeat for each action)
]

IMPORTANT:
- ballHolder extraction rules: ONLY first and last name (e.g., 'John Smith'), OR team name if individuals not specified (e.g., 'UGL', 'DNM Team')
- NO titles like 'Mr.', 'Dr.'
- NO email addresses
- If multiple people for same action, pick primary contact or use team name
- If only 1 action found, return array with 1 element: [{{""title"": ..., ...}}]
- If NO clear actions found, return empty array: []

Respond ONLY with valid JSON array, no additional text.";
        }

        private string BuildDeltaPrompt(string latestEmail, string previousEmail, string currentContext,
            string currentBallHolder)
        {
            string contextSection = string.IsNullOrWhiteSpace(previousEmail)
                ? ""
                : $@"
PREVIOUS EMAIL (for context - can clarify what the update is about):
{previousEmail}
---END OF PREVIOUS EMAIL---
";

            return $@"You are analyzing a new email in an existing action thread.

CURRENT ACTION CONTEXT:
{currentContext}

CURRENT BALL HOLDER: {currentBallHolder}

{contextSection}
LATEST EMAIL (*** PRIMARY SOURCE - responsibility and update intent come from here ***):
{latestEmail}
---END OF LATEST EMAIL---

CRITICAL INSTRUCTIONS - EMAIL PRIORITY:
- The LATEST EMAIL is the PRIMARY and MOST IMPORTANT source
- newBallHolder MUST come from the LATEST EMAIL (who is responsible NOW based on most recent message)
- The update intent comes from the LATEST EMAIL (what changed, what's the new status)
- HOWEVER, you CAN use the PREVIOUS EMAIL to understand WHAT the update is about if LATEST is vague

WEIGHTING RULES:
1. BallHolder: 100% from LATEST EMAIL (who is responsible based on the most recent message)
2. Update Intent/Status: 100% from LATEST EMAIL (what changed - completed, in progress, blocked, etc.)
3. Subject Matter/Context: Can use PREVIOUS EMAIL if LATEST is vague about what specifically changed

EXAMPLE:
- LATEST EMAIL: ""This is now complete""
- PREVIOUS EMAIL: ""Working on the structural drawings review""
- CORRECT extraction:
  - updateSummary: ""Structural drawings review is now complete"" (status from LATEST, ""structural drawings"" context from PREVIOUS)
  - newBallHolder: [person from LATEST if changed, otherwise keep '{currentBallHolder}']

Determine what changed and respond in JSON format:
{{
  ""newBallHolder"": ""Updated responsible person from LATEST EMAIL - MUST be first and last name only (e.g., 'John Smith', not titles, emails, or extra text). If no responsibility change in LATEST EMAIL, use '{currentBallHolder}'."",
  ""updateSummary"": ""2-3 sentence summary - update intent from LATEST, subject matter from PREVIOUS if LATEST is vague"",
  ""isComplete"": false
}}

Look for:
- Responsibility transfers FROM THE LATEST EMAIL (extract ONLY first and last name, no titles like 'Mr.', 'Dr.', no email addresses)
- Status updates/progress FROM THE LATEST EMAIL
- What specifically changed - can use PREVIOUS EMAIL to understand subject matter if LATEST is vague
- Completion signals FROM THE LATEST EMAIL

Respond ONLY with valid JSON, no additional text.";
        }

        private string BuildClosurePrompt(string latestEmail, string previousEmail, string actionContext)
        {
            string contextSection = string.IsNullOrWhiteSpace(previousEmail)
                ? ""
                : $@"
PREVIOUS EMAIL (for context):
{previousEmail}
---END OF PREVIOUS EMAIL---
";

            return $@"You are analyzing an email that is closing an action item. Generate a concise closure summary for the history log.

ACTION CONTEXT:
{actionContext}

{contextSection}
LATEST EMAIL (*** This email is closing the action ***):
{latestEmail}
---END OF LATEST EMAIL---

Generate a brief, professional closure summary (1-2 sentences max) that explains:
- What was resolved or completed
- Who confirmed or provided what (use first and last names only, no titles)
- Key outcome or deliverable if mentioned

EXAMPLES OF GOOD SUMMARIES:
- ""Byron confirmed structural drawings are approved. Wally provided the final calculations.""
- ""Review completed by John Smith. All comments addressed and package approved.""
- ""Maria Garcia confirmed installation is complete. Photos provided as requested.""

Return ONLY the closure summary text (no JSON, no quotes, no preamble). Keep it under 100 characters if possible.";
        }

        private (string latestEmail, string previousEmail) SplitRecentMessages(string emailBody, int messageCount)
        {
            if (string.IsNullOrWhiteSpace(emailBody))
                return (emailBody, string.Empty);

            // Common email thread delimiters
            string[] delimiters = new[]
            {
                "-----Original Message-----",
                "________________________________",
                "\nFrom:",
                "\nOn "  // Matches "On [date], [person] wrote:"
            };

            // Find the first delimiter position
            int firstDelimiterPos = -1;
            foreach (var delimiter in delimiters)
            {
                int pos = emailBody.IndexOf(delimiter, StringComparison.OrdinalIgnoreCase);
                if (pos != -1 && (firstDelimiterPos == -1 || pos < firstDelimiterPos))
                {
                    firstDelimiterPos = pos;
                }
            }

            // If no delimiter found, the whole body is the latest email
            if (firstDelimiterPos == -1)
            {
                string truncated = emailBody.Length > 4000
                    ? emailBody.Substring(0, 4000) + "\n\n[Truncated...]"
                    : emailBody;
                return (truncated, string.Empty);
            }

            // Split into latest and previous
            string latestEmail = emailBody.Substring(0, firstDelimiterPos).Trim();

            // Find the second delimiter for the previous email
            int secondDelimiterPos = -1;
            foreach (var delimiter in delimiters)
            {
                int pos = emailBody.IndexOf(delimiter, firstDelimiterPos + 1, StringComparison.OrdinalIgnoreCase);
                if (pos != -1 && (secondDelimiterPos == -1 || pos < secondDelimiterPos))
                {
                    secondDelimiterPos = pos;
                }
            }

            string previousEmail;
            if (secondDelimiterPos == -1)
            {
                // No second delimiter, take everything after first delimiter
                previousEmail = emailBody.Substring(firstDelimiterPos).Trim();
            }
            else
            {
                // Take content between first and second delimiter
                previousEmail = emailBody.Substring(firstDelimiterPos, secondDelimiterPos - firstDelimiterPos).Trim();
            }

            // Truncate if needed
            if (latestEmail.Length > 2000)
                latestEmail = latestEmail.Substring(0, 2000) + "\n\n[Truncated...]";

            if (previousEmail.Length > 2000)
                previousEmail = previousEmail.Substring(0, 2000) + "\n\n[Truncated...]";

            System.Diagnostics.Debug.WriteLine($"Split messages - Latest: {latestEmail.Length} chars, Previous: {previousEmail.Length} chars");

            return (latestEmail, previousEmail);
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

        private List<LLMExtractionResult> ParseMultipleExtractionResponse(string responseText)
        {
            var results = new List<LLMExtractionResult>();

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

                responseText = responseText.Trim();

                // Handle empty array case
                if (responseText == "[]")
                {
                    System.Diagnostics.Debug.WriteLine("LLM returned empty array - no actions found");
                    return results;
                }

                // Parse as JSON array
                var jsonArray = JArray.Parse(responseText);

                foreach (var item in jsonArray)
                {
                    if (item is JObject json)
                    {
                        results.Add(new LLMExtractionResult
                        {
                            Title = json["title"]?.ToString() ?? "New Action",
                            BallHolder = json["ballHolder"]?.ToString() ?? "Unknown",
                            Description = json["description"]?.ToString() ?? ""
                        });
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Parsed {results.Count} actions from LLM response");
                return results;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LLM multiple parse error: {ex.Message}");

                // Return single error action
                results.Add(new LLMExtractionResult
                {
                    Title = "Parse Error",
                    BallHolder = "Unknown",
                    Description = $"Failed to parse multiple actions: {ex.Message}\n\nRaw response:\n{responseText}"
                });
                return results;
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

        private string ParseClosureSummary(string responseText)
        {
            try
            {
                // Clean up the response
                responseText = responseText.Trim();

                // Remove markdown code blocks if present
                if (responseText.StartsWith("```"))
                    responseText = responseText.Substring(3);
                if (responseText.EndsWith("```"))
                    responseText = responseText.Substring(0, responseText.Length - 3);

                // Remove quotes if the LLM wrapped it
                responseText = responseText.Trim().Trim('"');

                // Limit length
                if (responseText.Length > 200)
                    responseText = responseText.Substring(0, 197) + "...";

                return string.IsNullOrWhiteSpace(responseText)
                    ? "Action closed"
                    : responseText;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Closure summary parse error: {ex.Message}");
                return "Action closed";
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
