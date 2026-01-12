using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace FastPMHelperAddin.Configuration
{
    public class ConfigurationManager
    {
        private static ConfigurationManager _instance;
        private Dictionary<string, string> _config;

        public static ConfigurationManager Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ConfigurationManager();
                return _instance;
            }
        }

        private ConfigurationManager()
        {
            LoadConfiguration();
        }

        private void LoadConfiguration()
        {
            _config = new Dictionary<string, string>();

            try
            {
                // Read .env from embedded resource (for VSTO add-ins)
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var resourceName = "FastPMHelperAddin..env";

                System.Diagnostics.Debug.WriteLine($"Loading .env from embedded resource: {resourceName}");

                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        // List available resources for debugging
                        var resources = assembly.GetManifestResourceNames();
                        System.Diagnostics.Debug.WriteLine($"Available embedded resources: {string.Join(", ", resources)}");
                        System.Diagnostics.Debug.WriteLine($"WARNING: Could not find embedded resource: {resourceName}");
                        return;
                    }

                    using (var reader = new System.IO.StreamReader(stream, System.Text.Encoding.UTF8))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                                continue;

                            // Remove BOM if present at start of line
                            string cleanLine = line.TrimStart('\uFEFF', '\u200B');

                            var parts = cleanLine.Split(new[] { '=' }, 2);
                            if (parts.Length == 2)
                            {
                                string key = parts[0].Trim();
                                string value = parts[1].Trim();

                                // Remove quotes if present
                                if (value.StartsWith("\"") && value.EndsWith("\""))
                                    value = value.Substring(1, value.Length - 2);

                                _config[key] = value;
                                System.Diagnostics.Debug.WriteLine($"Config loaded: {key} = {(key.Contains("KEY") || key.Contains("PASSWORD") ? "***" : value)}");
                            }
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Loaded {_config.Count} configuration values from embedded resource");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading .env from embedded resource: {ex.Message}");
            }
        }

        public string GetValue(string key, string defaultValue = "")
        {
            return _config.TryGetValue(key, out string value) ? value : defaultValue;
        }

        public string GeminiApiKey => GetValue("GEMINI_API_KEY");
        public string SharePointSiteUrl => GetValue("SHAREPOINT_SITE_URL");
        public string SharePointUsername => GetValue("SHAREPOINT_USERNAME");
        public string SharePointPassword => GetValue("SHAREPOINT_PASSWORD");
        public string AzureTenantId => GetValue("AZURE_TENANT_ID");
        public string AzureClientId => GetValue("AZURE_CLIENT_ID");

        // Google Sheets configuration
        public string GoogleClientId => GetValue("GOOGLE_CLIENT_ID");
        public string GoogleClientSecret => GetValue("GOOGLE_CLIENT_SECRET");
        public string GoogleSpreadsheetsId => GetValue("GOOGLE_SHEETS_SPREADSHEET_ID");
        public string GoogleSheetName => GetValue("GOOGLE_SHEETS_SHEET_NAME", "ProjectActions");
        public string GoogleAppName => GetValue("GOOGLE_APP_NAME", "FastPMHelperAddin");
        public string GoogleTokenCacheDir => GetValue("GOOGLE_TOKEN_CACHE_DIR", @"%APPDATA%\FastPMHelperAddin\tokens");

        public void ValidateGraphConfiguration()
        {
            var missingKeys = new List<string>();

            if (string.IsNullOrEmpty(SharePointSiteUrl))
                missingKeys.Add("SHAREPOINT_SITE_URL");
            if (string.IsNullOrEmpty(AzureTenantId))
                missingKeys.Add("AZURE_TENANT_ID");
            if (string.IsNullOrEmpty(AzureClientId))
                missingKeys.Add("AZURE_CLIENT_ID");

            if (missingKeys.Any())
            {
                throw new InvalidOperationException(
                    $"Missing required configuration keys: {string.Join(", ", missingKeys)}. " +
                    "Please check your .env file in the bin\\Debug folder.");
            }
        }

        public void ValidateGoogleSheetsConfiguration()
        {
            var missingKeys = new List<string>();

            if (string.IsNullOrEmpty(GoogleClientId))
                missingKeys.Add("GOOGLE_CLIENT_ID");
            if (string.IsNullOrEmpty(GoogleClientSecret))
                missingKeys.Add("GOOGLE_CLIENT_SECRET");
            if (string.IsNullOrEmpty(GoogleSpreadsheetsId))
                missingKeys.Add("GOOGLE_SHEETS_SPREADSHEET_ID");

            if (missingKeys.Any())
            {
                throw new InvalidOperationException(
                    $"Missing required Google Sheets configuration keys: {string.Join(", ", missingKeys)}. " +
                    "Please check your .env file. See README.md for setup instructions.");
            }
        }
    }
}
