using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace FastPMHelperAddin.Services
{
    public class GoogleSheetsAuthService
    {
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _applicationName;
        private readonly string _tokenCacheDir;
        private UserCredential _credential;

        public GoogleSheetsAuthService(string clientId, string clientSecret, string applicationName, string tokenCacheDir)
        {
            _clientId = clientId;
            _clientSecret = clientSecret;
            _applicationName = applicationName;

            // Expand environment variables in token cache dir
            _tokenCacheDir = Environment.ExpandEnvironmentVariables(tokenCacheDir);

            // Ensure directory exists
            if (!Directory.Exists(_tokenCacheDir))
                Directory.CreateDirectory(_tokenCacheDir);

            System.Diagnostics.Debug.WriteLine($"GoogleSheetsAuthService initialized:");
            System.Diagnostics.Debug.WriteLine($"  Application Name: {_applicationName}");
            System.Diagnostics.Debug.WriteLine($"  Client ID: {_clientId?.Substring(0, Math.Min(20, _clientId?.Length ?? 0))}...");
            System.Diagnostics.Debug.WriteLine($"  Client Secret Length: {_clientSecret?.Length ?? 0}");
            System.Diagnostics.Debug.WriteLine($"  Token cache directory: {_tokenCacheDir}");
        }

        public async Task<SheetsService> GetSheetsServiceAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                if (_credential == null)
                {
                    System.Diagnostics.Debug.WriteLine("Initializing Google OAuth2 authentication...");

                    var clientSecrets = new ClientSecrets
                    {
                        ClientId = _clientId,
                        ClientSecret = _clientSecret
                    };

                    // Authorize using Desktop app flow
                    _credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                        clientSecrets,
                        new[] { SheetsService.Scope.Spreadsheets },
                        "user",
                        cancellationToken,
                        new FileDataStore(_tokenCacheDir, true)
                    );

                    System.Diagnostics.Debug.WriteLine($"Authentication successful for user: {_credential.UserId}");
                }

                // Create Sheets service
                var service = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = _credential,
                    ApplicationName = _applicationName
                });

                return service;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Google authentication error: {ex.Message}");
                throw new Exception($"Failed to authenticate with Google: {ex.Message}", ex);
            }
        }

        public void ClearCredentials()
        {
            _credential = null;

            // Delete token files
            if (Directory.Exists(_tokenCacheDir))
            {
                try
                {
                    Directory.Delete(_tokenCacheDir, true);
                    System.Diagnostics.Debug.WriteLine("Token cache cleared");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to clear token cache: {ex.Message}");
                }
            }
        }
    }
}
