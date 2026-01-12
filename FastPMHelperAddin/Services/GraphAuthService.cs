using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace FastPMHelperAddin.Services
{
    public class GraphAuthService
    {
        private readonly IPublicClientApplication _msalClient;
        private readonly string[] _scopes = new[] { "User.Read", "Sites.ReadWrite.All" };

        public GraphAuthService(string tenantId, string clientId)
        {
            _msalClient = PublicClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithRedirectUri("http://localhost")
                .Build();
        }

        public async Task<string> GetAccessTokenAsync(CancellationToken ct = default)
        {
            try
            {
                // Try silent acquisition first (from cache)
                var accounts = await _msalClient.GetAccountsAsync();
                var firstAccount = accounts.FirstOrDefault();

                if (firstAccount != null)
                {
                    try
                    {
                        var result = await _msalClient
                            .AcquireTokenSilent(_scopes, firstAccount)
                            .ExecuteAsync(ct);

                        System.Diagnostics.Debug.WriteLine("Token acquired silently from cache");
                        return result.AccessToken;
                    }
                    catch (MsalUiRequiredException)
                    {
                        System.Diagnostics.Debug.WriteLine("Silent token acquisition failed, requiring interactive login");
                    }
                }

                // Interactive acquisition with parent window
                var interactiveResult = await _msalClient
                    .AcquireTokenInteractive(_scopes)
                    .WithParentActivityOrWindow(GetOutlookWindowHandle())
                    .WithPrompt(Prompt.SelectAccount)
                    .ExecuteAsync(ct);

                System.Diagnostics.Debug.WriteLine($"Token acquired interactively for: {interactiveResult.Account.Username}");
                return interactiveResult.AccessToken;
            }
            catch (MsalException ex)
            {
                System.Diagnostics.Debug.WriteLine($"MSAL authentication error: {ex.Message}");
                throw new Exception($"Authentication failed: {ex.Message}", ex);
            }
        }

        private IntPtr GetOutlookWindowHandle()
        {
            try
            {
                var outlookApp = Globals.ThisAddIn.Application;
                var activeExplorer = outlookApp.ActiveExplorer();
                if (activeExplorer != null)
                {
                    var hwndProp = activeExplorer.GetType().GetProperty("HWND");
                    if (hwndProp != null)
                    {
                        var hwndObj = hwndProp.GetValue(activeExplorer, null);
                        if (hwndObj != null)
                            return new IntPtr(Convert.ToInt64(hwndObj));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to get Outlook window handle: {ex.Message}");
            }
            return IntPtr.Zero;
        }
    }
}
