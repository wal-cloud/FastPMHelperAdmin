using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Kiota.Abstractions;
using ActionItemModel = FastPMHelperAddin.Models.ActionItem;

namespace FastPMHelperAddin.Services
{
    public class GraphSharePointService
    {
        private readonly GraphAuthService _authService;
        private readonly GraphServiceClient _graphClient;
        private readonly string _siteUrl;
        private readonly string _listName = "ProjectActions";

        private string _siteId;
        private string _listId;

        public GraphSharePointService(GraphAuthService authService, string siteUrl)
        {
            _authService = authService;
            _siteUrl = siteUrl;

            // Create Graph client with custom auth provider
            var authProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(_authService));
            _graphClient = new GraphServiceClient(authProvider);
        }

        private class TokenProvider : IAccessTokenProvider
        {
            private readonly GraphAuthService _authService;

            public TokenProvider(GraphAuthService authService)
            {
                _authService = authService;
            }

            public AllowedHostsValidator AllowedHostsValidator => new AllowedHostsValidator();

            public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = null, System.Threading.CancellationToken cancellationToken = default)
            {
                return await _authService.GetAccessTokenAsync(cancellationToken);
            }
        }

        private string GetFieldValue(IDictionary<string, object> fields, string key)
        {
            if (fields != null && fields.TryGetValue(key, out var value))
            {
                return value?.ToString() ?? string.Empty;
            }
            return string.Empty;
        }

        private async Task<string> GetSiteIdAsync()
        {
            if (!string.IsNullOrEmpty(_siteId))
                return _siteId;

            try
            {
                var uri = new Uri(_siteUrl);
                var hostname = uri.Host;
                var serverRelativePath = uri.AbsolutePath;

                var site = await _graphClient.Sites[$"{hostname}:{serverRelativePath}"].GetAsync();
                _siteId = site.Id;

                System.Diagnostics.Debug.WriteLine($"Site ID resolved: {_siteId}");
                return _siteId;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error resolving site ID: {ex.Message}");
                throw new Exception($"Failed to resolve SharePoint site: {ex.Message}", ex);
            }
        }

        private async Task<string> GetListIdAsync()
        {
            if (!string.IsNullOrEmpty(_listId))
                return _listId;

            try
            {
                var siteId = await GetSiteIdAsync();

                var lists = await _graphClient.Sites[siteId].Lists
                    .GetAsync(config => config.QueryParameters.Filter = $"displayName eq '{_listName}'");

                if (lists?.Value == null || !lists.Value.Any())
                    throw new Exception($"List '{_listName}' not found on site");

                _listId = lists.Value.First().Id;

                System.Diagnostics.Debug.WriteLine($"List ID resolved: {_listId}");
                return _listId;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error resolving list ID: {ex.Message}");
                throw new Exception($"Failed to find list '{_listName}': {ex.Message}", ex);
            }
        }

        public async Task<List<ActionItemModel>> FetchOpenActionsAsync()
        {
            try
            {
                var siteId = await GetSiteIdAsync();
                var listId = await GetListIdAsync();

                var items = await _graphClient.Sites[siteId].Lists[listId].Items
                    .GetAsync(config =>
                    {
                        config.QueryParameters.Expand = new[] { "fields" };
                        config.QueryParameters.Filter = "fields/Status ne 'Closed'";
                    });

                var actionItems = new List<ActionItemModel>();
                foreach (var item in items.Value)
                {
                    var fields = item.Fields.AdditionalData;
                    actionItems.Add(new ActionItemModel
                    {
                        Id = int.Parse(item.Id),
                        Title = GetFieldValue(fields, "Title"),
                        Status = GetFieldValue(fields, "Status"),
                        BallHolder = GetFieldValue(fields, "BallHolder"),
                        HistoryLog = GetFieldValue(fields, "HistoryLog"),
                        LinkedThreadIDs = GetFieldValue(fields, "LinkedThreadIDs"),
                        ActiveMessageIDs = GetFieldValue(fields, "ActiveMessageIDs")
                    });
                }

                System.Diagnostics.Debug.WriteLine($"Fetched {actionItems.Count} open actions");
                return actionItems;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Graph fetch error: {ex.Message}");
                throw;
            }
        }

        public async Task<int> CreateActionAsync(string title, string ballHolder,
            string conversationId, string internetMessageId, string initialNote)
        {
            try
            {
                var siteId = await GetSiteIdAsync();
                var listId = await GetListIdAsync();

                var fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        ["Title"] = title,
                        ["Status"] = "Open",
                        ["BallHolder"] = ballHolder,
                        ["LinkedThreadIDs"] = conversationId,
                        ["ActiveMessageIDs"] = internetMessageId,
                        ["HistoryLog"] = $"[{DateTime.Now:yyyy-MM-dd HH:mm}] Created: {initialNote}"
                    }
                };

                var newItem = new ListItem { Fields = fields };
                var created = await _graphClient.Sites[siteId].Lists[listId].Items.PostAsync(newItem);

                System.Diagnostics.Debug.WriteLine($"Created action item with ID: {created.Id}");
                return int.Parse(created.Id);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Graph create error: {ex.Message}");
                throw;
            }
        }

        public async Task UpdateActionAsync(int itemId, string newMessageId,
            string ballHolder, string updateNote)
        {
            try
            {
                var siteId = await GetSiteIdAsync();
                var listId = await GetListIdAsync();

                // Fetch current item
                var item = await _graphClient.Sites[siteId].Lists[listId].Items[itemId.ToString()]
                    .GetAsync(config => config.QueryParameters.Expand = new[] { "fields" });

                var fields = item.Fields.AdditionalData;

                // Append to ActiveMessageIDs
                string currentIds = GetFieldValue(fields, "ActiveMessageIDs");
                var idList = ActionItemModel.ParseIds(currentIds);
                if (!idList.Contains(newMessageId))
                    idList.Add(newMessageId);

                // Append to HistoryLog
                string currentLog = GetFieldValue(fields, "HistoryLog");
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                string updatedLog = $"{currentLog}\n\n[{timestamp}] {updateNote}";

                // Update fields
                var updateFields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        ["ActiveMessageIDs"] = string.Join("; ", idList),
                        ["BallHolder"] = ballHolder,
                        ["HistoryLog"] = updatedLog
                    }
                };

                await _graphClient.Sites[siteId].Lists[listId].Items[itemId.ToString()].Fields
                    .PatchAsync(updateFields);

                System.Diagnostics.Debug.WriteLine($"Updated action item {itemId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Graph update error: {ex.Message}");
                throw;
            }
        }

        public async Task SetStatusAsync(int itemId, string newStatus)
        {
            try
            {
                var siteId = await GetSiteIdAsync();
                var listId = await GetListIdAsync();

                var updateFields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        ["Status"] = newStatus
                    }
                };

                await _graphClient.Sites[siteId].Lists[listId].Items[itemId.ToString()].Fields
                    .PatchAsync(updateFields);

                System.Diagnostics.Debug.WriteLine($"Set status to '{newStatus}' for item {itemId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Graph status update error: {ex.Message}");
                throw;
            }
        }
    }
}
