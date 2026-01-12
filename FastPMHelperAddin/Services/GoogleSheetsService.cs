using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using FastPMHelperAddin.Models;
using ActionItemModel = FastPMHelperAddin.Models.ActionItem;

namespace FastPMHelperAddin.Services
{
    public class GoogleSheetsService
    {
        private readonly GoogleSheetsAuthService _authService;
        private readonly string _spreadsheetId;
        private readonly string _sheetName;
        private Dictionary<string, int> _headerIndexMap;

        // Column names (must match ActionItem properties)
        private readonly string[] _requiredColumns = new[]
        {
            "Id", "Project", "Package", "Title", "Status", "BallHolder", "SentOn", "DueDate", "HistoryLog", "LinkedThreadIDs", "ActiveMessageIDs"
        };

        public GoogleSheetsService(GoogleSheetsAuthService authService, string spreadsheetId, string sheetName)
        {
            _authService = authService;
            _spreadsheetId = spreadsheetId;
            _sheetName = sheetName;
        }

        private async Task EnsureHeaderRowAsync()
        {
            if (_headerIndexMap != null)
                return; // Already initialized

            var service = await _authService.GetSheetsServiceAsync();
            var range = $"{_sheetName}!A1:Z1";

            try
            {
                var request = service.Spreadsheets.Values.Get(_spreadsheetId, range);
                var response = await request.ExecuteAsync();

                if (response.Values == null || response.Values.Count == 0)
                {
                    // No header row - create it
                    System.Diagnostics.Debug.WriteLine("Creating header row in Google Sheets...");
                    await CreateHeaderRowAsync(service);
                }
                else
                {
                    // Parse existing header
                    var headerRow = response.Values[0];
                    _headerIndexMap = new Dictionary<string, int>();

                    for (int i = 0; i < headerRow.Count; i++)
                    {
                        _headerIndexMap[headerRow[i].ToString()] = i;
                    }

                    System.Diagnostics.Debug.WriteLine($"Header row found with {_headerIndexMap.Count} columns");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking header row: {ex.Message}");
                throw new Exception($"Failed to access spreadsheet: {ex.Message}", ex);
            }
        }

        private async Task CreateHeaderRowAsync(SheetsService service)
        {
            var range = $"{_sheetName}!A1:K1";
            var valueRange = new ValueRange
            {
                Values = new List<IList<object>> { _requiredColumns.Cast<object>().ToList() }
            };

            var request = service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

            await request.ExecuteAsync();

            // Rebuild header index map
            _headerIndexMap = new Dictionary<string, int>();
            for (int i = 0; i < _requiredColumns.Length; i++)
            {
                _headerIndexMap[_requiredColumns[i]] = i;
            }

            System.Diagnostics.Debug.WriteLine("Header row created successfully");
        }

        public async Task<List<ActionItemModel>> FetchOpenActionsAsync()
        {
            try
            {
                await EnsureHeaderRowAsync();
                var service = await _authService.GetSheetsServiceAsync();

                // Fetch all data rows (skip header)
                var range = $"{_sheetName}!A2:K";
                var request = service.Spreadsheets.Values.Get(_spreadsheetId, range);
                var response = await request.ExecuteAsync();

                var actionItems = new List<ActionItemModel>();

                if (response.Values != null)
                {
                    int rowNumber = 2; // Start at row 2 (after header)

                    foreach (var row in response.Values)
                    {
                        var status = GetCellValue(row, "Status");

                        // Only include non-closed actions
                        if (!status.Equals("Closed", StringComparison.OrdinalIgnoreCase))
                        {
                            actionItems.Add(new ActionItemModel
                            {
                                Id = rowNumber,
                                Project = GetCellValue(row, "Project"),
                                Package = GetCellValue(row, "Package"),
                                Title = GetCellValue(row, "Title"),
                                Status = status,
                                BallHolder = GetCellValue(row, "BallHolder"),
                                SentOn = ParseNullableDateTime(GetCellValue(row, "SentOn")),
                                DueDate = ParseNullableDateTime(GetCellValue(row, "DueDate")),
                                HistoryLog = GetCellValue(row, "HistoryLog"),
                                LinkedThreadIDs = GetCellValue(row, "LinkedThreadIDs"),
                                ActiveMessageIDs = GetCellValue(row, "ActiveMessageIDs")
                            });
                        }

                        rowNumber++;
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Fetched {actionItems.Count} open actions from Google Sheets");
                return actionItems;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Google Sheets fetch error: {ex.Message}");
                throw;
            }
        }

        public async Task<int> CreateActionAsync(string project, string package, string title, string ballHolder, string conversationId, string emailReference, string initialNote, DateTime sentOn, int dueDaysOffset)
        {
            try
            {
                await EnsureHeaderRowAsync();
                var service = await _authService.GetSheetsServiceAsync();

                // Get current row count to determine new row number
                var range = $"{_sheetName}!A:A";
                var request = service.Spreadsheets.Values.Get(_spreadsheetId, range);
                var response = await request.ExecuteAsync();

                int newRowNumber = response.Values != null ? response.Values.Count + 1 : 2;

                // Calculate due date
                var dueDate = sentOn.AddDays(dueDaysOffset);

                // Create new row
                var newRow = new List<object>
                {
                    newRowNumber.ToString(),                                          // Column A: Id
                    project,                                                          // Column B: Project
                    package,                                                          // Column C: Package
                    title,                                                            // Column D: Title
                    "Open",                                                           // Column E: Status
                    ballHolder,                                                       // Column F: BallHolder
                    sentOn.ToString("dd/MM/yyyy"),                                   // Column G: SentOn
                    dueDate.ToString("dd/MM/yyyy"),                                  // Column H: DueDate
                    $"[{DateTime.Now:dd/MM/yyyy}] Created: {initialNote}",          // Column I: HistoryLog
                    conversationId,                                                   // Column J: LinkedThreadIDs
                    emailReference                                                    // Column K: ActiveMessageIDs (NEW FORMAT: StoreID|EntryID|InternetMessageId)
                };

                var appendRange = $"{_sheetName}!A{newRowNumber}:K{newRowNumber}";
                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { newRow }
                };

                var appendRequest = service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, appendRange);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

                await appendRequest.ExecuteAsync();

                System.Diagnostics.Debug.WriteLine($"Created action with row number: {newRowNumber}");
                return newRowNumber;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Google Sheets create error: {ex.Message}");
                throw;
            }
        }

        public async Task UpdateActionAsync(int itemId, string newEmailReference, string ballHolder, string updateNote, DateTime sentOn, int dueDaysOffset)
        {
            try
            {
                await EnsureHeaderRowAsync();
                var service = await _authService.GetSheetsServiceAsync();

                // Fetch current row
                var rowRange = $"{_sheetName}!A{itemId}:K{itemId}";
                var getRequest = service.Spreadsheets.Values.Get(_spreadsheetId, rowRange);
                var response = await getRequest.ExecuteAsync();

                if (response.Values == null || response.Values.Count == 0)
                    throw new Exception($"Action with ID {itemId} not found");

                var currentRow = response.Values[0];

                // Append to ActiveMessageIDs
                string currentIds = GetCellValue(currentRow, "ActiveMessageIDs");
                var idList = ActionItemModel.ParseIds(currentIds);
                if (!idList.Contains(newEmailReference))
                    idList.Add(newEmailReference);

                // Append to HistoryLog
                string currentLog = GetCellValue(currentRow, "HistoryLog");
                string timestamp = DateTime.Now.ToString("dd/MM/yyyy");
                string updatedLog = $"{currentLog}\n\n[{timestamp}] {updateNote}";

                // Calculate new due date
                var dueDate = sentOn.AddDays(dueDaysOffset);

                // Update cells (BallHolder through ActiveMessageIDs)
                var updateRange = $"{_sheetName}!F{itemId}:K{itemId}";
                var updateRow = new List<object>
                {
                    ballHolder,                                          // Column F (BallHolder)
                    sentOn.ToString("dd/MM/yyyy"),                       // Column G (SentOn)
                    dueDate.ToString("dd/MM/yyyy"),                      // Column H (DueDate)
                    updatedLog,                                          // Column I (HistoryLog)
                    GetCellValue(currentRow, "LinkedThreadIDs"),         // Column J
                    string.Join("; ", idList)                            // Column K (ActiveMessageIDs)
                };

                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { updateRow }
                };

                var updateRequest = service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, updateRange);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

                await updateRequest.ExecuteAsync();

                System.Diagnostics.Debug.WriteLine($"Updated action {itemId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Google Sheets update error: {ex.Message}");
                throw;
            }
        }

        public async Task SetStatusAsync(int itemId, string newStatus, DateTime sentOn, int dueDaysOffset)
        {
            try
            {
                await EnsureHeaderRowAsync();
                var service = await _authService.GetSheetsServiceAsync();

                // Calculate new due date
                var dueDate = sentOn.AddDays(dueDaysOffset);

                // Update status, SentOn, and DueDate (Columns E, G, H)
                var updates = new List<ValueRange>
                {
                    new ValueRange
                    {
                        Range = $"{_sheetName}!E{itemId}",  // Status is now column E
                        Values = new List<IList<object>> { new List<object> { newStatus } }
                    },
                    new ValueRange
                    {
                        Range = $"{_sheetName}!G{itemId}",  // SentOn is now column G
                        Values = new List<IList<object>> { new List<object> { sentOn.ToString("dd/MM/yyyy") } }
                    },
                    new ValueRange
                    {
                        Range = $"{_sheetName}!H{itemId}",  // DueDate is now column H
                        Values = new List<IList<object>> { new List<object> { dueDate.ToString("dd/MM/yyyy") } }
                    }
                };

                var batchUpdateRequest = new BatchUpdateValuesRequest
                {
                    ValueInputOption = "RAW",
                    Data = updates
                };

                await service.Spreadsheets.Values.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync();

                System.Diagnostics.Debug.WriteLine($"Set status to '{newStatus}', SentOn to '{sentOn}', DueDate to '{dueDate}' for row {itemId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Google Sheets status update error: {ex.Message}");
                throw;
            }
        }

        public async Task CloseActionAsync(int itemId, string closureNote, string closingEmailReference, DateTime sentOn)
        {
            try
            {
                await EnsureHeaderRowAsync();
                var service = await _authService.GetSheetsServiceAsync();

                // Fetch current row to get HistoryLog and ActiveMessageIDs
                var rowRange = $"{_sheetName}!A{itemId}:K{itemId}";
                var getRequest = service.Spreadsheets.Values.Get(_spreadsheetId, rowRange);
                var response = await getRequest.ExecuteAsync();

                if (response.Values == null || response.Values.Count == 0)
                    throw new Exception($"Action with ID {itemId} not found");

                var currentRow = response.Values[0];

                // Append to HistoryLog with closure note
                string currentLog = GetCellValue(currentRow, "HistoryLog");
                string timestamp = DateTime.Now.ToString("dd/MM/yyyy");
                string updatedLog = $"{currentLog}\n\n[{timestamp}] Closed: {closureNote}";

                // Append to ActiveMessageIDs
                string currentIds = GetCellValue(currentRow, "ActiveMessageIDs");
                var idList = ActionItemModel.ParseIds(currentIds);
                if (!idList.Contains(closingEmailReference))
                    idList.Add(closingEmailReference);

                // Update: Status=Closed, BallHolder=empty, SentOn=closing email date, DueDate=empty, HistoryLog, ActiveMessageIDs
                var updates = new List<ValueRange>
                {
                    new ValueRange
                    {
                        Range = $"{_sheetName}!E{itemId}",  // Status
                        Values = new List<IList<object>> { new List<object> { "Closed" } }
                    },
                    new ValueRange
                    {
                        Range = $"{_sheetName}!F{itemId}",  // BallHolder (clear it)
                        Values = new List<IList<object>> { new List<object> { "" } }
                    },
                    new ValueRange
                    {
                        Range = $"{_sheetName}!G{itemId}",  // SentOn
                        Values = new List<IList<object>> { new List<object> { sentOn.ToString("dd/MM/yyyy") } }
                    },
                    new ValueRange
                    {
                        Range = $"{_sheetName}!H{itemId}",  // DueDate (clear it)
                        Values = new List<IList<object>> { new List<object> { "" } }
                    },
                    new ValueRange
                    {
                        Range = $"{_sheetName}!I{itemId}",  // HistoryLog
                        Values = new List<IList<object>> { new List<object> { updatedLog } }
                    },
                    new ValueRange
                    {
                        Range = $"{_sheetName}!K{itemId}",  // ActiveMessageIDs
                        Values = new List<IList<object>> { new List<object> { string.Join("; ", idList) } }
                    }
                };

                var batchUpdateRequest = new BatchUpdateValuesRequest
                {
                    ValueInputOption = "RAW",
                    Data = updates
                };

                await service.Spreadsheets.Values.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync();

                System.Diagnostics.Debug.WriteLine($"Closed action {itemId} with note: {closureNote}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Google Sheets close action error: {ex.Message}");
                throw;
            }
        }

        private string GetCellValue(IList<object> row, string columnName)
        {
            if (_headerIndexMap.TryGetValue(columnName, out int index) && index < row.Count)
            {
                return row[index]?.ToString() ?? string.Empty;
            }
            return string.Empty;
        }

        private DateTime? ParseNullableDateTime(string dateString)
        {
            if (string.IsNullOrWhiteSpace(dateString))
                return null;

            if (DateTime.TryParse(dateString, out DateTime result))
                return result;

            return null;
        }

        /// <summary>
        /// Fetch classification rules from Config_Rules tab
        /// </summary>
        public async Task<IList<IList<object>>> FetchConfigRulesAsync()
        {
            try
            {
                var service = await _authService.GetSheetsServiceAsync();

                // Fetch Config_Rules tab (columns A-F, all rows)
                var range = "Config_Rules!A:F";
                var request = service.Spreadsheets.Values.Get(_spreadsheetId, range);
                var response = await request.ExecuteAsync();

                System.Diagnostics.Debug.WriteLine($"Fetched {response.Values?.Count ?? 0} rows from Config_Rules");
                return response.Values;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error fetching Config_Rules: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Updates editable fields of an action (Project, Package, Title, BallHolder, DueDate)
        /// Also appends a history log entry
        /// </summary>
        public async Task UpdateActionFieldsAsync(
            int itemId,
            string project,
            string package,
            string title,
            string ballHolder,
            DateTime? dueDate)
        {
            try
            {
                await EnsureHeaderRowAsync();
                var service = await _authService.GetSheetsServiceAsync();

                // Fetch current row to get Status, SentOn, HistoryLog
                var rowRange = $"{_sheetName}!A{itemId}:K{itemId}";
                var getRequest = service.Spreadsheets.Values.Get(_spreadsheetId, rowRange);
                var response = await getRequest.ExecuteAsync();

                if (response.Values == null || response.Values.Count == 0)
                    throw new Exception($"Action with ID {itemId} not found");

                var currentRow = response.Values[0];

                // Append to HistoryLog
                string currentLog = GetCellValue(currentRow, "HistoryLog");
                string timestamp = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                string updatedLog = $"{currentLog}\n\n[{timestamp}] Manual edit: Updated fields";

                // Format due date
                string dueDateStr = dueDate.HasValue ? dueDate.Value.ToString("dd/MM/yyyy") : "";

                // Update columns B through I (Project, Package, Title, Status, BallHolder, SentOn, DueDate, HistoryLog)
                var updateRange = $"{_sheetName}!B{itemId}:I{itemId}";

                var updateRow = new List<object>
                {
                    project,                                         // Column B (Project)
                    package,                                         // Column C (Package)
                    title,                                           // Column D (Title)
                    GetCellValue(currentRow, "Status"),             // Column E (Status) - unchanged
                    ballHolder,                                      // Column F (BallHolder)
                    GetCellValue(currentRow, "SentOn"),             // Column G (SentOn) - unchanged
                    dueDateStr,                                      // Column H (DueDate)
                    updatedLog                                       // Column I (HistoryLog)
                };

                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { updateRow }
                };

                var updateRequest = service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, updateRange);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

                await updateRequest.ExecuteAsync();

                System.Diagnostics.Debug.WriteLine($"Updated action fields for row {itemId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Google Sheets field update error: {ex.Message}");
                throw;
            }
        }
    }
}
