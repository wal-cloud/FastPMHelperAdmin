using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastPMHelperAddin.Services
{
    /// <summary>
    /// Service for managing Outlook categories on emails to track processing status
    /// </summary>
    public class EmailCategoryService
    {
        private readonly Outlook.Application _app;
        private readonly Outlook.NameSpace _ns;

        private const string TRACKED_CATEGORY = "Tracked";
        private const string LLM_ERROR_CATEGORY = "LLM Error";
        private const string SHEETS_ERROR_CATEGORY = "Sheets Error";

        public EmailCategoryService()
        {
            _app = Globals.ThisAddIn.Application;
            _ns = _app.GetNamespace("MAPI");
        }

        /// <summary>
        /// Marks an email with the "Tracked" category to indicate successful processing.
        /// Also removes "LLM Error" category if present (success clears previous errors).
        /// </summary>
        /// <param name="mail">The MailItem to tag</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool MarkEmailAsTracked(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                System.Diagnostics.Debug.WriteLine("MarkEmailAsTracked: mail is null, skipping");
                return false;
            }

            try
            {
                // Ensure "Tracked" category exists in Master Category List
                EnsureTrackedCategoryExists();

                // Remove error categories if present (success clears previous errors)
                RemoveCategoryFromMail(mail, LLM_ERROR_CATEGORY);
                RemoveCategoryFromMail(mail, SHEETS_ERROR_CATEGORY);

                // Get current categories
                string currentCategories = mail.Categories ?? string.Empty;
                var categoryList = currentCategories
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(c => c.Trim())
                    .ToList();

                // Check if already tracked
                if (categoryList.Any(c => c.Equals(TRACKED_CATEGORY, StringComparison.OrdinalIgnoreCase)))
                {
                    System.Diagnostics.Debug.WriteLine($"Email already has '{TRACKED_CATEGORY}' category: {mail.Subject}");
                    return true; // Already tracked, consider this success
                }

                // Append "Tracked" category
                categoryList.Add(TRACKED_CATEGORY);
                mail.Categories = string.Join(", ", categoryList);

                // Save the mail item
                mail.Save();

                System.Diagnostics.Debug.WriteLine($"✓ Email marked as Tracked: {mail.Subject}");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR marking email as tracked: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Marks an email with the "LLM Error" category to indicate LLM processing failure.
        /// Does NOT remove "Tracked" category (preserves previous success).
        /// </summary>
        /// <param name="mail">The MailItem to tag</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool MarkEmailAsLLMError(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                System.Diagnostics.Debug.WriteLine("MarkEmailAsLLMError: mail is null, skipping");
                return false;
            }

            try
            {
                // Ensure "LLM Error" category exists in Master Category List
                EnsureLLMErrorCategoryExists();

                // Get current categories
                string currentCategories = mail.Categories ?? string.Empty;
                var categoryList = currentCategories
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(c => c.Trim())
                    .ToList();

                // Check if already has LLM Error
                if (categoryList.Any(c => c.Equals(LLM_ERROR_CATEGORY, StringComparison.OrdinalIgnoreCase)))
                {
                    System.Diagnostics.Debug.WriteLine($"Email already has '{LLM_ERROR_CATEGORY}' category: {mail.Subject}");
                    return true; // Already marked, consider this success
                }

                // Append "LLM Error" category
                categoryList.Add(LLM_ERROR_CATEGORY);
                mail.Categories = string.Join(", ", categoryList);

                // Save the mail item
                mail.Save();

                System.Diagnostics.Debug.WriteLine($"⚠ Email marked as LLM Error: {mail.Subject}");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR marking email as LLM error: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Marks an email with the "Sheets Error" category to indicate Google Sheets write failure.
        /// This is a CRITICAL error - the action was not recorded at all.
        /// Does NOT remove "Tracked" or "LLM Error" categories (preserves full history).
        /// </summary>
        /// <param name="mail">The MailItem to tag</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool MarkEmailAsSheetsError(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                System.Diagnostics.Debug.WriteLine("MarkEmailAsSheetsError: mail is null, skipping");
                return false;
            }

            try
            {
                // Ensure "Sheets Error" category exists in Master Category List
                EnsureSheetsErrorCategoryExists();

                // Get current categories
                string currentCategories = mail.Categories ?? string.Empty;
                var categoryList = currentCategories
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(c => c.Trim())
                    .ToList();

                // Check if already has Sheets Error
                if (categoryList.Any(c => c.Equals(SHEETS_ERROR_CATEGORY, StringComparison.OrdinalIgnoreCase)))
                {
                    System.Diagnostics.Debug.WriteLine($"Email already has '{SHEETS_ERROR_CATEGORY}' category: {mail.Subject}");
                    return true; // Already marked, consider this success
                }

                // Append "Sheets Error" category
                categoryList.Add(SHEETS_ERROR_CATEGORY);
                mail.Categories = string.Join(", ", categoryList);

                // Save the mail item
                mail.Save();

                System.Diagnostics.Debug.WriteLine($"🔴 Email marked as Sheets Error: {mail.Subject}");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR marking email as Sheets error: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Removes the "Tracked" category from an email (used when toggle is turned off).
        /// </summary>
        /// <param name="mail">The MailItem to update</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool RemoveTrackedCategory(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                System.Diagnostics.Debug.WriteLine("RemoveTrackedCategory: mail is null, skipping");
                return false;
            }

            try
            {
                bool removed = RemoveCategoryFromMail(mail, TRACKED_CATEGORY);

                if (removed)
                {
                    System.Diagnostics.Debug.WriteLine($"✓ Removed Tracked category: {mail.Subject}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"No Tracked category to remove: {mail.Subject}");
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR removing tracked category: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Removes a specific category from a mail item
        /// </summary>
        /// <param name="mail">The MailItem to update</param>
        /// <param name="categoryName">The category name to remove</param>
        /// <returns>True if category was removed, false if it wasn't present</returns>
        private bool RemoveCategoryFromMail(Outlook.MailItem mail, string categoryName)
        {
            if (mail == null || string.IsNullOrEmpty(categoryName))
                return false;

            try
            {
                // Get current categories
                string currentCategories = mail.Categories ?? string.Empty;
                var categoryList = currentCategories
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(c => c.Trim())
                    .ToList();

                // Check if category exists
                int initialCount = categoryList.Count;
                categoryList.RemoveAll(c => c.Equals(categoryName, StringComparison.OrdinalIgnoreCase));

                // If count changed, category was removed
                if (categoryList.Count < initialCount)
                {
                    // Update categories
                    mail.Categories = string.Join(", ", categoryList);
                    mail.Save();
                    return true;
                }

                return false; // Category wasn't present
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR removing category '{categoryName}': {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Ensures the "Tracked" category exists in the Master Category List
        /// </summary>
        private void EnsureTrackedCategoryExists()
        {
            try
            {
                // Check if "Tracked" category exists
                bool categoryExists = false;
                foreach (Outlook.Category cat in _ns.Categories)
                {
                    if (cat.Name.Equals(TRACKED_CATEGORY, StringComparison.OrdinalIgnoreCase))
                    {
                        categoryExists = true;
                        break;
                    }
                }

                // Create if doesn't exist
                if (!categoryExists)
                {
                    _ns.Categories.Add(TRACKED_CATEGORY, Outlook.OlCategoryColor.olCategoryColorDarkGreen);
                    System.Diagnostics.Debug.WriteLine($"✓ Created '{TRACKED_CATEGORY}' category in Master Category List");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WARNING: Could not ensure Tracked category exists: {ex.Message}");
                // Don't throw - category tagging will fail gracefully if category doesn't exist
            }
        }

        /// <summary>
        /// Ensures the "LLM Error" category exists in the Master Category List
        /// </summary>
        private void EnsureLLMErrorCategoryExists()
        {
            try
            {
                // Check if "LLM Error" category exists
                bool categoryExists = false;
                foreach (Outlook.Category cat in _ns.Categories)
                {
                    if (cat.Name.Equals(LLM_ERROR_CATEGORY, StringComparison.OrdinalIgnoreCase))
                    {
                        categoryExists = true;
                        break;
                    }
                }

                // Create if doesn't exist
                if (!categoryExists)
                {
                    _ns.Categories.Add(LLM_ERROR_CATEGORY, Outlook.OlCategoryColor.olCategoryColorYellow);
                    System.Diagnostics.Debug.WriteLine($"✓ Created '{LLM_ERROR_CATEGORY}' category in Master Category List");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WARNING: Could not ensure LLM Error category exists: {ex.Message}");
                // Don't throw - category tagging will fail gracefully if category doesn't exist
            }
        }

        /// <summary>
        /// Ensures the "Sheets Error" category exists in the Master Category List
        /// </summary>
        private void EnsureSheetsErrorCategoryExists()
        {
            try
            {
                // Check if "Sheets Error" category exists
                bool categoryExists = false;
                foreach (Outlook.Category cat in _ns.Categories)
                {
                    if (cat.Name.Equals(SHEETS_ERROR_CATEGORY, StringComparison.OrdinalIgnoreCase))
                    {
                        categoryExists = true;
                        break;
                    }
                }

                // Create if doesn't exist
                if (!categoryExists)
                {
                    _ns.Categories.Add(SHEETS_ERROR_CATEGORY, Outlook.OlCategoryColor.olCategoryColorRed);
                    System.Diagnostics.Debug.WriteLine($"✓ Created '{SHEETS_ERROR_CATEGORY}' category in Master Category List");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WARNING: Could not ensure Sheets Error category exists: {ex.Message}");
                // Don't throw - category tagging will fail gracefully if category doesn't exist
            }
        }
    }
}
