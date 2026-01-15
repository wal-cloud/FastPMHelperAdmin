using System;

namespace FastPMHelperAddin.Models
{
    /// <summary>
    /// Data model for storing deferred action execution instructions in draft emails.
    /// Serialized to JSON and stored in UserProperties["Deferred_Execution_Data"].
    /// </summary>
    public class DeferredActionData
    {
        /// <summary>
        /// The operation mode: "Create" or "Update"
        /// </summary>
        public string Mode { get; set; }

        /// <summary>
        /// The ID of the action to update (only used when Mode = "Update")
        /// </summary>
        public int? ActionID { get; set; }

        /// <summary>
        /// Optional manual title override (reserved for future use)
        /// </summary>
        public string ManualTitle { get; set; }

        /// <summary>
        /// Optional manual assignee override (reserved for future use)
        /// </summary>
        public string ManualAssignee { get; set; }
    }
}
