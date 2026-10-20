using System;
using Tracer.automation.level2;

namespace SL2Tracer.Extensions.Logging
{
    /// <summary>
    /// Configuration options for SL2Tracer logging provider.
    /// </summary>
    public class TraceManagerOptions
    {
        /// <summary>
        /// Application name prefix used for log file naming.
        /// </summary>
        public string Application { get; set; } = AppDomain.CurrentDomain.FriendlyName;

        /// <summary>
        /// Directory where log files are stored.
        /// </summary>
        public string Path { get; set; } = AppDomain.CurrentDomain.BaseDirectory;

        /// <summary>
        /// Header / submodule identifier added to log entries.
        /// </summary>
        public string Header { get; set; } = string.Empty;

        /// <summary>
        /// Whether this instance is the main tracer managing automatic background purger/cleaner.
        /// </summary>
        public bool IsMainTracer { get; set; } = true;

        /// <summary>
        /// Minimum trace level threshold to record.
        /// </summary>
        public TraceManager.TraceLevels MinimumLevel { get; set; } = TraceManager.TraceLevels.NOTSET;

        /// <summary>
        /// Formats dictionaries on a single delimited line instead of indented multi-line.
        /// </summary>
        public bool DictSingleTrace { get; set; } = false;

        /// <summary>
        /// Maximum number of historical rolled log files to retain.
        /// </summary>
        public short MaxFiles { get; set; } = 5;

        /// <summary>
        /// Maximum file size in Kilobytes before rolling.
        /// </summary>
        public long MaxFileSizeKb { get; set; } = 1024 * 1024;

        /// <summary>
        /// Maximum days to retain rolled log files.
        /// </summary>
        public short MaxDaysKeepingFiles { get; set; } = 7;
    }
}
