using System;
using System.Collections.Concurrent;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Tracer.automation.level2;

namespace SL2Tracer.Extensions.Logging
{
    /// <summary>
    /// LoggerProvider for SL2Tracer.
    /// </summary>
    public class TraceManagerLoggerProvider : ILoggerProvider
    {
        private readonly TraceManager _tracer;
        private readonly ConcurrentDictionary<string, TraceManagerLogger> _loggers = new ConcurrentDictionary<string, TraceManagerLogger>(StringComparer.OrdinalIgnoreCase);
        private bool _disposed = false;

        public TraceManagerLoggerProvider(IOptions<TraceManagerOptions> options)
        {
            var opt = options?.Value ?? new TraceManagerOptions();
            _tracer = new TraceManager(opt.Application, opt.Path, opt.IsMainTracer, opt.Header, opt.MinimumLevel)
            {
                DictSingleTrace = opt.DictSingleTrace,
                MaxFiles = opt.MaxFiles,
                MaxFileSize = opt.MaxFileSizeKb,
                MaxDaysKeepingFiles = opt.MaxDaysKeepingFiles
            };
        }

        public TraceManagerLoggerProvider(TraceManager tracer)
        {
            _tracer = tracer ?? throw new ArgumentNullException(nameof(tracer));
        }

        public ILogger CreateLogger(string categoryName)
        {
            return _loggers.GetOrAdd(categoryName, name => new TraceManagerLogger(name, _tracer));
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _tracer.Dispose();
                _loggers.Clear();
                _disposed = true;
            }
        }
    }
}
