using System;
using Microsoft.Extensions.Logging;
using Tracer.automation.level2;

namespace SL2Tracer.Extensions.Logging
{
    /// <summary>
    /// Microsoft.Extensions.Logging.ILogger implementation backed by TraceManager.
    /// </summary>
    public class TraceManagerLogger : ILogger
    {
        private readonly string _categoryName;
        private readonly TraceManager _tracer;

        public TraceManagerLogger(string categoryName, TraceManager tracer)
        {
            _categoryName = categoryName ?? string.Empty;
            _tracer = tracer ?? throw new ArgumentNullException(nameof(tracer));
        }

        public IDisposable? BeginScope<TState>(TState state) where TState : notnull
        {
            return NullScope.Instance;
        }

        public bool IsEnabled(LogLevel logLevel)
        {
            var level = MapLogLevel(logLevel);
            return (int)level >= (int)_tracer.TraceLevel;
        }

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
        {
            if (!IsEnabled(logLevel))
            {
                return;
            }

            if (formatter == null)
            {
                throw new ArgumentNullException(nameof(formatter));
            }

            var message = formatter(state, exception);
            if (string.IsNullOrEmpty(message) && exception == null)
            {
                return;
            }

            var formatted = !string.IsNullOrEmpty(_categoryName)
                ? $"[{_categoryName}] {message}"
                : message;

            if (exception != null)
            {
                formatted = $"{formatted} {exception}";
            }

            var traceLevel = MapLogLevel(logLevel);
            switch (traceLevel)
            {
                case TraceManager.TraceLevels.CRITICAL:
                    _tracer.CRITICAL(formatted);
                    break;
                case TraceManager.TraceLevels.ERROR:
                    _tracer.ERROR(formatted);
                    break;
                case TraceManager.TraceLevels.EXCEPTION:
                    _tracer.EXCEPTION(formatted);
                    break;
                case TraceManager.TraceLevels.WARNING:
                    _tracer.WARNING(formatted);
                    break;
                case TraceManager.TraceLevels.INFO:
                    _tracer.INFO(formatted);
                    break;
                case TraceManager.TraceLevels.DEBUG:
                    _tracer.DEBUG(formatted);
                    break;
                case TraceManager.TraceLevels.VERBOSE:
                default:
                    _tracer.VERBOSE(formatted);
                    break;
            }
        }

        private static TraceManager.TraceLevels MapLogLevel(LogLevel logLevel)
        {
            switch (logLevel)
            {
                case LogLevel.Critical:
                    return TraceManager.TraceLevels.CRITICAL;
                case LogLevel.Error:
                    return TraceManager.TraceLevels.ERROR;
                case LogLevel.Warning:
                    return TraceManager.TraceLevels.WARNING;
                case LogLevel.Information:
                    return TraceManager.TraceLevels.INFO;
                case LogLevel.Debug:
                    return TraceManager.TraceLevels.DEBUG;
                case LogLevel.Trace:
                default:
                    return TraceManager.TraceLevels.VERBOSE;
            }
        }

        private sealed class NullScope : IDisposable
        {
            public static NullScope Instance { get; } = new NullScope();
            public void Dispose() { }
        }
    }
}
