using System;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Logging;

namespace SL2Tracer.Extensions.Logging
{
    /// <summary>
    /// Extension methods for registering SL2Tracer with Microsoft.Extensions.Logging.
    /// </summary>
    public static class TraceManagerLoggingExtensions
    {
        /// <summary>
        /// Adds SL2Tracer to the logging builder.
        /// </summary>
        public static ILoggingBuilder AddSL2Tracer(this ILoggingBuilder builder)
        {
            builder.Services.TryAddEnumerable(ServiceDescriptor.Singleton<ILoggerProvider, TraceManagerLoggerProvider>());
            return builder;
        }

        /// <summary>
        /// Adds SL2Tracer to the logging builder with custom configuration options.
        /// </summary>
        public static ILoggingBuilder AddSL2Tracer(this ILoggingBuilder builder, Action<TraceManagerOptions> configure)
        {
            if (configure != null)
            {
                builder.Services.Configure(configure);
            }
            return builder.AddSL2Tracer();
        }
    }
}
