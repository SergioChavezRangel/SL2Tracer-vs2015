using System;
using System.IO;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using SL2Tracer.Extensions.Logging;
using Tracer.automation.level2;
using Xunit;

namespace Tracer.Tests
{
    public class SL2TracerExtensionsTests : IDisposable
    {
        private readonly string _testDir;
        private readonly string _appName;

        public SL2TracerExtensionsTests()
        {
            _appName = "ExtApp_" + Guid.NewGuid().ToString("N").Substring(0, 8);
            _testDir = Path.Combine(Path.GetTempPath(), "TracerExtTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_testDir);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_testDir))
                {
                    Directory.Delete(_testDir, true);
                }
            }
            catch
            {
                // Best effort
            }
        }

        [Fact]
        public void MicrosoftExtensionsLogging_IntegratesViaDependencyInjection()
        {
            var services = new ServiceCollection();
            services.AddLogging(builder =>
            {
                builder.AddSL2Tracer(options =>
                {
                    options.Application = _appName;
                    options.Path = _testDir;
                    options.IsMainTracer = false;
                    options.MinimumLevel = TraceManager.TraceLevels.INFO;
                });
            });

            using (var serviceProvider = services.BuildServiceProvider())
            {
                var logger = serviceProvider.GetRequiredService<ILogger<SL2TracerExtensionsTests>>();

                logger.LogTrace("This trace message should be filtered out");
                logger.LogDebug("This debug message should be filtered out");
                logger.LogInformation("Service started successfully for item {ItemId}", 101);
                logger.LogWarning("Potential degradation detected");
                logger.LogError(new InvalidOperationException("DB timeout"), "Failed to process transaction {TxId}", 9999);
            }

            var logFile = Path.Combine(_testDir, $"{_appName}.Tracer.txt");
            Assert.True(File.Exists(logFile), $"Log file {logFile} should exist");
            var content = File.ReadAllText(logFile);

            Assert.DoesNotContain("This trace message should be filtered out", content);
            Assert.DoesNotContain("This debug message should be filtered out", content);
            Assert.Contains("Service started successfully for item 101", content);
            Assert.Contains("Potential degradation detected", content);
            Assert.Contains("Failed to process transaction 9999", content);
            Assert.Contains("InvalidOperationException", content);
            Assert.Contains("[Tracer.Tests.SL2TracerExtensionsTests]", content);
        }
    }
}
