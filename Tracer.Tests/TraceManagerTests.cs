using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Tracer.automation.level2;
using Xunit;

namespace Tracer.Tests
{
    public class TraceManagerTests : IDisposable
    {
        private readonly string _testDir;
        private readonly string _appName;

        public TraceManagerTests()
        {
            _appName = "TestApp_" + Guid.NewGuid().ToString("N").Substring(0, 8);
            _testDir = Path.Combine(Path.GetTempPath(), "TracerTests_" + Guid.NewGuid().ToString("N"));
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
                // Cleanup best effort
            }
        }

        private string GetActiveLogPath()
        {
            return Path.Combine(_testDir, $"{_appName}.Tracer.txt");
        }

        [Fact]
        public void LogLevelFiltering_FiltersMessagesBelowThreshold()
        {
            using (var tracer = new TraceManager(_appName, _testDir, false, TraceManager.TraceLevels.INFO))
            {
                tracer.VERBOSE("verbose_msg");
                tracer.DEBUG("debug_msg");
                tracer.INFO("info_msg");
                tracer.WARNING("warning_msg");
                tracer.ERROR("error_msg");
                tracer.EXCEPTION("exception_msg");
                tracer.CRITICAL("critical_msg");
            }

            var logFile = GetActiveLogPath();
            Assert.True(File.Exists(logFile));
            var content = File.ReadAllText(logFile);

            Assert.DoesNotContain("verbose_msg", content);
            Assert.DoesNotContain("debug_msg", content);
            Assert.Contains("info_msg", content);
            Assert.Contains("warning_msg", content);
            Assert.Contains("error_msg", content);
            Assert.Contains("exception_msg", content);
            Assert.Contains("critical_msg", content);
            Assert.Contains("[INFO]", content);
            Assert.Contains("[ERROR]", content);
            Assert.Contains("[CRITICAL]", content);
        }

        [Fact]
        public void DictionaryLogging_SupportsGenericAndNestedDictionariesWithoutCrash()
        {
            using (var tracer = new TraceManager(_appName, _testDir, false, TraceManager.TraceLevels.DEBUG))
            {
                // Standard Hashtable
                var ht = new Hashtable
                {
                    { "Key1", "Value1" },
                    { "Key2", 42 }
                };
                tracer.INFO(ht);

                // Generic Dictionary<string, string>
                var genericDict = new Dictionary<string, string>
                {
                    { "GenKeyA", "GenValA" },
                    { "GenKeyB", "GenValB" }
                };
                tracer.INFO(genericDict);

                // Nested Dictionary
                var nestedDict = new Dictionary<string, object>
                {
                    { "ParentKey", "ParentValue" },
                    { "ChildDict", new Dictionary<string, int> { { "ChildCount", 10 } } }
                };
                tracer.INFO(nestedDict);
            }

            var content = File.ReadAllText(GetActiveLogPath());
            Assert.Contains("Key1", content);
            Assert.Contains("Value1", content);
            Assert.Contains("GenKeyA", content);
            Assert.Contains("GenValA", content);
            Assert.Contains("ParentKey", content);
            Assert.Contains("ChildCount", content);
            Assert.Contains("10", content);
        }

        [Fact]
        public void DictSingleTrace_FormatsOnSingleLine()
        {
            using (var tracer = new TraceManager(_appName, _testDir, false, TraceManager.TraceLevels.DEBUG))
            {
                tracer.DictSingleTrace = true;
                var dict = new Dictionary<string, string>
                {
                    { "Alpha", "One" },
                    { "Beta", "Two" }
                };
                tracer.INFO(dict);
            }

            var lines = File.ReadAllLines(GetActiveLogPath());
            var dictLine = lines.FirstOrDefault(l => l.Contains("Alpha=One"));
            Assert.NotNull(dictLine);
            Assert.Contains("Beta=Two", dictLine);
        }

        [Fact]
        public void ThreadSafety_ConcurrentWrites_DoesNotThrowAndPreservesAllEntries()
        {
            const int threadCount = 10;
            const int messagesPerThread = 50;

            using (var tracer = new TraceManager(_appName, _testDir, false, TraceManager.TraceLevels.VERBOSE))
            {
                Parallel.For(0, threadCount, t =>
                {
                    for (int i = 0; i < messagesPerThread; i++)
                    {
                        tracer.INFO($"thread_{t}_msg_{i}");
                    }
                });
            }

            var content = File.ReadAllText(GetActiveLogPath());
            for (int t = 0; t < threadCount; t++)
            {
                for (int i = 0; i < messagesPerThread; i++)
                {
                    Assert.Contains($"thread_{t}_msg_{i}", content);
                }
            }
        }

        [Fact]
        public void PurgeMaxFiles_EnforcesExactFileLimit_FixesOffByOne()
        {
            const short maxFiles = 3;

            // Create 6 historical rolled files
            for (int i = 1; i <= 6; i++)
            {
                var filePath = Path.Combine(_testDir, $"{_appName}.Tracer_2026010{i}_120000.txt");
                File.WriteAllText(filePath, $"historical log {i}");
                File.SetLastWriteTime(filePath, DateTime.Now.AddHours(-10 + i));
            }

            using (var tracer = new TraceManager(_appName, _testDir, false, TraceManager.TraceLevels.DEBUG))
            {
                tracer.MaxFiles = maxFiles;
                tracer.INFO("Active log trigger");
                tracer.TriggerPurgeCheck();
            }

            var historicalFiles = Directory.GetFiles(_testDir, $"{_appName}.Tracer_*.txt");
            Assert.Equal(maxFiles, historicalFiles.Length);
        }

        [Fact]
        public void CleanFiles_RemovesFilesOlderThanKeepDays()
        {
            const short keepDays = 5;

            var oldFilePath = Path.Combine(_testDir, $"{_appName}.Tracer_20250101_120000.txt");
            File.WriteAllText(oldFilePath, "very old log");
            File.SetLastWriteTime(oldFilePath, DateTime.Now.AddDays(-20));

            var freshFilePath = Path.Combine(_testDir, $"{_appName}.Tracer_20261020_120000.txt");
            File.WriteAllText(freshFilePath, "fresh log");
            File.SetLastWriteTime(freshFilePath, DateTime.Now.AddDays(-1));

            using (var tracer = new TraceManager(_appName, _testDir, false, TraceManager.TraceLevels.DEBUG))
            {
                tracer.MaxDaysKeepingFiles = keepDays;
                tracer.TriggerPurgeCheck();
            }

            Assert.False(File.Exists(oldFilePath), "Old file should have been deleted");
            Assert.True(File.Exists(freshFilePath), "Fresh file should be retained");
        }

        [Fact]
        public void HeaderProperty_IncludedInFormattedLogEntry()
        {
            using (var tracer = new TraceManager(_appName, _testDir, false, "SubModule", TraceManager.TraceLevels.INFO))
            {
                tracer.INFO("message with header");
            }

            var content = File.ReadAllText(GetActiveLogPath());
            Assert.Contains($"[{_appName}].[SubModule]", content);
            Assert.Contains("message with header", content);
        }

        [Fact]
        public void ExceptionOverload_LogsExceptionDetails()
        {
            using (var tracer = new TraceManager(_appName, _testDir, false, TraceManager.TraceLevels.DEBUG))
            {
                try
                {
                    throw new InvalidOperationException("Simulated failure for tracer test");
                }
                catch (Exception ex)
                {
                    tracer.EXCEPTION(ex);
                }
            }

            var content = File.ReadAllText(GetActiveLogPath());
            Assert.Contains("Simulated failure for tracer test", content);
            Assert.Contains("InvalidOperationException", content);
            Assert.Contains("[EXCEPTION]", content);
        }
    }
}
