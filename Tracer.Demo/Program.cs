using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Tracer.automation.level2;

namespace Tracer.Demo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("==================================================");
            Console.WriteLine("       SL2Tracer Cross-Platform Console Demo      ");
            Console.WriteLine("==================================================");
            Console.ResetColor();

            string logDir = Path.Combine(Directory.GetCurrentDirectory(), "demo_logs");
            Console.WriteLine($"[1] Initializing TraceManager in directory: {logDir}");

            using (var logger = new TraceManager("DemoApp", logDir, true, TraceManager.TraceLevels.VERBOSE))
            {
                logger.Header = "DemoModule";

                Console.WriteLine("\n[2] Testing standard log levels:");
                logger.CRITICAL("Critical system event occurred");
                logger.ERROR("An unexpected error occurred during processing");
                logger.EXCEPTION("Sample handled exception message");
                logger.WARNING("Low disk space warning threshold triggered");
                logger.INFO("Application initialized successfully");
                logger.DEBUG("Cache lookup completed in 1.4ms");
                logger.VERBOSE("Entering dispatch loop step 0");
                Console.WriteLine("    -> Written CRITICAL, ERROR, EXCEPTION, WARNING, INFO, DEBUG, VERBOSE");

                Console.WriteLine("\n[3] Testing Collections & Dictionary tracing:");
                var nestedDict = new Dictionary<string, object>
                {
                    { "Server", "prod-worker-01" },
                    { "Port", 8080 },
                    { "Metrics", new Dictionary<string, int> { { "ActiveConnections", 42 }, { "QueueDepth", 3 } } }
                };
                logger.INFO(nestedDict);
                Console.WriteLine("    -> Multi-line nested dictionary logged");

                Console.WriteLine("\n[4] Testing Single-Line Dictionary Tracing:");
                logger.DictSingleTrace = true;
                var singleLineDict = new Dictionary<string, string>
                {
                    { "UserId", "USR-9941" },
                    { "Role", "Administrator" },
                    { "IP", "192.168.1.100" }
                };
                logger.INFO(singleLineDict);
                Console.WriteLine("    -> Single-line dictionary logged");

                Console.WriteLine("\n[5] Triggering Purge and Retention check:");
                logger.TriggerPurgeCheck();
                Console.WriteLine("    -> Purger check executed cleanly");
            }

            string logFile = Path.Combine(logDir, "DemoApp.Tracer.txt");
            if (File.Exists(logFile))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\n[SUCCESS] Log file created: {logFile}");
                Console.ResetColor();
                Console.WriteLine("\n--- Tail of Log File ---");
                var lines = File.ReadAllLines(logFile);
                int start = Math.Max(0, lines.Length - 10);
                for (int i = start; i < lines.Length; i++)
                {
                    Console.WriteLine(lines[i]);
                }
                Console.WriteLine("------------------------\n");
            }

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Demo completed successfully!");
            Console.ResetColor();
        }
    }
}
