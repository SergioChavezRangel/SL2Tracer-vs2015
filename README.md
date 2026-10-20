# SL2Tracer

[![.NET 8.0](https://img.shields.io/badge/.NET-8.0-512BD4?logo=dotnet&logoColor=white)](https://dotnet.microsoft.com/)
[![.NET Standard 2.0](https://img.shields.io/badge/.NET%20Standard-2.0-blue.svg)](https://docs.microsoft.com/en-us/dotnet/standard/net-standard)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)](https://dotnet.microsoft.com/)
[![Build and Test](https://github.com/SergioChavezRangel/SL2Tracer-vs2015/actions/workflows/build-and-test.yml/badge.svg)](https://github.com/SergioChavezRangel/SL2Tracer-vs2015/actions/workflows/build-and-test.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A robust, cross-platform file-based logging and tracing library designed for .NET applications. It features automatic file size rolling, daily scheduled rotation, retention purging, hierarchical collection tracing (nested dictionaries and hashtables), and native `Microsoft.Extensions.Logging` integration for modern Dependency Injection.

---

## Architecture & Solution Structure

```
SL2Tracer-vs2015/
├── Tracer/                         # Core library (netstandard2.0 & net8.0)
│   ├── TraceManager.vb             # Main TraceManager logger engine
│   └── PurgeConfig.xml             # XML configuration for file rolling and purger
├── SL2Tracer.Extensions.Logging/   # Microsoft.Extensions.Logging adapter (netstandard2.0 & net8.0)
│   ├── TraceManagerLogger.cs       # ILogger implementation
│   ├── TraceManagerLoggerProvider.cs # ILoggerProvider
│   └── TraceManagerLoggingExtensions.cs # builder.AddSL2Tracer()
├── Tracer.Tests/                   # Automated cross-platform unit & integration test suite (xUnit)
│   ├── TraceManagerTests.cs        # Tests for thread safety, levels, rolling & purging
│   └── SL2TracerExtensionsTests.cs # Tests for DI and Microsoft.Extensions.Logging
├── Tracer.Demo/                    # Cross-platform interactive console demo (net8.0)
│   └── Program.cs
└── TracerTest/                     # Legacy Windows Forms test application (net8.0-windows)
    └── TestApp.cs
```

---

## Key Features

- **Cross-Platform**: Multi-targets `.NET Standard 2.0` and `.NET 8.0`. Runs natively on **Windows, macOS, and Linux**.
- **Thread-Safe**: Fully synchronized file I/O prevents concurrent access collisions (`IOException`) in multi-threaded workloads.
- **Trace Levels**:
  - `CRITICAL (50)`
  - `ERROR (45)`
  - `EXCEPTION (40)`
  - `WARNING (30)`
  - `INFO (20)`
  - `NOTSET (15)` *(default threshold)*
  - `DEBUG (10)`
  - `VERBOSE (0)`
- **Hierarchical Collection Tracing**:
  - Traces `Hashtable`, generic `Dictionary<TKey, TValue>`, and arbitrary nested key-value collections.
  - Multi-line indented format or single-line delimited format (`DictSingleTrace = true`).
- **Automatic File Rolling & Purging**:
  - **Size Rolling**: Automatically archives log file when size exceeds `maxFileSize` (e.g. 10 MB).
  - **Daily Rotation**: Automatically rotates files at a specified minute after midnight (`time4Clean`, e.g. 07:00 AM).
  - **History Purging**: Automatically deletes old rolled logs when total files exceed `maxFiles` or age exceeds `keepLogs` days.
- **Modern DI Integration**: Compatible with `Microsoft.Extensions.Logging` via `builder.AddSL2Tracer()`.
- **Resource Management**: Fully implements `IDisposable` for deterministic timer cleanup and resource teardown.

---

## Quickstart

### 1. Direct API Usage (C#)

```csharp
using Tracer.automation.level2;

// Initialize tracer (ApplicationName, LogDirectory, IsMainTracer, LogLevel)
using var logger = new TraceManager("MyApp", "logs", isMainTracer: true, TraceManager.TraceLevels.INFO);

logger.Header = "PaymentModule";

// Standard logging
logger.INFO("Payment gateway connected");
logger.WARNING("Retry attempt 1 of 3");
logger.ERROR("Failed to process payment");

// Tracing generic dictionaries and collections
var metadata = new Dictionary<string, object>
{
    { "OrderId", "ORD-12345" },
    { "Amount", 99.50 },
    { "Customer", new Dictionary<string, string> { { "Id", "CUST-9" }, { "Region", "NA" } } }
};
logger.INFO(metadata);
```

### 2. Direct API Usage (VB.NET)

```vb
Imports Tracer.automation.level2

Using logger As New TraceManager("MyApp", "logs", True, TraceManager.TraceLevels.INFO)
    logger.Header = "Orders"
    logger.INFO("Processing incoming order")
    logger.ERROR("Order failed validation")
End Using
```

### 3. Modern .NET / ASP.NET Core Dependency Injection

Add the adapter package project reference `SL2Tracer.Extensions.Logging`, then configure in `Program.cs`:

```csharp
using SL2Tracer.Extensions.Logging;

var builder = WebApplication.CreateBuilder(args);

// Register SL2Tracer
builder.Logging.AddSL2Tracer(options =>
{
    options.Application = "OrderService";
    options.Path = Path.Combine(AppContext.BaseDirectory, "logs");
    options.MinimumLevel = TraceManager.TraceLevels.INFO;
    options.MaxFiles = 10;
    options.MaxFileSizeKb = 10240; // 10 MB
    options.MaxDaysKeepingFiles = 14;
});

var app = builder.Build();

app.MapGet("/", (ILogger<Program> logger) =>
{
    logger.LogInformation("Handling root request");
    return "Service healthy!";
});

app.Run();
```

---

## Configuration (`PurgeConfig.xml`)

Place `PurgeConfig.xml` in your application base directory:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<config>
  <!-- Retain rolled logs up to N days -->
  <keepLogs value="10" />
  <!-- Daily cleanup time in minutes after midnight (420 = 07:00 AM) -->
  <time4Clean value="420" />
  <!-- Max active file size before rolling (in KB) -->
  <maxFileSize value="10000" />
  <!-- Max number of historical archive files to retain -->
  <maxFiles value="20" />
</config>
```

---

## Improvements & Modernization in v2.0.0

| Issue in Legacy v1.0 | Modernized in v2.0.0 |
|---|---|
| **Target Framework** | Upgraded from legacy .NET Framework 4.0 / 4.5.2 to modern SDK-style multi-targeting `netstandard2.0` and `net8.0`. |
| **Cross-Platform** | Replaced Windows-only `My.Computer.FileSystem` and hardcoded backslashes `\` with portable `System.IO` (`Path.Combine`, `File.Move`, `FileInfo.Delete`). Builds and runs on Linux, macOS, and Windows. |
| **Thread Safety** | Replaced unsynchronized file writes with `SyncLock` synchronization and `using` file stream blocks, eliminating concurrent `IOException` file sharing crashes and silent log dropping. |
| **Purger Off-by-One Bug** | Fixed inclusive `0 To FilesToDelete` loop in `checkMaxFiles` that deleted 1 more file than `maxFiles`. |
| **Invalid Trace Levels** | Fixed purger methods that called `writeLog(..., 1)` and `writeLog(..., 2)` which were filtered out by `TraceLevel >= 15`. Purger status messages now log at `INFO` and `DEBUG`. |
| **Collection Crashes** | Replaced rigid `Dictionary(Of String, Hashtable)` casting with universal `IDictionary` recursion. Generic dictionaries like `Dictionary<string, string>` and `Dictionary<int, object>` no longer throw `InvalidCastException`. |
| **Missing `ERROR` Level** | Added `ERROR (45)` to `TraceLevels` and implemented `logger.ERROR(msge)`. |
| **Resource Disposal** | Implemented `IDisposable` on `TraceManager` to properly stop and dispose background timers (`System.Timers.Timer`). |
| **Testing** | Added an automated unit test suite (`Tracer.Tests`) with 9 cross-platform xUnit tests. |
| **DI Integration** | Added `SL2Tracer.Extensions.Logging` providing standard `ILogger` and `ILoggerProvider` support. |
| **Console Demo** | Added `Tracer.Demo` cross-platform console runner so the logger can be tested without Windows Forms. |

---

## Building and Running Tests

### Prerequisites
- [.NET 8 SDK](https://dotnet.microsoft.com/download) or later

### Build the Solution
```bash
dotnet build Tracer.sln
```

### Run Unit Tests
```bash
dotnet test Tracer.sln
```

### Run the Interactive Demo
```bash
dotnet run --project Tracer.Demo/Tracer.Demo.csproj
```
