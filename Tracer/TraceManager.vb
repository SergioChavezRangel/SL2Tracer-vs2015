Imports System.Collections
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports System.Xml

Namespace automation.level2

    ''' <summary>
    ''' Thread-safe, cross-platform logger with file rolling, purging, and hierarchical collection tracing.
    ''' </summary>
    Public Class TraceManager
        Implements IDisposable

        Private WithEvents timMinute As Timers.Timer

        Private ReadOnly mLock As New Object()
        Private mDisposed As Boolean = False
        Private clean As Boolean = False
        Private purge As Boolean = False
        Private Config As XmlDocument

        Private mDictSingleTrace As Boolean = False
        Private mMainTracer As Boolean = False
        Private mApplication As String = AppDomain.CurrentDomain.FriendlyName
        Private mTraceLevel As TraceLevels = TraceLevels.NOTSET
        Private mPath As String = AppDomain.CurrentDomain.BaseDirectory
        Private mHeader As String = ""
        Private mTimeFormat As String = "yyyy-MM-dd HH:mm:ss"
        Private mTimeCleanup As Int16 = 420
        Private mMaxFiles As Int16 = 5
        Private mMaxFileSize As Long = (1024 * 1024)
        Private mMaxDaysKeepingFiles As Int16 = 7

#Region "Public Properties"
        Public Property DictSingleTrace() As Boolean
            Get
                Return mDictSingleTrace
            End Get
            Set(ByVal value As Boolean)
                mDictSingleTrace = value
            End Set
        End Property

        Public Property MainTracer() As Boolean
            Get
                Return mMainTracer
            End Get
            Set(ByVal value As Boolean)
                mMainTracer = value
            End Set
        End Property

        Public Property Application() As String
            Get
                Return mApplication
            End Get
            Set(ByVal value As String)
                mApplication = value
            End Set
        End Property

        Public Property TraceLevel() As TraceLevels
            Get
                Return mTraceLevel
            End Get
            Set(ByVal value As TraceLevels)
                mTraceLevel = value
            End Set
        End Property

        Public Property Path() As String
            Get
                Return mPath
            End Get
            Set(ByVal value As String)
                mPath = value
            End Set
        End Property

        Public Property Header() As String
            Get
                Return mHeader
            End Get
            Set(ByVal value As String)
                mHeader = value
            End Set
        End Property

        Public Property TimeFormat() As String
            Get
                Return mTimeFormat
            End Get
            Set(ByVal value As String)
                mTimeFormat = value
            End Set
        End Property

        Public Property TimeCleanup() As Int16
            Get
                Return mTimeCleanup
            End Get
            Set(ByVal value As Int16)
                mTimeCleanup = value
            End Set
        End Property

        Public Property MaxFiles() As Int16
            Get
                Return mMaxFiles
            End Get
            Set(ByVal value As Int16)
                mMaxFiles = value
            End Set
        End Property

        Public Property MaxFileSize() As Long
            Get
                Return mMaxFileSize
            End Get
            Set(ByVal value As Long)
                mMaxFileSize = value
            End Set
        End Property

        Public Property MaxDaysKeepingFiles() As Int16
            Get
                Return mMaxDaysKeepingFiles
            End Get
            Set(ByVal value As Int16)
                mMaxDaysKeepingFiles = value
            End Set
        End Property
#End Region

        Public Enum TraceLevels
            CRITICAL = 50
            [ERROR] = 45
            EXCEPTION = 40
            WARNING = 30
            INFO = 20
            NOTSET = 15
            DEBUG = 10
            VERBOSE = 0
        End Enum

        Private Sub loghandler(ByVal msge As Object, ByVal traceLevel As TraceLevels)
            If msge Is Nothing Then
                writeLog("<null>", traceLevel)
                Return
            End If

            If TypeOf msge Is IDictionary Then
                logDictionary(DirectCast(msge, IDictionary), traceLevel)
            ElseIf TypeOf msge Is Exception Then
                Dim ex As Exception = DirectCast(msge, Exception)
                writeLog(ex.ToString(), traceLevel)
            Else
                writeLog(msge.ToString(), traceLevel)
            End If
        End Sub

        Public Sub CRITICAL(ByVal msge As Object)
            loghandler(msge, TraceLevels.CRITICAL)
        End Sub

        Public Sub [ERROR](ByVal msge As Object)
            loghandler(msge, TraceLevels.[ERROR])
        End Sub

        Public Sub EXCEPTION(ByVal msge As Object)
            loghandler(msge, TraceLevels.EXCEPTION)
        End Sub

        Public Sub WARNING(ByVal msge As Object)
            loghandler(msge, TraceLevels.WARNING)
        End Sub

        Public Sub INFO(ByVal msge As Object)
            loghandler(msge, TraceLevels.INFO)
        End Sub

        Public Sub DEBUG(ByVal msge As Object)
            loghandler(msge, TraceLevels.DEBUG)
        End Sub

        Public Sub VERBOSE(ByVal msge As Object)
            loghandler(msge, TraceLevels.VERBOSE)
        End Sub

        Public Sub New(ByVal sApplication As String, ByVal sPath As String, ByVal isMainTracer As Boolean, ByVal iTraceLevel As TraceLevels)
            MyBase.New()
            newLogger(sApplication, sPath, isMainTracer, iTraceLevel, "")
        End Sub

        Public Sub New(ByVal sApplication As String, ByVal sPath As String, ByVal isMainTracer As Boolean, ByVal Header As String, ByVal iTraceLevel As TraceLevels)
            MyBase.New()
            newLogger(sApplication, sPath, isMainTracer, iTraceLevel, Header)
        End Sub

        Public Sub New()
            MyBase.New()
            newLogger(mApplication, mPath, MainTracer, mTraceLevel, mHeader)
        End Sub

        Public Sub newLogger(ByVal sApplication As String, ByVal sPath As String, ByVal isMainTracer As Boolean, ByVal iTraceLevel As TraceLevels, ByVal Header As String)
            mHeader = Header
            mApplication = sApplication
            mPath = sPath
            TraceLevel = iTraceLevel
            MainTracer = isMainTracer

            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)

            If MainTracer Then
                writeLog("'Tracer' Starting", TraceLevels.INFO)
                writeLog("    Path '" & mPath & "'", TraceLevels.INFO)
                writeLog("    File '" & mApplication & "'", TraceLevels.INFO)
                writeLog("'Tracer' Started OK", TraceLevels.INFO)
                PurgerStart(mApplication, mPath)
            End If
        End Sub

        ''' <summary>
        ''' Logs any IDictionary (Hashtable, Dictionary, etc.) recursively with single-trace or multi-line formatting.
        ''' </summary>
        Public Sub logDictionary(ByVal d2Log As IDictionary, ByVal iTraceLevel As TraceLevels)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            If iTraceLevel >= TraceLevel AndAlso d2Log IsNot Nothing Then
                Try
                    If mDictSingleTrace Then
                        Dim sb As New StringBuilder()
                        Dim first As Boolean = True
                        For Each de As DictionaryEntry In d2Log
                            If Not first Then
                                sb.Append(vbTab)
                            End If
                            first = False

                            sb.Append(If(de.Key IsNot Nothing, de.Key.ToString(), "<null>")).Append("=")
                            If TypeOf de.Value Is IDictionary Then
                                Dim subFirst As Boolean = True
                                For Each subDe As DictionaryEntry In DirectCast(de.Value, IDictionary)
                                    If Not subFirst Then
                                        sb.Append(";")
                                    End If
                                    subFirst = False
                                    sb.Append(If(subDe.Key IsNot Nothing, subDe.Key.ToString(), "<null>"))
                                    sb.Append(":")
                                    sb.Append(If(subDe.Value IsNot Nothing, subDe.Value.ToString(), "<null>"))
                                Next
                            Else
                                sb.Append(If(de.Value IsNot Nothing, de.Value.ToString(), "<null>"))
                            End If
                        Next
                        writeLog("   " & sb.ToString(), iTraceLevel)
                    Else
                        For Each de As DictionaryEntry In d2Log
                            writeLog("   " & If(de.Key IsNot Nothing, de.Key.ToString(), "<null>"), iTraceLevel)
                            If TypeOf de.Value Is IDictionary Then
                                logDictionary(DirectCast(de.Value, IDictionary), iTraceLevel)
                            Else
                                writeLog("       " & If(de.Value IsNot Nothing, de.Value.ToString(), "<null>"), iTraceLevel)
                            End If
                        Next
                    End If
                Catch ex As Exception
                    writeLog("Unable to trace object dictionary! " & ex.Message, TraceLevels.EXCEPTION)
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Legacy overload for Hashtable logging.
        ''' </summary>
        Public Sub logHT(ByVal ht2Log As Hashtable, ByVal iTraceLevel As TraceLevels)
            logDictionary(DirectCast(ht2Log, IDictionary), iTraceLevel)
        End Sub

        ''' <summary>
        ''' Legacy overload for Dictionary(Of String, Hashtable) logging.
        ''' </summary>
        Public Sub logDictionary(ByVal d2Log As Dictionary(Of String, Hashtable), ByVal iTraceLevel As TraceLevels)
            logDictionary(DirectCast(d2Log, IDictionary), iTraceLevel)
        End Sub

        Private Sub writeLog(ByVal log As String, ByVal iTraceLevel As TraceLevels)
            If iTraceLevel >= TraceLevel Then
                Dim msg As String
                Dim Level As String = "[" & iTraceLevel.ToString() & "]"
                Dim time_stamp As String = DateTime.Now.ToString(mTimeFormat)
                If Not String.IsNullOrEmpty(mHeader) Then
                    msg = String.Format("{0} {2,-11} [{1}].[{4}] {3}", time_stamp, mApplication, Level, log, mHeader)
                Else
                    msg = String.Format("{0} {2,-11} [{1}] {3}", time_stamp, mApplication, Level, log)
                End If

                SyncLock mLock
                    Try
                        If Not IO.Directory.Exists(mPath) Then
                            IO.Directory.CreateDirectory(mPath)
                        End If
                        Dim fileName As String = IO.Path.Combine(mPath, mApplication & ".Tracer.txt")
                        Using writer As StreamWriter = IO.File.AppendText(fileName)
                            writer.WriteLine(msg)
                        End Using
                    Catch ex As Exception
                        System.Diagnostics.Trace.WriteLine("SL2Tracer write error: " & ex.Message)
                    End Try
                End SyncLock
            End If
        End Sub

#Region "Purge"
        Private Sub createFile()
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            writeLog("I: File Log Restarted", TraceLevels.INFO)
            Try
                writeLog("I: Process running Since " & Process.GetCurrentProcess().StartTime.ToString(), TraceLevels.INFO)
            Catch ex As Exception
                ' Ignore if process start time cannot be queried on current platform
            End Try
        End Sub

        Private Sub PurgerStart(ByVal sApplication As String, ByVal sPath As String)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            mApplication = sApplication
            mPath = sPath

            Try
                Dim configPath As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PurgeConfig.xml")
                If IO.File.Exists(configPath) Then
                    Config = New XmlDocument()
                    Config.Load(configPath)

                    Dim keepLogsNode = Config.SelectSingleNode("/config/keepLogs")
                    If keepLogsNode IsNot Nothing AndAlso keepLogsNode.Attributes("value") IsNot Nothing Then
                        mMaxDaysKeepingFiles = CShort(keepLogsNode.Attributes("value").Value)
                    End If

                    Dim time4CleanNode = Config.SelectSingleNode("/config/time4Clean")
                    If time4CleanNode IsNot Nothing AndAlso time4CleanNode.Attributes("value") IsNot Nothing Then
                        mTimeCleanup = CShort(time4CleanNode.Attributes("value").Value)
                    End If

                    Dim maxFileSizeNode = Config.SelectSingleNode("/config/maxFileSize")
                    If maxFileSizeNode IsNot Nothing AndAlso maxFileSizeNode.Attributes("value") IsNot Nothing Then
                        mMaxFileSize = CLng(maxFileSizeNode.Attributes("value").Value)
                    End If

                    Dim maxFilesNode = Config.SelectSingleNode("/config/maxFiles")
                    If maxFilesNode IsNot Nothing AndAlso maxFilesNode.Attributes("value") IsNot Nothing Then
                        mMaxFiles = CShort(maxFilesNode.Attributes("value").Value)
                    End If
                Else
                    writeLog("PurgeConfig.xml not found; default values used", TraceLevels.DEBUG)
                End If
            Catch ex As Exception
                writeLog("Default Values used: " & ex.Message, TraceLevels.EXCEPTION)
            End Try

            writeLog("Object " & sApplication & " Loading.......", TraceLevels.INFO)
            writeLog("'Purger' Started", TraceLevels.INFO)
            writeLog("    Path '" & mPath & "'", TraceLevels.INFO)
            writeLog("    File '" & mApplication & "'", TraceLevels.INFO)
            writeLog("    Max Days for Old Files -> " & mMaxDaysKeepingFiles.ToString(), TraceLevels.INFO)
            writeLog("    Max Size for Files Kb  -> " & mMaxFileSize.ToString(), TraceLevels.INFO)
            writeLog("    Max Files              -> " & mMaxFiles.ToString(), TraceLevels.INFO)

            Dim mins As Long
            Dim hrs As Long
            hrs = Math.DivRem(mTimeCleanup, 60, mins)

            writeLog("    Time for Cleanup       -> " & String.Format("{0:00}:{1:00}", hrs, mins), TraceLevels.INFO)
            writeLog("'Purger' Configure OK", TraceLevels.INFO)

            If timMinute IsNot Nothing Then
                timMinute.Stop()
                timMinute.Dispose()
            End If

            timMinute = New Timers.Timer(60000)
            AddHandler timMinute.Elapsed, AddressOf timMinute_Tick
            timMinute.AutoReset = True
            timMinute.Enabled = True
            timMinute.Start()

            checkMaxSize(mMaxFileSize)
            checkMaxFiles(mMaxFiles)
            cleanFiles(mMaxDaysKeepingFiles)
        End Sub

        Private Sub timMinute_Tick(ByVal sender As System.Object, ByVal e As Timers.ElapsedEventArgs)
            Dim hr As Integer = DateTime.Now.Hour
            Dim min As Integer = DateTime.Now.Minute

            Dim spanMins As Integer = (hr * 60) + min
            Dim time4CleanEnd As Integer = mTimeCleanup + 10

            If (spanMins >= mTimeCleanup) AndAlso (spanMins <= time4CleanEnd) AndAlso (clean = False) Then
                clean = True
                renameFile()
                cleanFiles(mMaxDaysKeepingFiles)
            ElseIf (spanMins < mTimeCleanup) OrElse (spanMins > time4CleanEnd) Then
                clean = False
            End If

            If (min < 2) AndAlso (purge = False) Then
                purge = True
                checkMaxSize(mMaxFileSize)
                checkMaxFiles(mMaxFiles)
            ElseIf (min > 1) Then
                purge = False
            End If
        End Sub

        Public Sub TriggerPurgeCheck()
            checkMaxSize(mMaxFileSize)
            checkMaxFiles(mMaxFiles)
            cleanFiles(mMaxDaysKeepingFiles)
        End Sub

        Private Sub renameFile()
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            Dim currentFile As String = IO.Path.Combine(mPath, mApplication & ".Tracer.txt")
            writeLog("LogFile-> " & currentFile, TraceLevels.INFO)
            Try
                If IO.File.Exists(currentFile) Then
                    writeLog("LogFile Already Exist", TraceLevels.DEBUG)
                    Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                    Dim newFileName As String = String.Format("{0}.Tracer_{1}.txt", mApplication, timestamp)
                    Dim newFullPath As String = IO.Path.Combine(mPath, newFileName)

                    Dim counter As Integer = 1
                    While IO.File.Exists(newFullPath)
                        newFileName = String.Format("{0}.Tracer_{1}_{2}.txt", mApplication, timestamp, counter)
                        newFullPath = IO.Path.Combine(mPath, newFileName)
                        counter += 1
                    End While

                    SyncLock mLock
                        IO.File.Move(currentFile, newFullPath)
                    End SyncLock

                    writeLog("OldLogFile-> " & newFileName, TraceLevels.INFO)
                    createFile()
                Else
                    writeLog("LogFile Not Found", TraceLevels.DEBUG)
                    writeLog("Can't rename File ", TraceLevels.WARNING)
                End If
            Catch ex As Exception
                writeLog("Can't rename File " & ex.Message, TraceLevels.EXCEPTION)
            End Try
        End Sub

        Private Sub checkMaxFiles(ByVal maxFiles As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            Try
                Dim dirInfo As New DirectoryInfo(mPath)
                If Not dirInfo.Exists Then Return

                Dim orderedFiles = dirInfo.GetFiles(mApplication & ".Tracer*.txt") _
                    .Where(Function(f) Not f.Name.Equals(mApplication & ".Tracer.txt", StringComparison.OrdinalIgnoreCase)) _
                    .OrderBy(Function(x) x.LastWriteTime) _
                    .ToList()

                Dim filesFound As Integer = orderedFiles.Count
                If filesFound > maxFiles Then
                    Dim filesToDelete As Integer = filesFound - maxFiles
                    writeLog("'Purger' Deleting " & filesToDelete & " Files, MaxFiles Full", TraceLevels.INFO)
                    For i As Integer = 0 To filesToDelete - 1
                        Try
                            orderedFiles(i).Delete()
                            writeLog("File deleted-> " & orderedFiles(i).Name, TraceLevels.INFO)
                        Catch delEx As Exception
                            writeLog("Error deleting file " & orderedFiles(i).Name & ": " & delEx.Message, TraceLevels.WARNING)
                        End Try
                    Next
                End If
            Catch ex As Exception
                writeLog("Something was wrong MaxFiles Full -> " & ex.Message, TraceLevels.EXCEPTION)
            End Try
        End Sub

        Private Sub checkMaxSize(ByVal maxfileSizeKb As Long)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            Try
                Dim currentFile As String = IO.Path.Combine(mPath, mApplication & ".Tracer.txt")
                If IO.File.Exists(currentFile) Then
                    Dim fileInfo As New FileInfo(currentFile)
                    Dim sizeInKb As Long = fileInfo.Length \ 1024
                    If sizeInKb > maxfileSizeKb Then
                        writeLog("'Purger' Change File, MaxSize Full", TraceLevels.INFO)
                        renameFile()
                    End If
                End If
            Catch ex As Exception
                writeLog("'Purger' Change File Error -> '" & ex.Message & "'", TraceLevels.EXCEPTION)
            End Try
        End Sub

        Private Sub cleanFiles(ByVal keepHistoryDays As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            Dim currentFile As String = IO.Path.Combine(mPath, mApplication & ".Tracer.txt")
            writeLog("LogFileCleanup-> " & currentFile, TraceLevels.INFO)
            Try
                Dim dirInfo As New DirectoryInfo(mPath)
                If Not dirInfo.Exists Then Return

                Dim rolledFiles = dirInfo.GetFiles(mApplication & ".Tracer*.txt") _
                    .Where(Function(f) Not f.Name.Equals(mApplication & ".Tracer.txt", StringComparison.OrdinalIgnoreCase)) _
                    .ToList()

                For Each rolledFile As FileInfo In rolledFiles
                    Dim daysOld As Double = (DateTime.Now - rolledFile.LastWriteTime).TotalDays
                    If daysOld > keepHistoryDays Then
                        Try
                            rolledFile.Delete()
                            writeLog("File deleted-> " & rolledFile.Name, TraceLevels.INFO)
                        Catch delEx As Exception
                            writeLog("Can't delete File-> " & delEx.Message, TraceLevels.WARNING)
                        End Try
                    End If
                Next
            Catch ex As Exception
                writeLog("Can't delete File-> " & ex.Message, TraceLevels.EXCEPTION)
            End Try
        End Sub
#End Region

#Region "IDisposable Support"
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not mDisposed Then
                If disposing Then
                    If timMinute IsNot Nothing Then
                        timMinute.Stop()
                        timMinute.Dispose()
                        timMinute = Nothing
                    End If
                End If
                mDisposed = True
            End If
        End Sub
#End Region

    End Class

End Namespace