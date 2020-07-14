Imports System.IO
Imports System.Reflection
Imports System.Xml

Namespace automation.level2

    Public Class TraceManager
        Private WithEvents timMinute As Timers.Timer

        Private clean As Boolean = False
        Private Config As XmlDocument

        Private List As XmlNodeList
        Private log As String
        Private Node As XmlNode
        Private purge As Boolean = False
        Private mDictSingleTrace = False
        Private mMainTracer As Boolean = False
        Private mApplication As String = AppDomain.CurrentDomain.FriendlyName
        Private mTraceLevel As TraceLevels = 15
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

        Enum TraceLevels
            CRITICAL = 50
            EXCEPTION = 40
            WARNING = 30
            INFO = 20
            NOTSET = 15
            DEBUG = 10
            VERBOSE = 0
        End Enum

        Private Sub loghandler(ByVal msge As Object, ByVal traceLevel As TraceLevels)
            If TypeOf msge Is Hashtable Then
                Dim obj As New Hashtable
                obj = msge
                logHT(obj, traceLevel)
            ElseIf msge.GetType().ToString().IndexOf("Dictionary") >= 0 Then
                'msge.GetType().GetGenericTypeDefinition().IsAssignableFrom(TypeOf (Dictionary <,>) Is IDictionary Then
                'o is IDictionary && o.GetType().IsGenericType && o.GetType().GetGenericTypeDefinition().IsAssignableFrom(TypeOf (Dictionary <,>))
                'Dim obj As New Dictionary(Of Object, Object)
                'obj = msge

                logDictionary(msge, traceLevel)
            Else
                writeLog(msge.ToString(), traceLevel)
            End If
        End Sub

        Public Sub CRITICAL(ByVal msge As Object)
            loghandler(msge, TraceLevels.CRITICAL)
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
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            'System.Diagnostics.EventLog.WriteEntry(Application, "Instance Tracer")

            newLogger(sApplication, sPath, isMainTracer, iTraceLevel, "")

        End Sub

        Public Sub New(ByVal sApplication As String, ByVal sPath As String, ByVal isMainTracer As Boolean, ByVal Header As String, ByVal iTraceLevel As TraceLevels)
            MyBase.New()

            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            'System.Diagnostics.EventLog.WriteEntry(Application, "Instance Tracer")

            newLogger(sApplication, sPath, isMainTracer, iTraceLevel, Header)

        End Sub

        Public Sub newLogger(ByVal sApplication As String, ByVal sPath As String, ByVal isMainTracer As Boolean, ByVal iTraceLevel As TraceLevels, ByVal Header As String)
            mHeader = Header
            mApplication = sApplication
            mPath = sPath
            TraceLevel = iTraceLevel
            MainTracer = isMainTracer
            If MainTracer Then
                writeLog("'Tracer' Starting", TraceLevels.INFO)
                writeLog("    Path '" & mPath & "'", TraceLevels.INFO)
                writeLog("    File '" & mApplication & "'", TraceLevels.INFO)
                writeLog("'Tracer' Started OK", TraceLevels.INFO)
                PurgerStart(mApplication, mPath)
            End If
        End Sub

        Public Sub New()
            MyBase.New()
            newLogger(mApplication, mPath, MainTracer, mTraceLevel, mHeader)
        End Sub



        Private Sub logDictionary(ByVal d2Log As Dictionary(Of String, Hashtable), ByVal iTraceLevel As TraceLevels)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            If iTraceLevel >= TraceLevel Then
                Try
                    If mDictSingleTrace Then
                        Dim sTempVar As String = ""
                        For Each kvp As KeyValuePair(Of String, Hashtable) In d2Log
                            sTempVar = sTempVar & kvp.Key.ToString & "="
                            If TypeOf (kvp.Value) Is Hashtable Then
                                For Each de As DictionaryEntry In kvp.Value
                                    sTempVar = sTempVar & de.Key.ToString & ":" & de.Value.ToString & vbTab
                                Next
                            Else
                                sTempVar = sTempVar & kvp.Value.ToString & vbTab
                            End If

                        Next
                    Else
                        For Each kvp As KeyValuePair(Of String, Hashtable) In d2Log
                            'If (TypeOf (kvp.Key) Is String) Or (TypeOf (kvp.Key) Is Int16) Or (TypeOf (kvp.Key) Is Int32) Or (TypeOf (kvp.Key) Is Int64) Or (TypeOf (kvp.Key) Is Integer) Or (TypeOf (kvp.Key) Is Double) Then
                            writeLog("   " & kvp.Key.ToString, iTraceLevel)
                            'End If

                            'If TypeOf (kvp.Value) Is Hashtable Then
                            logHT(kvp.Value, iTraceLevel)
                            'ElseIf (TypeOf (kvp.Value) Is String) Or (TypeOf (kvp.Value) Is Int16) Or (TypeOf (kvp.Value) Is Int32) Or (TypeOf (kvp.Value) Is Int64) Or (TypeOf (kvp.Value) Is Integer) Or (TypeOf (kvp.Value) Is Double) Then
                            '    writeLog("       " & kvp.Value.ToString, iTraceLevel)
                            'End If

                            'If kvp.Key.GetType.ToString = "System.String" Then
                            writeLog("       " & kvp.Key.ToString, iTraceLevel)
                            ' End If
                        Next
                    End If
                Catch ex As Exception
                    writeLog("Unable to trace object dictionary! " & ex.Message, TraceLevels.EXCEPTION)
                End Try
            End If
        End Sub

        'Private Sub logDictionary(ByVal d2Log As Dictionary(Of String, Hashtable), ByVal iTraceLevel As Int16)
        '    writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 12)
        '    If iTraceLevel <= TraceLevel Then
        '        Try
        '            For Each kvp As KeyValuePair(Of String, Hashtable) In d2Log
        '                writeLog("   " & kvp.Key.ToString, 1)
        '                logHT(kvp.Value, iTraceLevel)
        '                writeLog("       " & kvp.Key.ToString, 1)
        '            Next
        '        Catch ex As Exception
        '            writeLog("Unable to trace object dictionary! " & ex.Message, TraceLevels.EXCEPTION)
        '        End Try
        '    End If
        'End Sub

        Private Sub logHT(ByVal ht2Log As Hashtable, ByVal iTraceLevel As TraceLevels) ', ByVal bHTrace As Boolean
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), TraceLevels.VERBOSE)
            If iTraceLevel >= TraceLevel Then
                Try
                    If mDictSingleTrace Then
                        Dim sTempVar As String = ""
                        For Each de As DictionaryEntry In ht2Log
                            sTempVar = sTempVar & de.Key.ToString & ":"
                            If TypeOf (de.Value) Is Hashtable Then
                                For Each de2 As DictionaryEntry In de.Value
                                    sTempVar = sTempVar & de2.Key.ToString & ":" & de2.Value.ToString & vbTab
                                Next
                            Else
                                sTempVar = sTempVar & de.Value.ToString & vbTab
                            End If
                        Next
                        writeLog("   " & sTempVar, iTraceLevel)
                    Else
                        For Each de As DictionaryEntry In ht2Log
                            If TypeOf (de.Value) Is Hashtable Then
                                writeLog("       " & de.Key.ToString & " -> ", iTraceLevel)
                                logHT(de.Value, iTraceLevel)
                            Else
                                writeLog("       " & de.Key.ToString & " -> " & de.Value.ToString, iTraceLevel)
                            End If
                        Next
                    End If
                Catch ex As Exception
                    writeLog("E: Unable to trace hashtable! ", TraceLevels.EXCEPTION)
                End Try
            End If
        End Sub

        'Private Sub logHT(ByVal ht2Log As Hashtable, ByVal iTraceLevel As Int16)
        '    writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 12)
        '    If iTraceLevel <= TraceLevel Then
        '        Try
        '            For Each de As DictionaryEntry In ht2Log
        '                writeLog("       " & de.Key.ToString & " -> " & de.Value.ToString, 1)
        '            Next
        '        Catch ex As Exception
        '            writeLog("E: Unable to trace hashtable! ", 1)
        '        End Try
        '    End If
        'End Sub

        'Private Sub logHT(ByVal ht2Log As Hashtable, ByVal sIndent As String, ByVal iTraceLevel As Int16)
        '    writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 12)
        '    If iTraceLevel <= TraceLevel Then
        '        Try
        '            If sIndent = "" Then
        '                sIndent = "       "
        '            End If
        '            For Each de As DictionaryEntry In ht2Log
        '                writeLog(sIndent & de.Key.ToString & " -> " & de.Value.ToString, 1)
        '            Next
        '        Catch ex As Exception
        '            writeLog("E: Unable to trace hashtable! ", 1)
        '        End Try
        '    End If
        'End Sub

        'Private Sub writeLog(ByVal chType As String, ByVal log As String, ByVal iTraceLevel As TraceLevels)
        '    'writeLog(MethodInfo.GetCurrentMethod().ToString(), 2)
        '    If iTraceLevel <= TraceLevel Then
        '        Dim msg As String
        '        msg = mApplication & " " & chType & ": " & log
        '        'Dim fecha As String = System.DateTime.Now.ToString(LogHeader)
        '        Try
        '            Dim fileName As String = mPath & mApplication & ".Tracer.txt"
        '            Dim writer As StreamWriter = File.AppendText(fileName)
        '            writer.WriteLine(System.DateTime.Now.ToString(mTimeFormat) & mHeader & msg)
        '            writer.Close()
        '        Catch ex As Exception
        '            'Do Nothing
        '        End Try
        '    End If
        'End Sub

        Private Sub writeLog(ByVal log As String, ByVal iTraceLevel As TraceLevels)
            If iTraceLevel >= TraceLevel Then
                Dim msg As String
                Dim Level As String = "[" & iTraceLevel.ToString() & "]"
                Dim time_stamp As String = DateTime.Now.ToString(mTimeFormat)
                If mHeader.Length > 0 Then
                    msg = String.Format("{0} {2,-11} [{1}].[{4}] {3}", time_stamp, mApplication, Level, log, mHeader)
                Else
                    msg = String.Format("{0} {2,-11} [{1}] {3}", time_stamp, mApplication, Level, log)
                End If
                'Dim fecha As String = System.DateTime.Now.ToString(LogHeader)
                Try
                    Dim fileName As String = String.Format("{0}{1}.Tracer.txt", mPath, mApplication)
                    Dim writer As StreamWriter = File.AppendText(fileName)
                    writer.WriteLine(msg)
                    writer.Close()
                Catch ex As Exception
                    'Do Nothing
                End Try
            End If
        End Sub

#Region "Purge"
        Private Sub createFile()
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            log = "I: File Log Restarted"
            writeLog(log, TraceLevels.INFO)
            log = "I: Process running Since " & Process.GetCurrentProcess.StartTime
            writeLog(log, TraceLevels.INFO)

        End Sub

        Private Sub PurgerStart(ByVal sApplication As String, ByVal sPath As String) 'ByVal maxDays As Int16, ByVal timeClean As Int16, ByVal maxSize As Int16)

            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            mApplication = sApplication
            mPath = sPath

            Try
                Config = New XmlDocument()
                Config.Load(System.AppDomain.CurrentDomain.BaseDirectory() & "PurgeConfig.xml")
                'List = Config.SelectNodes("/config/tags/tag")
                mMaxDaysKeepingFiles = CInt(Config.SelectSingleNode("/config/keepLogs").Attributes.GetNamedItem("value").Value)
                mTimeCleanup = CInt(Config.SelectSingleNode("/config/time4Clean").Attributes.GetNamedItem("value").Value)
                mMaxFileSize = CLng(Config.SelectSingleNode("/config/maxFileSize").Attributes.GetNamedItem("value").Value)
                mMaxFiles = CInt(Config.SelectSingleNode("/config/maxFiles").Attributes.GetNamedItem("value").Value)
            Catch ex As Exception
                writeLog("Default Values used", TraceLevels.EXCEPTION)
            End Try

            writeLog("Object " & sApplication & " Loading.......", 1)

            writeLog("'Purger' Started", 1)
            writeLog("    Path '" & mPath & "'", 1)
            writeLog("    File '" & mApplication & "'", 1)
            writeLog("    Max Days for Old Files -> " & mMaxDaysKeepingFiles.ToString, 1)
            writeLog("    Max Size for Files Kb  -> " & mMaxFileSize.ToString, 1)
            writeLog("    Max Files              -> " & mMaxFiles.ToString, 1)

            Dim mins As Long
            Dim hrs As Long
            hrs = Math.DivRem(mTimeCleanup, 60, mins)

            writeLog("    Time for Cleanup       -> " & String.Format("{0:00}:{1:00}", hrs, mins), 1)
            writeLog("'Purger' Configure OK", 1)

            timMinute = New Timers.Timer(60000)
            timMinute.Enabled = True
            timMinute.Start()

            checkMaxSize(mMaxFileSize)
            checkMaxFiles(mMaxFiles)
            cleanFiles(mMaxDaysKeepingFiles)
        End Sub

        Private Sub timMinute_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timMinute.Elapsed
            Dim hr, min As Int16

            hr = DateTime.Now.Hour
            min = DateTime.Now.Minute

            Dim spanMins As Integer
            spanMins = (hr * 60) + min
            Dim time4CleanEnd As Integer
            time4CleanEnd = mTimeCleanup + 10

            If (spanMins >= mTimeCleanup) And (spanMins <= time4CleanEnd) And (clean = False) Then
                clean = True
                renameFile()
                cleanFiles(mMaxDaysKeepingFiles)
            ElseIf (spanMins < mTimeCleanup) Or (spanMins > time4CleanEnd) Then
                clean = False
            End If

            If (min < 2) And (purge = False) Then
                purge = True
                checkMaxSize(mMaxFileSize)
                checkMaxFiles(mMaxFiles)
            ElseIf (min > 1) Then
                purge = False
            End If

        End Sub

        Private Sub renameFile()
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            Dim siNo As Boolean = False
            log = "LogFile-> " & mPath & mApplication & ".Tracer.txt"
            writeLog(log, TraceLevels.INFO)
            Try
                If File.Exists(mPath & mApplication & ".Tracer.txt") Then
                    siNo = True
                    writeLog("LogFile Already Exist", 1)
                Else
                    writeLog("LogFile Not Found", 1)
                End If
                If siNo Then
                    My.Computer.FileSystem.RenameFile(mPath & mApplication & ".Tracer.txt", mApplication & ".Tracer" & DateAndTime.Now.ToString(" ddMMyyyy HHmmss") & ".txt")
                    log = "OldLogFile-> " & mApplication & ".Tracer" & DateAndTime.Now.ToString(" ddMMyyyy HHmmss") & ".txt"
                    writeLog(log, TraceLevels.INFO)
                    createFile()
                Else
                    log = "Can't rename File "
                    writeLog(log, TraceLevels.INFO)
                End If
            Catch ex As Exception
                log = "Can't rename File " & ex.Message
                writeLog(log, TraceLevels.EXCEPTION)
            End Try
        End Sub

        Private Sub checkMaxFiles(ByVal maxFiles As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            Try
                Dim fileNames() As String
                fileNames = Directory.GetFiles(mPath, mApplication & ".Tracer *.txt")

                Dim orderedFiles = New DirectoryInfo(mPath).GetFiles(mApplication & ".Tracer *.txt").OrderBy(Function(x) x.LastWriteTime)
                For Each f As FileInfo In orderedFiles
                    Console.WriteLine(String.Format("{0,-15} {1,12}", f.Name, f.CreationTime.ToString))
                Next

                Dim FilesFound As Int16
                FilesFound = orderedFiles.Count
                If FilesFound > maxFiles Then
                    Dim FilesToDelete As Int16
                    FilesToDelete = (FilesFound - maxFiles)
                    writeLog("'Purger' Deleting " & (FilesToDelete + 1) & " Files, MaxFiles Full", 1)
                    For i = 0 To FilesToDelete
                        My.Computer.FileSystem.DeleteFile(orderedFiles(i).FullName)
                        log = "File deleted-> " & orderedFiles(i).Name
                        writeLog(log, TraceLevels.INFO)
                    Next
                End If
            Catch ex As Exception
                log = "Something was wrong MaxFiles Full -> " & ex.Message.ToString
                writeLog(log, TraceLevels.INFO)
            End Try
        End Sub

        Private Sub checkMaxSize(ByVal maxfileSize As Long)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            Try
                If File.Exists(mPath & mApplication & ".Tracer.txt") Then
                    Dim fileInfo As FileInfo
                    fileInfo = My.Computer.FileSystem.GetFileInfo(mPath & mApplication & ".Tracer.txt")
                    Dim x As Long
                    x = fileInfo.Length
                    x = x / 1024
                    If x > maxfileSize Then
                        writeLog("'Purger' Change File, MaxSize Full", 1)
                        renameFile()
                    End If
                End If
            Catch ex As Exception
                writeLog("'Purger' Change File Error -> '" & ex.Message & "'", 1)
            End Try
        End Sub

        Private Sub cleanFiles(ByVal keepHistory As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            log = "LogFileCleanup-> " & mPath & mApplication & ".Tracer.txt"
            writeLog(log, TraceLevels.INFO)
            Try
                Dim fileNames As New Collection
                For Each file As String In Directory.GetFiles(mPath, mApplication & ".Tracer *.txt") 'My.Application.Info.DirectoryPath.ToString, Application & ".Tracer *.txt")
                    log = "OldLogFiles-> " & file
                    writeLog(log, TraceLevels.INFO)
                    fileNames.Add(file.ToString)
                Next
                For Each fileName As String In fileNames
                    Dim info As New FileInfo(fileName)
                    Dim infoDate As Date
                    infoDate = DateAndTime.Now
                    Dim days As Int32 = 0
                    days = DateDiff(DateInterval.Day, info.LastWriteTime, DateAndTime.Now)
                    If days > keepHistory Then
                        My.Computer.FileSystem.DeleteFile(fileName)
                        log = "File deleted-> " & fileName.Substring(fileName.LastIndexOf("\") + 1)
                        writeLog(log, TraceLevels.INFO)
                    End If
                Next
            Catch ex As Exception
                log = "Can't delete File-> " & ex.Message.ToString
                writeLog(log, TraceLevels.INFO)
            End Try
        End Sub
#End Region

    End Class

End Namespace