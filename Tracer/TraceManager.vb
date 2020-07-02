
Imports System.IO
Imports System.Reflection
Imports System.Xml

Namespace AHMSA.BOFyCC

    Public Class TraceManager
        Dim log As String
        Private Application As String = "AppNameNotSet"
        Private Path As String = "C:\"
        Private TraceLevel As Int16 = 1


        ' Purger Var Names
        'Dim traceLogger As Tracer.AHMSA.BOFyCC.TraceManager
        Dim keepLogs As Int16
        Dim time4Clean As Int16
        Dim maxFileSize As Long
        Dim maxFiles As Int16
        Dim clean As Boolean = False
        Dim purge As Boolean = False
        Dim WithEvents timMinute As System.Timers.Timer
        Dim Config As XmlDocument
        Dim List As XmlNodeList
        Dim Node As XmlNode
        Private LogRowHeader As String
        Private LogTimeFormat As String

        Private Sub PurgerStart(ByVal sApplication As String, ByVal sPath As String) 'ByVal maxDays As Int16, ByVal timeClean As Int16, ByVal maxSize As Int16)

            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            Application = sApplication
            Path = sPath

            Config = New XmlDocument()
            Config.Load(System.AppDomain.CurrentDomain.BaseDirectory() & "PurgeConfig.xml")
            List = Config.SelectNodes("/config/tags/tag")
            keepLogs = CInt(Config.SelectSingleNode("/config/keepLogs").Attributes.GetNamedItem("value").Value)
            time4Clean = CInt(Config.SelectSingleNode("/config/time4Clean").Attributes.GetNamedItem("value").Value)
            maxFileSize = CLng(Config.SelectSingleNode("/config/maxFileSize").Attributes.GetNamedItem("value").Value)
            maxFiles = CInt(Config.SelectSingleNode("/config/maxFiles").Attributes.GetNamedItem("value").Value)

            writeLog("I", "Object " & sApplication & " Loading.......", 1)

            writeLog("I", "'Purger' Started", 1)
            writeLog("    Path '" & Path & "'", 1)
            writeLog("    File '" & Application & "'", 1)
            writeLog("    Max Days for Old Files -> " & keepLogs.ToString, 1)
            writeLog("    Max Size for Files Kb  -> " & maxFileSize.ToString, 1)
            writeLog("    Max Files              -> " & maxFiles.ToString, 1)

            Dim mins As Long
            Dim hrs As Long
            hrs = Math.DivRem(time4Clean, 60, mins)

            writeLog("    Time for Cleanup       -> " & String.Format("{0:00}:{1:00}", hrs, mins), 1)
            writeLog("I", "'Purger' Configure OK", 1)

            timMinute = New System.Timers.Timer(60000)
            timMinute.Enabled = True
            timMinute.Start()

            checkMaxSize(maxFileSize)
            checkMaxFiles(maxFiles)
            cleanFiles(keepLogs)
        End Sub

        Private Sub timMinute_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timMinute.Elapsed
            Dim hr, min As Int16

            hr = DateTime.Now.Hour
            min = DateTime.Now.Minute

            Dim spanMins As Integer
            spanMins = (hr * 60) + min
            Dim time4CleanEnd As Integer
            time4CleanEnd = time4Clean + 10

            If (spanMins >= time4Clean) And (spanMins <= time4CleanEnd) And (clean = False) Then
                clean = True
                renameFile()
                cleanFiles(keepLogs)
            ElseIf (spanMins < time4Clean) Or (spanMins > time4CleanEnd) Then
                clean = False
            End If

            If (min < 2) And (purge = False) Then
                purge = True
                checkMaxSize(maxFileSize)
                checkMaxFiles(maxFiles)
            ElseIf (min > 1) Then
                purge = False
            End If

        End Sub

        Public Sub New(ByVal sApplication As String, ByVal sPath As String, ByVal isMainTracer As Boolean, ByVal iTraceLevel As Int16)
            MyBase.New()
            LogTimeFormat = "dd-MM-yyyy HH:mm:ss "
            LogRowHeader = ""
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            'System.Diagnostics.EventLog.WriteEntry(Application, "Instance Tracer")
            Application = sApplication
            Path = sPath
            TraceLevel = iTraceLevel
            If isMainTracer Then
                writeLog("I", " 'Tracer' Started", 1)
                writeLog("    Path '" & Path & "'", 1)
                writeLog("    File '" & Application & "'", 1)
                writeLog("I", " 'Tracer' Started OK", 1)
                PurgerStart(sApplication, sPath)
            End If


        End Sub

        Public Sub New(ByVal sApplication As String, ByVal sPath As String, ByVal isMainTracer As Boolean, ByVal Header As String, ByVal iTraceLevel As Int16)
            MyBase.New()
            LogRowHeader = Header
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            'System.Diagnostics.EventLog.WriteEntry(Application, "Instance Tracer")
            Application = sApplication
            Path = sPath
            TraceLevel = iTraceLevel
            If isMainTracer Then
                writeLog("I", " 'Tracer' Started", 1)
                writeLog("    Path '" & Path & "'", 1)
                writeLog("    File '" & Application & "'", 1)
                writeLog("I", " 'Tracer' Started OK", 1)
                PurgerStart(sApplication, sPath)
            End If


        End Sub

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub writeLog(ByVal chType As String, ByVal log As String, ByVal iTraceLevel As Int16)
            'writeLog(MethodInfo.GetCurrentMethod().ToString(), 2)
            If iTraceLevel <= TraceLevel Then
                Dim msg As String
                msg = Application & " " & chType & ": " & log
                'Dim fecha As String = System.DateTime.Now.ToString(LogHeader)
                Try
                    Dim fileName As String = Path & Application & ".Tracer.txt"
                    Dim writer As StreamWriter = File.AppendText(fileName)
                    writer.WriteLine(System.DateTime.Now.ToString(LogTimeFormat) & LogRowHeader & msg)
                    writer.Close()
                Catch ex As Exception
                    'Do Nothing
                End Try
            End If
        End Sub

        Public Sub writeLog(ByVal log As String, ByVal iTraceLevel As Int16)
            If iTraceLevel <= TraceLevel Then
                Dim msg As String
                msg = Application & " " & log
                'Dim fecha As String = System.DateTime.Now.ToString(LogHeader)
                Try
                    Dim fileName As String = Path & Application & ".Tracer.txt"
                    Dim writer As StreamWriter = File.AppendText(fileName)
                    writer.WriteLine(System.DateTime.Now.ToString(LogTimeFormat) & LogRowHeader & msg)
                    writer.Close()
                Catch ex As Exception
                    'Do Nothing
                End Try
            End If
        End Sub

        Private Sub createFile()
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            log = "I: File Log Restarted"
            writeLog(log, 1)
            log = "I: Process running Since " & Process.GetCurrentProcess.StartTime
            writeLog(log, 1)

        End Sub

        Public Sub renameFile()
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            Dim siNo As Boolean = False
            log = "LogFile-> " & Path & Application & ".Tracer.txt"
            writeLog("I", log, 1)
            Try
                If File.Exists(Path & Application & ".Tracer.txt") Then
                    siNo = True
                    writeLog("I", "LogFile Already Exist", 1)
                Else
                    writeLog("I", "LogFile Not Found", 1)
                End If
                If siNo Then
                    My.Computer.FileSystem.RenameFile(Path & Application & ".Tracer.txt", Application & ".Tracer" & DateAndTime.Now.ToString(" ddMMyyyy HHmmss") & ".txt")
                    log = "OldLogFile-> " & Application & ".Tracer" & DateAndTime.Now.ToString(" ddMMyyyy HHmmss") & ".txt"
                    writeLog("I", log, 1)
                    createFile()
                Else
                    log = "Can't rename File "
                    writeLog("E", log, 1)
                End If
            Catch ex As Exception
                log = "Can't rename File " & ex.Message
                writeLog("E", log, 1)
            End Try
        End Sub

        Public Sub cleanFiles(ByVal keepHistory As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            log = "LogFileCleanup-> " & Path & Application & ".Tracer.txt"
            writeLog("I", log, 1)
            Try
                Dim fileNames As New Collection
                For Each file As String In Directory.GetFiles(Path, Application & ".Tracer *.txt") 'My.Application.Info.DirectoryPath.ToString, Application & ".Tracer *.txt")
                    log = "OldLogFiles-> " & file
                    writeLog("I", log, 1)
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
                        writeLog("I", log, 1)
                    End If
                Next
            Catch ex As Exception
                log = "Can't delete File-> " & ex.Message.ToString
                writeLog("E", log, 1)
            End Try
        End Sub

        Public Sub checkMaxSize(ByVal maxfileSize As Long)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            Try
                If File.Exists(Path & Application & ".Tracer.txt") Then
                    Dim fileInfo As FileInfo
                    fileInfo = My.Computer.FileSystem.GetFileInfo(Path & Application & ".Tracer.txt")
                    Dim x As Long
                    x = fileInfo.Length
                    x = x / 1024
                    If x > maxfileSize Then
                        writeLog("I", "'Purger' Change File, MaxSize Full", 1)
                        renameFile()
                    End If
                End If
            Catch ex As Exception
                writeLog("E", "'Purger' Change File Error -> '" & ex.Message & "'", 1)
            End Try
        End Sub

        Public Sub checkMaxFiles(ByVal maxFiles As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 2)
            Try
                Dim fileNames() As String
                fileNames = Directory.GetFiles(Path, Application & ".Tracer *.txt")

                Dim orderedFiles = New System.IO.DirectoryInfo(Path).GetFiles(Application & ".Tracer *.txt").OrderBy(Function(x) x.LastWriteTime)
                For Each f As System.IO.FileInfo In orderedFiles
                    Console.WriteLine(String.Format("{0,-15} {1,12}", f.Name, f.CreationTime.ToString))
                Next

                Dim FilesFound As Int16
                FilesFound = orderedFiles.Count
                If FilesFound > maxFiles Then
                    Dim FilesToDelete As Int16
                    FilesToDelete = (FilesFound - maxFiles)
                    writeLog("I", "'Purger' Deleting " & (FilesToDelete + 1) & " Files, MaxFiles Full", 1)
                    For i = 0 To FilesToDelete
                        My.Computer.FileSystem.DeleteFile(orderedFiles(i).FullName)
                        log = "File deleted-> " & orderedFiles(i).Name
                        writeLog("I", log, 1)
                    Next
                End If
            Catch ex As Exception
                log = "Something was wrong MaxFiles Full -> " & ex.Message.ToString
                writeLog("E", log, 1)
            End Try
        End Sub

        Public Sub logHT(ByVal ht2Log As Hashtable, ByVal bHTrace As Boolean, ByVal iTraceLevel As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 12)
            If iTraceLevel <= TraceLevel Then
                Try
                    If bHTrace Then
                        Dim sTempVar As String = ""
                        For Each de As DictionaryEntry In ht2Log
                            sTempVar = sTempVar & de.Key.ToString & ":" & de.Value.ToString & vbTab
                        Next
                        writeLog("   " & sTempVar, 1)
                    Else
                        For Each de As DictionaryEntry In ht2Log
                            writeLog("       " & de.Key.ToString & " -> " & de.Value.ToString, 1)
                        Next
                    End If
                Catch ex As Exception
                    writeLog("E: Cannot trace hashtable! ", 1)
                End Try
            End If
        End Sub

        Public Sub logHT(ByVal ht2Log As Hashtable, ByVal iTraceLevel As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 12)
            If iTraceLevel <= TraceLevel Then
                Try
                    For Each de As DictionaryEntry In ht2Log
                        writeLog("       " & de.Key.ToString & " -> " & de.Value.ToString, 1)
                    Next

                Catch ex As Exception
                    writeLog("E: Cannot trace hashtable! ", 1)
                End Try
            End If
        End Sub

        Public Sub logHT(ByVal ht2Log As Hashtable, ByVal sIndent As String, ByVal iTraceLevel As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 12)
            If iTraceLevel <= TraceLevel Then
                Try
                    If sIndent = "" Then
                        sIndent = "       "
                    End If
                    For Each de As DictionaryEntry In ht2Log
                        writeLog(sIndent & de.Key.ToString & " -> " & de.Value.ToString, 1)
                    Next
                Catch ex As Exception
                    writeLog("E: Cannot trace hashtable! ", 1)
                End Try
            End If
        End Sub

        Public Sub logDictionary(ByVal d2Log As Dictionary(Of Object, Object), ByVal iTraceLevel As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 12)
            If iTraceLevel <= TraceLevel Then
                Try
                    For Each kvp As KeyValuePair(Of Object, Object) In d2Log
                        If (TypeOf (kvp.Key) Is String) Or (TypeOf (kvp.Key) Is Int16) Or (TypeOf (kvp.Key) Is Int32) Or (TypeOf (kvp.Key) Is Int64) Or (TypeOf (kvp.Key) Is Integer) Or (TypeOf (kvp.Key) Is Double) Then
                            writeLog("   " & kvp.Key.ToString, 1)
                        End If

                        If TypeOf (kvp.Value) Is Hashtable Then
                            logHT(kvp.Value, iTraceLevel)
                        ElseIf (TypeOf (kvp.Value) Is String) Or (TypeOf (kvp.Value) Is Int16) Or (TypeOf (kvp.Value) Is Int32) Or (TypeOf (kvp.Value) Is Int64) Or (TypeOf (kvp.Value) Is Integer) Or (TypeOf (kvp.Value) Is Double) Then
                            writeLog("       " & kvp.Value.ToString, 1)
                        End If

                        If kvp.Key.GetType.ToString = "System.String" Then
                            writeLog("       " & kvp.Key.ToString, 1)
                        End If
                    Next
                Catch ex As Exception
                    writeLog("E: Cannot trace object dictionary! " & ex.Message, 1)
                End Try
            End If
        End Sub

        Public Sub logDictionary(ByVal d2Log As Dictionary(Of String, Hashtable), ByVal iTraceLevel As Int16)
            writeLog(">>>>>>>>> " + MethodInfo.GetCurrentMethod().ToString(), 12)
            If iTraceLevel <= TraceLevel Then
                Try
                    For Each kvp As KeyValuePair(Of String, Hashtable) In d2Log
                        writeLog("   " & kvp.Key.ToString, 1)
                        logHT(kvp.Value, iTraceLevel)
                        writeLog("       " & kvp.Key.ToString, 1)
                    Next
                Catch ex As Exception
                    writeLog("E: Cannot trace object dictionary! " & ex.Message, 1)
                End Try
            End If
        End Sub
    End Class

End Namespace

