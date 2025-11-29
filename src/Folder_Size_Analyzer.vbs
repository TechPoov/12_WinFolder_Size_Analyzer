'===============================================================
' WinFolderSizeCrawler.vbs
' TechPoov / WinFolderSizeCrawler
'
' - Reads jobs from <scriptBaseName>.config (same name as .vbs)
'   Example: WinFolderSizeCrawler.vbs -> WinFolderSizeCrawler.config
'            6.vbs                    -> 6.config
'
' - For each [JobX], scans folders/files and writes:
'     * Data CSV
'     * Log CSV
'     * Optional Debug log
' - Also writes a run-level summary CSV in the script folder:
'     * summary_YYYYMMDD_HHMM.csv (one row per job, or CONFIG_NOT_FOUND)
'
' Design:
'   - All core actions are Boolean-return Functions:
'       bResult = Function(...)
'       If Not bResult Then ...
'   - Uses NoError() helper to centralize Err handling.
'   - Uses BuildCsvLine(fields) for all CSV rows.
'   - No WScript.Echo, no GoTo: Task-Scheduler friendly.
'===============================================================

Option Explicit

'-----------------------------
' Application constants
'-----------------------------
Const APP_NAME              = "WinFolderSizeCrawler"
Const APP_VERSION           = "1.0"
Const CONFIG_FILE_EXTENSION = ".config"

' Mode constants
Const MODE_FILES   = "FILES"
Const MODE_FOLDER  = "FOLDER"
Const MODE_BOTH    = "BOTH"

' Log level constants
Const LOG_LEVEL_INFO  = "INFO"
Const LOG_LEVEL_WARN  = "WARN"
Const LOG_LEVEL_ERROR = "ERROR"

' Status constants
Const STATUS_START           = "START"
Const STATUS_INIT            = "INIT"
Const STATUS_VALIDATE        = "VALIDATE"
Const STATUS_SCAN_START      = "SCAN_START"
Const STATUS_SCAN_OK         = "SCAN_OK"
Const STATUS_SCAN_SKIP       = "SCAN_SKIP"
Const STATUS_OK              = "OK"
Const STATUS_FAIL            = "FAIL"
Const STATUS_RENAME          = "RENAME"
Const STATUS_END             = "END"
Const STATUS_CONFIG_MISSING  = "CONFIG_NOT_FOUND"

'-----------------------------
' Global runtime objects/vars
'-----------------------------
Dim gFSO
Dim gRunContext
Dim gErrorContext
Dim gJobs()
Dim gJobCount

Dim gConfigLines
Dim gConfigFilePath

' Per job globals
Dim gCurrentJob
Dim gJobName
Dim gRootFolderPath
Dim gOutputFolderPath
Dim gMode
Dim gMaxDepth
Dim gDebugMode

Dim gMainCsvPath
Dim gLogCsvPath
Dim gDebugFilePath

Dim gDataTS
Dim gLogTS
Dim gDebugTS

Dim gStartTime
Dim gJobStartTime

Dim gTotalFolders
Dim gTotalFiles
Dim gTotalErrors

' Summary file
Dim gSummaryTS
Dim gSummaryFilePath
Dim gScriptFolder

'===============================================================
' Entry point
'===============================================================
Sub Main()
    Dim bResult

    bResult = InitializeEnvironment()
    If bResult Then bResult = InitSummaryFile()

    If bResult Then
        bResult = LoadConfigFile(gConfigFilePath)
        If Not bResult Then
            Call WriteConfigMissingSummaryRow()
            Call CloseSummaryFile()
            Exit Sub
        End If
    End If

    If bResult Then bResult = ParseConfigIntoJobs()
    If bResult Then bResult = ValidateJobs()
    If bResult Then bResult = RunAllJobs()

    Call CloseSummaryFile()
End Sub

Main

'===============================================================
' Common error helper
'===============================================================
Function NoError()
    NoError = (Err.Number = 0)
    If Not NoError Then Err.Clear
End Function

'===============================================================
' Initialization
'===============================================================
Function InitializeEnvironment()
    On Error Resume Next
    InitializeEnvironment = False

    Set gFSO = CreateObject("Scripting.FileSystemObject")
    If Not NoError() Then Exit Function

    gScriptFolder = gFSO.GetParentFolderName(WScript.ScriptFullName)

    Dim scriptBaseName
    scriptBaseName = gFSO.GetBaseName(WScript.ScriptName)

    gConfigFilePath = gScriptFolder & "\" & scriptBaseName & CONFIG_FILE_EXTENSION

    Set gRunContext   = CreateObject("Scripting.Dictionary")
    Set gErrorContext = CreateObject("Scripting.Dictionary")
    If Not NoError() Then Exit Function

    gRunContext.RemoveAll
    gErrorContext.RemoveAll

    ReDim gJobs(-1)
    gJobCount = -1

    gStartTime = Now

    InitializeEnvironment = True
    On Error GoTo 0
End Function

Function InitSummaryFile()
    On Error Resume Next
    InitSummaryFile = False

    Dim tsStamp
    tsStamp = BuildTimeStampFrom(gStartTime)

    gSummaryFilePath = gScriptFolder & "\summary_" & tsStamp & ".csv"
    Set gSummaryTS = gFSO.CreateTextFile(gSummaryFilePath, True)
    If Not NoError() Then Exit Function

    gSummaryTS.WriteLine BuildCsvLine(Array( _
        "JobName", "ScanFolder", "OutputFolder", "Mode", "MaxDepth", "Debug", _
        "Status", "FolderCount", "FileCount", "ErrorCount", _
        "JobStart", "JobEnd", "DurationSeconds", _
        "DataCsvPath", "LogCsvPath", "DebugLogPath" _
    ))
    If Not NoError() Then Exit Function

    InitSummaryFile = True
    On Error GoTo 0
End Function

Sub CloseSummaryFile()
    On Error Resume Next
    If Not gSummaryTS Is Nothing Then
        gSummaryTS.Close
        Set gSummaryTS = Nothing
    End If
    On Error GoTo 0
End Sub

'===============================================================
' Config loading & parsing
'===============================================================
Function LoadConfigFile(configPath)
    On Error Resume Next
    LoadConfigFile = False

    Dim ts, contents

    If Not gFSO.FileExists(configPath) Then Exit Function

    Set ts = gFSO.OpenTextFile(configPath, 1)
    If Not NoError() Then Exit Function

    contents = ts.ReadAll
    ts.Close
    If Not NoError() Then Exit Function

    gConfigLines = Split(contents, vbCrLf)

    LoadConfigFile = True
    On Error GoTo 0
End Function

Function ParseConfigIntoJobs()
    On Error Resume Next
    ParseConfigIntoJobs = False

    Dim i, line, trimmed, sectionName
    Dim currentJob

    Set currentJob = Nothing
    ReDim gJobs(-1)
    gJobCount = -1

    For i = 0 To UBound(gConfigLines)
        line = gConfigLines(i)
        trimmed = Trim(line)

        If trimmed = "" Then
        ElseIf Left(trimmed, 1) = "#" Then
        ElseIf Left(trimmed, 1) = "[" And InStr(trimmed, "]") > 1 Then
            If Not currentJob Is Nothing Then
                If currentJob.Count > 0 Then
                    gJobCount = gJobCount + 1
                    ReDim Preserve gJobs(gJobCount)
                    Set gJobs(gJobCount) = currentJob
                End If
            End If

            sectionName = Mid(trimmed, 2, InStr(trimmed, "]") - 2)
            sectionName = Trim(sectionName)

            Set currentJob = CreateObject("Scripting.Dictionary")
            currentJob.RemoveAll
            currentJob("jobname") = sectionName

        ElseIf InStr(trimmed, "=") > 0 Then
            If Not currentJob Is Nothing Then
                Dim pos, key, value, hashPos
                pos = InStr(trimmed, "=")
                key = Trim(Left(trimmed, pos - 1))
                value = Trim(Mid(trimmed, pos + 1))

                hashPos = InStr(value, "#")
                If hashPos > 0 Then
                    value = Trim(Left(value, hashPos - 1))
                End If

                If key <> "" Then
                    currentJob(LCase(key)) = value
                End If
            End If
        End If
    Next

    If Not currentJob Is Nothing Then
        If currentJob.Count > 0 Then
            gJobCount = gJobCount + 1
            ReDim Preserve gJobs(gJobCount)
            Set gJobs(gJobCount) = currentJob
        End If
    End If

    ParseConfigIntoJobs = (gJobCount >= 0)

    On Error GoTo 0
End Function

Function ValidateJobs()
    Dim i, job, okAll, b
    okAll = True

    For i = 0 To gJobCount
        Set job = gJobs(i)
        b = ValidateJob(job)
        If Not b Then okAll = False
    Next

    ValidateJobs = okAll
End Function

Function ValidateJob(job)
    Dim hasError
    hasError = False

    If Not job.Exists("scanfolder") Then hasError = True
    If Not job.Exists("outputfolder") Then hasError = True
    If Not job.Exists("mode") Then hasError = True

    If Not job.Exists("maxdepth") Then job("maxdepth") = "0"
    If Not job.Exists("debug") Then job("debug") = "0"

    ValidateJob = Not hasError
End Function

'===============================================================
' Summary row when config is missing
'===============================================================
Sub WriteConfigMissingSummaryRow()
    On Error Resume Next

    If gSummaryTS Is Nothing Then Exit Sub

    Dim fields
    fields = Array( _
        "(GLOBAL)", _
        gConfigFilePath, _
        "", _
        "", _
        "", _
        "", _
        STATUS_CONFIG_MISSING, _
        "", "", "", _
        "", "", "", _
        "", "", "" _
    )

    gSummaryTS.WriteLine BuildCsvLine(fields)
    On Error GoTo 0
End Sub

'===============================================================
' Running all jobs
'===============================================================
Function RunAllJobs()
    Dim i, job, okAll, b
    okAll = True

    For i = 0 To gJobCount
        Set job = gJobs(i)
        b = RunJob(job)
        If Not b Then okAll = False
    Next

    RunAllJobs = okAll
End Function

Function RunJob(job)
    On Error Resume Next
    RunJob = False

    Dim jobStatus
    Dim jobEndTime
    Dim b

    jobStatus = "FAIL"

    Set gCurrentJob = job
    gJobName = job("jobname")

    gJobStartTime = Now
    gTotalFolders = 0
    gTotalFiles   = 0
    gTotalErrors  = 0

    b = InitializeJobContext(job)
    If Not b Then
        jobEndTime = Now
        b = WriteSummaryRow(job, jobStatus, gJobStartTime, jobEndTime)
        Exit Function
    End If

    b = BuildOutputFileNames()
    If Not b Then
        gTotalErrors = gTotalErrors + 1
        jobEndTime = Now
        b = WriteSummaryRow(job, jobStatus, gJobStartTime, jobEndTime)
        Exit Function
    End If

    b = InitCsvFiles()
    If Not b Then
        gTotalErrors = gTotalErrors + 1
        jobEndTime = Now
        b = WriteSummaryRow(job, jobStatus, gJobStartTime, jobEndTime)
        Exit Function
    End If

    b = LogEvent(LOG_LEVEL_INFO, "Job", STATUS_START, _
                 "Job [" & gJobName & "] started. Root=" & gRootFolderPath, 0)

    b = RunFolderScan()
    If Not b Then
        gTotalErrors = gTotalErrors + 1
        Call LogEvent(LOG_LEVEL_ERROR, "RunFolderScan", STATUS_FAIL, _
                      "RunFolderScan failed for root: " & gRootFolderPath, 0)
    End If

    b = LogJobEndStatus()

    jobEndTime = Now
    If gTotalErrors > 0 Then
        jobStatus = "FAIL"
    Else
        jobStatus = "OK"
    End If

    b = WriteSummaryRow(job, jobStatus, gJobStartTime, jobEndTime)

    Call CloseCsvFiles()

    RunJob = True
    On Error GoTo 0
End Function

Function WriteSummaryRow(job, jobStatus, jobStartTime, jobEndTime)
    On Error Resume Next
    WriteSummaryRow = False

    If gSummaryTS Is Nothing Then
        WriteSummaryRow = True
        Exit Function
    End If

    Dim durationSec
    durationSec = DateDiff("s", jobStartTime, jobEndTime)

    Dim debugPath
    debugPath = ""
    If gDebugMode = 1 Then
        debugPath = gDebugFilePath
    End If

    Dim fields
    fields = Array( _
        job("jobname"), _
        job("scanfolder"), _
        job("outputfolder"), _
        gMode, _
        gMaxDepth, _
        gDebugMode, _
        jobStatus, _
        gTotalFolders, _
        gTotalFiles, _
        gTotalErrors, _
        FormatTimestamp(jobStartTime), _
        FormatTimestamp(jobEndTime), _
        durationSec, _
        gMainCsvPath, _
        gLogCsvPath, _
        debugPath _
    )

    gSummaryTS.WriteLine BuildCsvLine(fields)
    If Not NoError() Then Exit Function

    WriteSummaryRow = True
    On Error GoTo 0
End Function

Function InitializeJobContext(job)
    On Error Resume Next
    InitializeJobContext = False

    gRootFolderPath   = Trim(CStr(job("scanfolder")))
    gOutputFolderPath = Trim(CStr(job("outputfolder")))
    gMode             = Trim(CStr(job("mode")))
    gMaxDepth         = CLng(Trim(CStr(job("maxdepth"))))
    gDebugMode        = CInt(Trim(CStr(job("debug"))))

    Dim m
    m = UCase(gMode)

    Select Case m
        Case "FILES", "FILE"
            gMode = MODE_FILES
        Case "FOLDERS", "FOLDER"
            gMode = MODE_FOLDER
        Case "BOTH"
            gMode = MODE_BOTH
        Case Else
            Exit Function
    End Select

    If Not gFSO.FolderExists(gRootFolderPath) Then Exit Function
    If Not gFSO.FolderExists(gOutputFolderPath) Then Exit Function

    InitializeJobContext = True
    On Error GoTo 0
End Function

'===============================================================
' Output file naming & initialization
'===============================================================
Function BuildOutputFileNames()
    On Error Resume Next
    BuildOutputFileNames = False

    Dim timeStamp, prefix

    timeStamp = BuildTimeStamp()

    prefix = gJobName & "_" & gMode
    prefix = SanitizeFileNamePart(prefix)

    gMainCsvPath   = gOutputFolderPath & "\" & prefix & "_"       & timeStamp & ".csv"
    gLogCsvPath    = gOutputFolderPath & "\" & prefix & "_Log_"   & timeStamp & ".csv"
    gDebugFilePath = gOutputFolderPath & "\" & prefix & "_Debug_" & timeStamp & ".log"

    If Not NoError() Then Exit Function

    BuildOutputFileNames = True
    On Error GoTo 0
End Function

Function BuildTimeStamp()
    BuildTimeStamp = BuildTimeStampFrom(Now)
End Function

Function BuildTimeStampFrom(dt)
    BuildTimeStampFrom = _
        Year(dt) & _
        Right("0" & Month(dt), 2) & _
        Right("0" & Day(dt), 2) & "_" & _
        Right("0" & Hour(dt), 2) & _
        Right("0" & Minute(dt), 2)
End Function

Function SanitizeFileNamePart(originalName)
    On Error Resume Next

    Dim namePart, badChars, i, c
    namePart = originalName
    badChars = "\/:*?""<>|"

    For i = 1 To Len(badChars)
        c = Mid(badChars, i, 1)
        namePart = Replace(namePart, c, "_")
    Next

    SanitizeFileNamePart = namePart
    On Error GoTo 0
End Function

Function InitCsvFiles()
    On Error Resume Next
    InitCsvFiles = False

    Set gDataTS = gFSO.CreateTextFile(gMainCsvPath, True)
    If Not NoError() Then Exit Function

    gDataTS.WriteLine BuildCsvLine(Array( _
        "ItemPath", "ItemType", "SizeBytes", _
        "CreatedOn", "LastAccessOn", "LastModifiedOn", "FilesCount" _
    ))
    If Not NoError() Then Exit Function

    Set gLogTS = gFSO.CreateTextFile(gLogCsvPath, True)
    If Not NoError() Then Exit Function

    gLogTS.WriteLine BuildCsvLine(Array( _
        "Timestamp", "Level", "StepName", "Status", _
        "Message", "ErrorNumber", "DurationSeconds" _
    ))
    If Not NoError() Then Exit Function

    If gDebugMode = 1 Then
        Set gDebugTS = gFSO.CreateTextFile(gDebugFilePath, True)
        If NoError() Then
            gDebugTS.WriteLine "Timestamp | StepName | Message"
            If Not NoError() Then Exit Function
        Else
            Err.Clear
        End If
    End If

    InitCsvFiles = True
    On Error GoTo 0
End Function

Sub CloseCsvFiles()
    On Error Resume Next

    If Not gDataTS Is Nothing Then
        gDataTS.Close
        Set gDataTS = Nothing
    End If

    If Not gLogTS Is Nothing Then
        gLogTS.Close
        Set gLogTS = Nothing
    End If

    If Not gDebugTS Is Nothing Then
        gDebugTS.Close
        Set gDebugTS = Nothing
    End If

    On Error GoTo 0
End Sub

'===============================================================
' CSV helpers
'===============================================================
Function BuildCsvLine(fields)
    Dim i, out
    out = ""

    For i = 0 To UBound(fields)
        If i > 0 Then out = out & ","
        out = out & SafeCsv(fields(i))
    Next

    BuildCsvLine = out
End Function

Function SafeCsv(value)
    Dim s
    s = CStr(value)
    s = Replace(s, """", """""")
    SafeCsv = """" & s & """"
End Function

Function FormatTimestamp(dt)
    If IsDate(dt) Then
        Dim d
        d = CDate(dt)
        FormatTimestamp = _
            Year(d) & "-" & _
            Right("0" & Month(d), 2) & "-" & _
            Right("0" & Day(d), 2) & " " & _
            Right("0" & Hour(d), 2) & ":" & _
            Right("0" & Minute(d), 2) & ":" & _
            Right("0" & Second(d), 2)
    Else
        FormatTimestamp = ""
    End If
End Function

'===============================================================
' Folder scanning (FUNCTION version)
'===============================================================
Function RunFolderScan()
    On Error Resume Next
    RunFolderScan = False

    Dim totalSize, totalFiles
    Dim b

    totalSize  = 0
    totalFiles = 0

    b = LogEvent(LOG_LEVEL_INFO, "RunFolderScan", STATUS_SCAN_START, _
                 "Starting scan at root: " & gRootFolderPath, 0)

    b = DebugLog("RunFolderScan", "Root=" & gRootFolderPath & _
                 ", Mode=" & gMode & ", MaxDepth=" & gMaxDepth)

    b = ScanFolder(gRootFolderPath, 1, totalSize, totalFiles)

    RunFolderScan = b
    On Error GoTo 0
End Function

Function ScanFolder(folderPath, currentDepth, ByRef totalSize, ByRef totalFiles)
    On Error Resume Next
    ScanFolder = False

    Dim f, file, subFolder
    Dim folderSize, folderFiles
    Dim childSize, childFiles
    Dim b

    totalSize  = 0
    totalFiles = 0

    Set f = gFSO.GetFolder(folderPath)
    If Err.Number <> 0 Then
        b = HandleError("ScanFolder", "Failed to get folder", folderPath)
        b = LogEvent(LOG_LEVEL_WARN, "ScanFolder", STATUS_SCAN_SKIP, _
                     "Skipped folder (GetFolder failed): " & folderPath, Err.Number)
        gTotalErrors = gTotalErrors + 1
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    gTotalFolders = gTotalFolders + 1
    b = DebugLog("ScanFolder", "Enter Path=" & folderPath & ", Depth=" & currentDepth)

    folderSize  = 0
    folderFiles = 0

    For Each file In f.Files
        folderSize  = folderSize + file.Size
        folderFiles = folderFiles + 1
        gTotalFiles = gTotalFiles + 1

        If gMode = MODE_FILES Or gMode = MODE_BOTH Then
            b = WriteDataRow(file.Path, "FILE", file.Size, _
                             file.DateCreated, file.DateLastAccessed, file.DateLastModified, 1)
        End If

        If Err.Number <> 0 Then
            b = HandleError("ScanFolder", "Error processing file", file.Path)
            gTotalErrors = gTotalErrors + 1
            Err.Clear
        End If
    Next

    If gMaxDepth = 0 Or currentDepth < gMaxDepth Then
        For Each subFolder In f.SubFolders
            childSize  = 0
            childFiles = 0

            b = ScanFolder(subFolder.Path, currentDepth + 1, childSize, childFiles)

            folderSize  = folderSize + childSize
            folderFiles = folderFiles + childFiles
        Next
    End If

    If gMode = MODE_FOLDER Or gMode = MODE_BOTH Then
        b = WriteDataRow(f.Path, "FOLDER", folderSize, _
                         f.DateCreated, f.DateLastAccessed, f.DateLastModified, folderFiles)
    End If

    b = DebugLog("ScanFolder", "SCAN_OK Path=" & f.Path)

    totalSize  = folderSize
    totalFiles = folderFiles

    b = DebugLog("ScanFolder", "Exit Path=" & folderPath & ", Size=" & _
                 folderSize & ", Files=" & folderFiles)

    ScanFolder = True
    On Error GoTo 0
End Function

Function WriteDataRow(itemPath, itemType, sizeBytes, createdOn, lastAccessOn, lastModifiedOn, filesCount)
    On Error Resume Next
    WriteDataRow = False

    If gDataTS Is Nothing Then
        WriteDataRow = True
        Exit Function
    End If

    Dim fields
    fields = Array( _
        itemPath, _
        itemType, _
        sizeBytes, _
        FormatTimestamp(createdOn), _
        FormatTimestamp(lastAccessOn), _
        FormatTimestamp(lastModifiedOn), _
        filesCount _
    )

    gDataTS.WriteLine BuildCsvLine(fields)
    If Not NoError() Then Exit Function

    WriteDataRow = True
    On Error GoTo 0
End Function

'===============================================================
' Logging helpers
'===============================================================
Function LogEvent(level, stepName, status, message, errNumber)
    LogEvent = LogEventWithDuration(level, stepName, status, message, errNumber, "")
End Function

Function LogEventWithDuration(level, stepName, status, message, errNumber, durationSec)
    On Error Resume Next
    LogEventWithDuration = False

    If gLogTS Is Nothing Then
        LogEventWithDuration = True
        Exit Function
    End If

    Dim fields
    fields = Array( _
        FormatTimestamp(Now), _
        level, _
        stepName, _
        status, _
        message, _
        errNumber, _
        durationSec _
    )

    gLogTS.WriteLine BuildCsvLine(fields)
    If Not NoError() Then Exit Function

    LogEventWithDuration = True
    On Error GoTo 0
End Function

Function LogJobEndStatus()
    On Error Resume Next
    LogJobEndStatus = False

    Dim endTime, durationSec, msg

    endTime     = Now
    durationSec = DateDiff("s", gJobStartTime, endTime)

    msg = "Job [" & gJobName & "] completed. Folders=" & gTotalFolders & _
          ", Files=" & gTotalFiles & ", Errors=" & gTotalErrors

    LogJobEndStatus = LogEventWithDuration(LOG_LEVEL_INFO, "Job", STATUS_END, msg, 0, durationSec)

    If gTotalErrors > 0 Then
        Call LogEvent( _
            LOG_LEVEL_WARN, _
            "Job", _
            STATUS_FAIL, _
            "Job [" & gJobName & "] completed with " & gTotalErrors & _
                " errors. Check ERROR entries or debug log for details.", _
            0 _
        )
    End If

    On Error GoTo 0
End Function

Function DebugLog(stepName, message)
    On Error Resume Next
    DebugLog = False

    If gDebugMode = 0 Then
        DebugLog = True
        Exit Function
    End If
    If gDebugTS Is Nothing Then
        DebugLog = True
        Exit Function
    End If

    gDebugTS.WriteLine FormatTimestamp(Now) & " | " & stepName & " | " & message
    If Not NoError() Then Exit Function

    DebugLog = True
    On Error GoTo 0
End Function

'===============================================================
' Error handler
'===============================================================
Function HandleError(funcName, customMessage, filePath)
    On Error Resume Next
    HandleError = False

    gErrorContext("Function") = funcName
    gErrorContext("Message")  = customMessage
    gErrorContext("FilePath") = filePath
    gErrorContext("Number")   = Err.Number

    Dim fullMsg
    fullMsg = customMessage
    If filePath <> "" Then
        fullMsg = fullMsg & " | Path: " & filePath
    End If
    If Err.Number <> 0 Then
        fullMsg = fullMsg & " | Err " & Err.Number & ": " & Err.Description
    End If

    Dim b
    b = LogEvent(LOG_LEVEL_ERROR, funcName, STATUS_FAIL, fullMsg, Err.Number)
    b = DebugLog(funcName, "ERROR: " & fullMsg)

    Err.Clear
    HandleError = True
    On Error GoTo 0
End Function
