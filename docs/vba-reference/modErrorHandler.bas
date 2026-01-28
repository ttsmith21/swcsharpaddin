Attribute VB_Name = "modErrorHandler"
Option Explicit

' Handles errors, logs details, and optionally displays messages
Sub HandleError(Optional ByVal SubName As String = "Unknown", _
                Optional ByVal AdditionalInfo As String = "", _
                Optional ByVal Severity As String = "Error")

    Dim errMsg As String
    Dim errLine As String
    Dim logSuccess As Boolean

    ' Capture line number if available
    If Erl <> 0 Then
        errLine = "Line Number: " & Erl & vbCrLf
    Else
        errLine = "Line Number: (Not Available)" & vbCrLf
    End If

    ' Construct error message
    errMsg = "---------- ERROR REPORT ----------" & vbCrLf & _
             "Date & Time: " & Now & vbCrLf & _
             "Subroutine: " & SubName & vbCrLf & _
             errLine & _
             "Error Number: " & Err.Number & vbCrLf & _
             "Error Description: " & Err.Description & vbCrLf & _
             "Severity: " & Severity & vbCrLf

    If AdditionalInfo <> "" Then
        errMsg = errMsg & "Additional Info: " & AdditionalInfo & vbCrLf
    End If

    ' Log error if logging is enabled
    If modConfig.LOG_ENABLED Then logSuccess = LogError(errMsg)

    ' Show message based on severity level
    Select Case Severity
        Case "Fatal"
            MsgBox errMsg, vbCritical, "Fatal Error - Execution Halted"
            ' Ensure CleanupExcel is handled safely
            On Error Resume Next
            If modConfig.AUTO_CLOSE_EXCEL Then
                modExcelHelper.CleanupExcel
            End If
            On Error GoTo 0
            Exit Sub

        Case "Critical"
            MsgBox errMsg, vbCritical, "Critical Error"

        Case "Warning"
            If modConfig.SHOW_WARNINGS Then MsgBox "Warning: " & AdditionalInfo, vbExclamation, "Warning"

        Case "Info"
            If modConfig.ENABLE_DEBUG_MODE Then Debug.Print "INFO: " & AdditionalInfo
    End Select

    ' Notify user if logging failed
    If modConfig.LOG_ENABLED And Not logSuccess And Severity = "Critical" Then
        MsgBox "Failed to write error log. Check permissions for: " & modConfig.ERROR_LOG_PATH, vbExclamation, "Logging Error"
    End If

    ' Clear the error
    Err.Clear
End Sub

' Logs error messages to a text file with retry mechanism
Function LogError(ByVal errMsg As String) As Boolean
    Dim logFile As String
    Dim fso As Object
    Dim fileHandle As Integer
    Dim retryCount As Integer
    Dim success As Boolean

    logFile = modConfig.ERROR_LOG_PATH
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ensure log directory exists
    Dim logFolder As String
    logFolder = fso.GetParentFolderName(logFile)

    If Not fso.FolderExists(logFolder) Then
        On Error GoTo DirectoryErrorHandler
        fso.CreateFolder logFolder
        On Error GoTo 0
    End If

    ' Retry file writing in case of file locks
    retryCount = 0
    success = False
    Do While retryCount < 3 And Not success
        On Error Resume Next
        fileHandle = FreeFile
        Open logFile For Append As #fileHandle
        Print #fileHandle, Now & " - " & errMsg
        Close #fileHandle
        success = (Err.Number = 0)
        On Error GoTo 0

        If Not success Then retryCount = retryCount + 1
    Loop

    If Not success Then GoTo FileErrorHandler

    LogError = True
    Exit Function

' Directory creation error handler
DirectoryErrorHandler:
    HandleError "LogError", "Could not create log directory - " & logFolder, "Error"
    LogError = False
    Exit Function

' File write error handler
FileErrorHandler:
    HandleError "LogError", "Failed to write to log file: " & logFile, "Error"
    LogError = False
    Exit Function
End Function


