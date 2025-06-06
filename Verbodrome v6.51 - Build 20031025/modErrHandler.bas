Attribute VB_Name = "modErrHandler"
Option Explicit
'
' error handler constants
'
'Public Const errExit As Byte = 0
'Public Const errResume As Byte = 1
'Public Const errSelect As Byte = 2
Public Const errNext As Byte = 3
'
' public stack constants
'
Public errStackFlag As Integer
Public Const errPush As Byte = 0
Public Const errPop As Byte = 1
'Public Const errList As Byte = 2
'Public Const errFile As Byte = 3
Public gblNProc As Integer ' stack pointer
Public gblAProc() As String ' array of routines
'
Private blnInErrorHandling As Boolean
'
Public Function errHandler(lngErrNumber, strErrDescription) As Integer
    On Error Resume Next
    '
    If blnPM_Presentation Or blnInErrorHandling Then
        errHandler = errNext
    Else
        Dim bytOriginalMousePointer As Byte: bytOriginalMousePointer = Screen.MousePointer
        Dim intUserResponse As Integer
        Dim strMsg As String
        '
        ' build message
        '
        blnInErrorHandling = True
        Screen.MousePointer = vbDefault
        '
        strMsg = Replace(Replace(Replace(GetTxt("CriticalError"), "$err.number$", lngErrNumber), "$err.description$", strErrDescription), "$procedurename$", gblAProc(UBound(gblAProc)))
        intUserResponse = MsgBox(strMsg, vbCritical + vbAbortRetryIgnore, AppName)
        Call LogEvent(Error, gblAProc(UBound(gblAProc)) & "$" & lngErrNumber & "$" & strErrDescription & "$" & IIf(intUserResponse = vbAbort, "Abort", "Retry"))
        Call errWriteLogFile(strMsg, IIf(intUserResponse = vbAbort, "Abort", "Retry"))
        '
        ' handle option
        '
        Select Case intUserResponse
            Case vbAbort
               GoTo errHandlerEnd
            Case vbIgnore, vbRetry
               errHandler = errNext
               Screen.MousePointer = bytOriginalMousePointer
               blnInErrorHandling = False
        End Select
    End If
    '
    Exit Function
errHandlerEnd:
        Screen.MousePointer = vbDefault
        End
End Function
'
Public Sub errProcStack(nStackAction, cProcName)
   On Error Resume Next
   '
   Dim strMsg As String
   Dim X As Integer
   Dim nFile As Integer
   Dim cFile As String
   '
   ' skip it if toggle is off
   If Not (blnPM_DebugLog) Or blnInErrorHandling Then
      GoTo errProcStackExit
   End If
   '
   ' handle stack action
   Select Case nStackAction
      Case Is = errPush
         ' add new procedure to stack
         gblNProc = gblNProc + 1
         ReDim Preserve gblAProc(gblNProc)
         gblAProc(gblNProc) = cProcName
         '
         If blnPM_DebugLog Then Call LogStack
      Case Is = errPop
         ' remove procedure from stack
         gblNProc = gblNProc - 1
         '
         ReDim Preserve gblAProc(gblNProc)
   End Select
   '
errProcStackExit::
End Sub
'
Public Sub errWriteLogFile(strMessage As String, strUserResponse As String)
    On Error Resume Next
    '
    ' Write error message to log file
    '
    Dim strFile As String
    Dim BNr As Integer
    Dim i As Integer
    '
    BNr = FreeFile()
    strFile = App.Path & "\Errorlog_" + Format(Now, "mmddhhss") + ".txt"
    '
    Open strFile For Output As BNr
    Print #BNr, "*** ERROR REPORT - [" + App.EXEName + "] ***"
    Print #BNr, ""
    Print #BNr, "DATE: " + Format(Now, "ddmmyyyyhhmmss")
    Print #BNr, ""
    Print #BNr, "ERROR MESSAGE:"
    Print #BNr, "**************"
    Print #BNr, strMessage & vbCrLf & "User response: " & strUserResponse
    Print #BNr, ""
    Print #BNr, "ACTIVE FORM LIST:"
    Print #BNr, "*****************"
    Print #BNr, CreateActiveFormList
    'Print #BNr, ""
    Print #BNr, "PROCEDURE STACK DUMP:"
    Print #BNr, "*********************"
        For i = gblNProc To 1 Step -1
            Print #BNr, Trim(CStr(i)) + " - " + gblAProc(i)
        Next i
    Print #BNr, ""
    Print #BNr, "QUESTION STACK DUMP MESSAGE:"
    Print #BNr, "****************************"
    Print #BNr, mySession.QuestionStackDump
    Print #BNr, "*** EOF ***"
    Close BNr
End Sub

'
' DEBUG INFO LOGGER
'
    Private Sub LogStack()
        Call WriteToFile(App.Path & "\Debug\" & CStr(lngSessionID) & "-StackDump.txt", Replace(Space(gblNProc), " ", vbTab) & gblAProc(UBound(gblAProc)))
    End Sub
    '
    Public Sub LogDebug(strLine As String)
        Call WriteToFile(App.Path & "\Debug\" & CStr(lngSessionID) & "-StackDump.txt", Replace(Space(gblNProc + 1), " ", vbTab) & strLine)
    End Sub
    '
    Private Sub WriteToFile(strFile As String, strLine As String)
        Dim BNr: BNr = FreeFile
        '
        Open strFile For Append As #BNr
            Print #BNr, Date & "-" & Time & " - " & strLine
        Close #BNr
    End Sub
    '
    Private Function CreateActiveFormList() As String
        Dim intfrmCount As Integer
        Dim intfrmClose As Integer
        '
        intfrmCount = Forms.Count
        If intfrmCount >= 1 Then
            For intfrmClose = intfrmCount - 1 To 0 Step -1
                If Forms(intfrmClose).Visible Then
                    CreateActiveFormList = CreateActiveFormList & "Loaded & Visible: " & Forms(intfrmClose).Name & vbCrLf
                Else
                    CreateActiveFormList = CreateActiveFormList & "Loaded: " & Forms(intfrmClose).Name & vbCrLf
                End If
            Next
        End If
    End Function
