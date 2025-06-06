VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
Public myScoreBoard                         As New clsScoreBoard
Private myQuestions()                       As clsQuestion
Private sbrSessionInfo                      As StatusBar
'
Private strSessionFile                      As String
Private bytTimeLimit                        As Byte
Private blnTimeLimitMode                    As Boolean
'
' Statistiekvariabelen
Private intCurrentQuestionNumber            As Integer
Private intCurrentQuestionTime              As Integer
Private intCurrentAverageQuestionTime       As Integer
Private dtmCurrentTotalSessionTime          As Date
'
' Session Identifiers
Private Const SID_Verbe                     As String = "Verbe"
Private Const SID_Temps                     As String = "Temps"
Private Const SID_Personne                  As String = "Personne"
Private Const SID_PreviousAnswers           As String = "PreviousAnswers"
Private Const SID_TotalQuestions            As String = "TotalQuestions"
Private Const SID_CurrentQuestioNumber      As String = "QuestionNumber"
Private Const SID_CurrentTotalSessionTime   As String = "TotalSessionTime"
Private Const SID_CurrentQuestionTime       As String = "QuestionTime"
Private Const SID_DateStamp                 As String = "DateStamp"
Private Const SID_TimeStamp                 As String = "TimeStamp"
'
' Session File Parts
Private Const SFP_General                   As String = "General"
Private Const SFP_Question                  As String = "Question_"
Private Const SFP_Score                     As String = "Score"
'
' API Calls
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
'
' AFGEVOERDE IDEEN:
'   Public Sub NextQuestion(clsQuestion as Question)
'   InsertQuestion
'
' SUBS
'
    ' Constructor
    Private Sub Class_Initialize()
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.Class_Initialize"
        '
        ReDim myQuestions(1 To 1)
        intCurrentQuestionNumber = 1
        '
        If Not (blnPM_DUF_Install Or blnPM_Open_DUF Or blnPM_DBEdit Or blnPM_Session_Editor Or blnPM_DUF_Editor) Then
            'Hoofdscherm getoond, statusbar koppelen
            Set sbrSessionInfo = frmMain.sbrInfo
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    ' Destructor
    Private Sub Class_Terminate()
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.Class_Terminate"
        '
        Dim CurrentQuestion As Variant
        '
        For Each CurrentQuestion In myQuestions
            Set CurrentQuestion = Nothing
        Next
        '
        Set myScoreBoard = Nothing
        Set CurrentQuestion = Nothing
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    ' CopySlot
    Public Sub CopySlot()
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.CopySlot"
        '
        'Willekeurige bovengrens zoeken
        Randomize
        '
        Dim intTo As Integer
        intTo = Int(Rnd * 10) + 2 + intCurrentQuestionNumber
        '
        Do
            If QuestionSlotEmpty(intTo) Then
                If UBound(myQuestions) < intTo Then ReDim Preserve myQuestions(LBound(myQuestions) To intTo)
                Set myQuestions(intTo) = myQuestions(intCurrentQuestionNumber)
                Exit Do
            End If
            '
            intTo = intTo + 1
        Loop
        '
        If blnPM_DebugLog Then Call LogDebug("Copied " & intCurrentQuestionNumber & " to " & intTo)
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    ' AddQuestion
    Public Sub AddQuestion(intSlot As Integer, intVerbe As Integer, bytTemps As Byte, bytPersonne As Byte, Optional blnFromSessionFile As Boolean)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.AddQuestion"
        '
        If UBound(myQuestions) < intSlot Then ReDim Preserve myQuestions(LBound(myQuestions) To intSlot)
        '
        If myQuestions(intSlot) Is Nothing Then
            Set myQuestions(intSlot) = New clsQuestion
        End If
        '
        With myQuestions(intSlot)
            .VerbeIndex = intVerbe
            .TempsIndex = bytTemps
            .PersonneIndex = bytPersonne
            .FromSessionFile = blnFromSessionFile
            '
            If blnPM_DebugLog Then Call LogDebug("Question added: " & .Verbe & ", " & .Temps & ", " & .Personne)
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    ' Timer
    Public Sub Timer()
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.Timer"
        '
        Screen.MousePointer = vbArrow
        dtmCurrentTotalSessionTime = DateAdd("s", 1, dtmCurrentTotalSessionTime)
        '
        With frmMain
            'If .mnuFileSettingsTimeLimitStatusOn.Checked Then
            If blnTimeLimitMode Then
                'Werken met tijdslimiet
                If intCurrentQuestionTime = 1 Then
                    'Tijd verstreken
                    .txtInput.Text = IIf(.txtInput.Text = "", GetTxt("TimeUp"), .txtInput.Text & " (" & GetTxt("TimeUp") & ")")
                    Call frmMain.cmdOK_Click
                Else
                    'Nog tijd over
                    intCurrentQuestionTime = intCurrentQuestionTime - 1
                    '
                    Panel(pnlTijd) = CStr(intCurrentQuestionTime & "s")
                    Panel(pnlTotTijd) = CStr(dtmCurrentTotalSessionTime)
                End If
            Else
                'Gewone afloop
                intCurrentQuestionTime = intCurrentQuestionTime + 1
                '
                If .mnuViewStatusbarShowTime.Checked Then
                    Panel(pnlTijd) = CStr(intCurrentQuestionTime & "s")
                    Panel(pnlTotTijd) = CStr(dtmCurrentTotalSessionTime)
                End If
            End If
        End With
        '
        'DEMO: Slechts een kwartier gebruiken
        'Elke minuut controleren
        If Second(dtmCurrentTotalSessionTime) = 0 Then
            If RegisteredVersion = False And Minute(Now - dtmProgramStart) > 14 Then
                MsgBox GetTxt("Registration_SessionExpired"), vbCritical, AppName
                Call ExitProgram
            End If
        End If
        
        '
        GoTo ExitProc
ErrorOccured::
        Resume Next
ExitProc::
        'DoEvents om flikkering tegen te gaan. Verwijder volgende regel voor optimalisatie
        DoEvents
        Screen.MousePointer = vbDefault
        '
        errProcStack errPop, ""
    End Sub
    '
    ' ShowSaveSessionDialog
    Public Sub ShowSaveSessionDialog()
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.ShowSaveSessionDialog"
        '
        With frmCommonDialog.dlgCommonDialog
            .DialogTitle = GetTxt("GiveSessionFileName")
            .Filter = GetTxt("SessionFile") & " (*" & gcSessionFileExtention & ")|*" & gcSessionFileExtention
            .FileName = ""
            .CancelError = True
            .ShowSave
            '
            strSessionFile = .FileName
            '
            If Dir(strSessionFile) = StrReverse(Left(StrReverse(strSessionFile), InStr(StrReverse(strSessionFile), "\") - 1)) Then
                'Dit bestand bestaat
                If MsgBox(GetTxt("WarningFileExists"), vbQuestion + vbYesNo, AppName) = vbNo Then GoTo ExitProc
            End If
            '
            Call SaveSessionFile
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If Err.Number = 32755 Then
            'Op Cancel gedrukt
            GoTo ExitProc
        Else
            If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
        End If
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    ' DeleteQuestion
    Public Sub DeleteQuestion(intIndex As Integer)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.DeleteQuestion"
        '
        Dim i As Integer
        '
        For i = intIndex To UBound(myQuestions) - 1
            Set myQuestions(i) = myQuestions(i + 1)
        Next i
        '
        Set myQuestions(UBound(myQuestions)) = Nothing
        '
        If LBound(myQuestions) <> UBound(myQuestions) Then
            ReDim Preserve myQuestions(LBound(myQuestions) To UBound(myQuestions) - 1)
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    ' ResetQuestion
    Public Sub ResetQuestion(intSlot As Integer)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.ResetQuestion"
        '
        If intSlot <= UBound(myQuestions) Then
            Set myQuestions(intSlot) = Nothing
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
'
' FUNCTIONS
'
    '
    ' ShowOpenSessionDialog
    Public Function ShowOpenSessionDialog() As Boolean
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.ShowOpenSessionDialog"
        '
        If intCurrentQuestionNumber <> 1 Then
            If MsgBox(GetTxt("WarningCurrentSessionLoss"), vbQuestion + vbYesNo, AppName) = vbNo Then GoTo ExitProc
        End If
        '
        With frmCommonDialog.dlgCommonDialog
            .DialogTitle = GetTxt("SelectSessionFile")
            .Filter = GetTxt("SessionFile") & " (*" & gcSessionFileExtention & ")|*" & gcSessionFileExtention
            .FileName = ""
            .CancelError = True
            .ShowOpen
            '
            strSessionFile = .FileName
            '
            If ValidSessionFile Then
                ShowOpenSessionDialog = OpenSessionFile
            Else
                MsgBox GetTxt("InvalidSessionFile"), vbExclamation, AppName
                GoTo ExitProc
            End If
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If Err.Number = 32755 Then
            'Op Cancel gedrukt
            GoTo ExitProc
        Else
            If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
        End If
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    ' OpenSessionFile
    Public Function OpenSessionFile() As Boolean
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.OpenSessionFile"
        '
        'Geeft terug of het openen succesvol was, we nemen aan dat dit een geldig bestand is
        '
        Dim strPreviousAnswers() As String
        Dim intTotalQuestions As Integer
        Dim i As Integer
        '
        Dim intVerbe As Integer
        Dim bytTemps As Byte
        Dim bytPersonne As Byte
        '
        Screen.MousePointer = vbHourglass
        ' Header inladen
        Call ReadHeader(intTotalQuestions, intCurrentQuestionNumber, dtmCurrentTotalSessionTime, intCurrentQuestionTime)
        '
        frmLoad.Show
        frmLoad.lblCurrentJob.Caption = GetTxt("OpeningSession")
        frmLoad.prgDatabase.Max = IIf(intTotalQuestions = 0, 1, intTotalQuestions)
        DoEvents
        '
        ' Werkwoorden inladen
        For i = 1 To intTotalQuestions
            frmLoad.prgDatabase.Value = i
            '
            intVerbe = myFormes.GetIndexNumber((ReadINI(SFP_Question & CStr(i), SID_Verbe, strSessionFile, 0)))
            bytTemps = CByte(ReadINI(SFP_Question & CStr(i), SID_Temps, strSessionFile, 0))
            bytPersonne = CByte(ReadINI(SFP_Question & CStr(i), SID_Personne, strSessionFile, 0))
            '
            If intVerbe <> 0 And bytTemps <> 0 And bytPersonne <> 0 Then
                Call AddQuestion(i, intVerbe, bytTemps, bytPersonne, True)
                '
                strPreviousAnswers = Split(ReadINI(SFP_Question & CStr(i), SID_PreviousAnswers, strSessionFile, ""), ";")
                If UBound(strPreviousAnswers) > -1 Then myQuestions(i).PreviousAnswers = strPreviousAnswers
            Else
                If Not (intVerbe = 0 And bytTemps = 0 And bytPersonne = 0) Then
                    'Dit kan niet !!!
                    OpenSessionFile = False
                    GoTo ExitProc
                End If
            End If
        Next i
        '
        ' Scores inladen
        If blnPM_Session_Editor = False Then
            Dim strInput As String
            '
            myScoreBoard.Reset
            '
            For i = 1 To 22
                strInput = ReadINI(SFP_Score, Convert(AbsNumber, Abreviation, i), strSessionFile, "0/0")
                '
                myScoreBoard.BehSc(i) = CInt(Left(strInput, InStr(strInput, "/") - 1))
                myScoreBoard.MaxSc(i) = CInt(Mid(strInput, InStr(strInput, "/") + 1))
            Next i
            '
            ' Statusbar initialiseren
            Panel(pnlOefNr) = intCurrentQuestionNumber
            '
            If myScoreBoard.TotalNumberOfEntries > 0 Then
                Panel(pnlScore) = myScoreBoard.TotalScore(100) & "%" 'Round(intBehSc / intMaxSc * 100, 0) & "%"
                Panel(pnlGemTijd) = Round((Hour(dtmCurrentTotalSessionTime) * 60 * 60 + Minute(dtmCurrentTotalSessionTime) * 60 + Second(dtmCurrentTotalSessionTime)) / (intCurrentQuestionNumber), 1) & "s"
            Else
                Panel(pnlScore) = ""
                Panel(pnlGemTijd) = ""
            End If
            Panel(pnlTijd) = intCurrentQuestionTime & "s"
            Panel(pnlTotTijd) = CStr(dtmCurrentTotalSessionTime)
            '
            Call PrintQuestion(intCurrentQuestionNumber)
        End If
        '
        OpenSessionFile = True
        '
        Unload frmLoad
        '
        Screen.MousePointer = vbDefault
        '
        GoTo ExitProc
ErrorOccured::
        OpenSessionFile = False
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    ' SaveSessionFile
    Public Function SaveSessionFile()
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.SaveSessionFile"
        '
        'Schrijft de huidige sessie naar een bestand, overschrijft het desnoods
        '
        Dim intSaveStartQuestion As Integer
        '
        frmLoad.Show
        '
        'Laatste sessiebestand verwijderen
        If Dir(strSessionFile) = StrReverse(Left(StrReverse(strSessionFile), InStr(StrReverse(strSessionFile), "\") - 1)) Then Kill strSessionFile
        frmLoad.lblCurrentJob.Caption = GetTxt("SavingSession")
        frmLoad.prgDatabase.Max = IIf(TotalQuestions = 0, 1, TotalQuestions)
        DoEvents
        '
        Dim i As Integer
        '
        Screen.MousePointer = vbHourglass
        '
        '1. Aantal vragen wegschrijven, andere statistieken
        If blnPM_Session_Editor Then
            intCurrentQuestionNumber = 1
            CurrentTotalSessionTime = 0
            intCurrentQuestionTime = 0
        End If
        '
        Call WriteINI(SFP_General, SID_TotalQuestions, TotalQuestions, strSessionFile)
        Call WriteINI(SFP_General, SID_CurrentQuestioNumber, CStr(intCurrentQuestionNumber), strSessionFile)
        Call WriteINI(SFP_General, SID_CurrentTotalSessionTime, CStr(CurrentTotalSessionTime), strSessionFile)
        Call WriteINI(SFP_General, SID_CurrentQuestionTime, CStr(intCurrentQuestionTime), strSessionFile)
        Call WriteINI(SFP_General, SID_DateStamp, CStr(Date), strSessionFile)
        Call WriteINI(SFP_General, SID_TimeStamp, CStr(Time), strSessionFile)
        '
        '2. Vragen wegschrijven
        If (GetDBSetting("SaveAllQuestions", False) = True Or blnPM_Session_Editor) Then
            'Volledig sessiebestand: alle vragen wegschrijven
            intSaveStartQuestion = 1
        Else
            'Incrementeel sessiebestand: enkel nuttige vragen opslaan
            intSaveStartQuestion = CurrentQuestionNumber
        End If
        '
        For i = intSaveStartQuestion To TotalQuestions
            frmLoad.prgDatabase.Value = i
            '
            With myQuestions(i)
                If Not (QuestionSlotEmpty(i)) Then
                    Call WriteINI(SFP_Question & CStr(i), SID_Verbe, myFormes.GetInfinitive(.VerbeIndex), strSessionFile)
                    Call WriteINI(SFP_Question & CStr(i), SID_Temps, CStr(.TempsIndex), strSessionFile)
                    Call WriteINI(SFP_Question & CStr(i), SID_Personne, CStr(.PersonneIndex), strSessionFile)
                    If Join(.PreviousAnswers) <> "" And Not (blnPM_Session_Editor) Then
                        'Als er geen PreviousAnswers zijn, of we bewerken een sessie, niets wegschrijven
                        Call WriteINI(SFP_Question & CStr(i), SID_PreviousAnswers, Join(.PreviousAnswers, ";"), strSessionFile)
                    End If
                End If
            End With
        Next i
        '
        '3. Score wegschrijven
        If Not (blnPM_Session_Editor) Then
            For i = 1 To 22
                If myScoreBoard.NumberOfEntries(i) <> 0 Then
                    Call WriteINI(SFP_Score, Convert(AbsNumber, Abreviation, i), myScoreBoard.Score(i) & "/" & myScoreBoard.NumberOfEntries(i), strSessionFile)
                End If
            Next i
        End If
        '
        Unload frmLoad
        '
        Screen.MousePointer = vbDefault
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    ' CurrentQuestion
    Public Function CurrentQuestion() As clsQuestion
        Set CurrentQuestion = myQuestions(intCurrentQuestionNumber)
    End Function
    '
    ' QuestionAllowed
    Public Function QuestionAllowed(intSlot As Integer) As Boolean
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.QuestionAllowed"
        '
        If intSlot > UBound(myQuestions) Then
            QuestionAllowed = True
        Else
            If myQuestions(intSlot) Is Nothing Then
                QuestionAllowed = True
            Else
                With myQuestions(intSlot)
                    QuestionAllowed = myFormes.Allowed(.VerbeIndex, .TempsIndex, .PersonneIndex)
                End With
            End If
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    ' QuestionSlotEmpty
    Public Function QuestionSlotEmpty(intSlot As Integer) As Boolean
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.QuestionSlotEmpty"
        '
        If intSlot > UBound(myQuestions) Then
            QuestionSlotEmpty = True
        Else
            QuestionSlotEmpty = (myQuestions(intSlot) Is Nothing)
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    ' QuestionAlreadyPosed
    Public Function QuestionAlreadyPosed(intSlot As Integer) As Boolean
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.QuestionAlreadyPosed"
        '
        With myQuestions(intSlot)
            QuestionAlreadyPosed = myFormes.Posed(.VerbeIndex, .TempsIndex, .PersonneIndex)
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    ' ReadHeader
    Public Function ReadHeader(Optional ByRef intTotalQuestions As Integer, Optional ByRef intCurrentQuestionNumber As Integer, Optional ByRef dtmCurrentTotalSessionTime As Date, Optional ByRef intCurrentQuestionTime As Integer, Optional ByRef strDatestamp As String, Optional ByRef strTimestamp As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.ReadHeader"
        '
        intTotalQuestions = CInt(ReadINI(SFP_General, SID_TotalQuestions, strSessionFile))
        intCurrentQuestionNumber = CInt(ReadINI(SFP_General, SID_CurrentQuestioNumber, strSessionFile))
        dtmCurrentTotalSessionTime = CDate(ReadINI(SFP_General, SID_CurrentTotalSessionTime, strSessionFile))
        intCurrentQuestionTime = CInt(ReadINI(SFP_General, SID_CurrentQuestionTime, strSessionFile))
        strDatestamp = ReadINI(SFP_General, SID_DateStamp, strSessionFile)
        strTimestamp = ReadINI(SFP_General, SID_TimeStamp, strSessionFile)
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    ' QuestionStackDump
    Public Function QuestionStackDump() As String
        On Error Resume Next
        '
        QuestionStackDump = "Number of questions: " & UBound(myQuestions) & vbCrLf
        QuestionStackDump = QuestionStackDump & "Current position: " & intCurrentQuestionNumber & vbCrLf & vbCrLf
        '
        Dim i As Integer
        '
        For i = LBound(myQuestions) To UBound(myQuestions)
            If myQuestions(i) Is Nothing Then
                QuestionStackDump = QuestionStackDump & "Question " & CStr(i) & " is nothing" & vbCrLf & vbCrLf
            Else
                With myQuestions(i)
                    QuestionStackDump = QuestionStackDump & "Question " & CStr(i) & vbCrLf
                    QuestionStackDump = QuestionStackDump & vbTab & "Verbe:             " & .Verbe & vbCrLf
                    QuestionStackDump = QuestionStackDump & vbTab & "Temps:             " & .Temps & vbCrLf
                    QuestionStackDump = QuestionStackDump & vbTab & "Personne:          " & .Personne & vbCrLf
                    QuestionStackDump = QuestionStackDump & vbTab & "QuestionText:      " & .QuestionText & vbCrLf
                    QuestionStackDump = QuestionStackDump & vbTab & "CorrectAnswer:     " & .CorrectAnswer & vbCrLf
                    QuestionStackDump = QuestionStackDump & vbTab & "PreviousAnswers:   " & IIf(Join(.PreviousAnswers) <> "", Join(.PreviousAnswers, ";"), "none") & vbCrLf & vbCrLf
                End With
            End If
        Next i
    End Function
    '
    ' ReadINI
    Private Function ReadINI(strSection As String, strKey As String, strDatabaseLocation As String, Optional strDefault As String) As String
        On Error Resume Next
        '
        Dim sData As String: sData = Space(8192)
        Dim lDataLen As Long: lDataLen = GetPrivateProfileString(strSection, strKey, strDefault, sData, Len(sData), strDatabaseLocation)
        '
        ReadINI = Left(sData, lDataLen)
        '
        If blnPM_DebugLog Then Call LogDebug("ReadINI: Section: " & strSection & ", Key: " & strKey & ", Default: " & strDefault & ", Result: " & ReadINI)
    End Function
    '
    ' WriteINI
    Private Function WriteINI(strSection As String, strKey As String, strSetting As String, strDatabaseLocation As String)
        On Error Resume Next
        '
        Call WritePrivateProfileString(strSection, strKey, strSetting, strDatabaseLocation)
        '
        If blnPM_DebugLog Then Call LogDebug("WriteINI: Section: " & strSection & ", Key: " & strKey & ", Setting: " & strSetting)
    End Function
    '
    ' PromptOpenLastSession
    Public Function PromptOpenLastSession() As Boolean
        strSessionFile = App.Path & "\LastSession" & gcSessionFileExtention
        '
        'Controleren of dit een geldig sessiebestand is
        If ValidSessionFile = True And Not (blnPM_NoAskSessionResume) Then
            Dim strDatestamp As String, strTimestamp As String, strResumeQuestion As String
            Call ReadHeader(, , , , strDatestamp, strTimestamp)
            '
            'Vragen of de gebruiker de vorige sessie wilt hervatten
            strResumeQuestion = GetTxt("ResumeSession")
            strResumeQuestion = Replace(strResumeQuestion, "$date$", strDatestamp)
            strResumeQuestion = Replace(strResumeQuestion, "$time$", strTimestamp)
            '
            If MsgBox(strResumeQuestion, vbQuestion + vbYesNo, AppName) = vbYes Then
                'Enkel als het openen succesvol is, naar PromptOpenLastSession op true zetten
                PromptOpenLastSession = OpenSessionFile
            End If
        End If
    End Function
'
' PROPERTIES
'
    ' CurrentQuestionNumber
    Public Property Get CurrentQuestionNumber() As Integer
        CurrentQuestionNumber = intCurrentQuestionNumber
    End Property
    '
    Public Property Let CurrentQuestionNumber(intNewCurrentQuestionNumber As Integer)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.CurrentQuestionNumber [PropertyLet]"
        '
        intCurrentQuestionNumber = intNewCurrentQuestionNumber
        '
        'Statusbar bijwerken
        Panel(pnlOefNr) = intCurrentQuestionNumber
        '
        If myScoreBoard.TotalNumberOfEntries > 0 Then
            Panel(pnlScore) = myScoreBoard.TotalScore(100) & "%"
        End If
        '
        If frmMain.mnuViewStatusbarShowTime.Checked Then
            If myScoreBoard.TotalNumberOfEntries > 0 Then
                intCurrentAverageQuestionTime = Round((Hour(dtmCurrentTotalSessionTime) * 60 * 60 + Minute(dtmCurrentTotalSessionTime) * 60 + Second(dtmCurrentTotalSessionTime)) / (intCurrentQuestionNumber - 1), 1)
                Panel(pnlGemTijd) = intCurrentAverageQuestionTime & "s"
            End If
            '
            If Not (blnTimeLimitMode) Then
                'We werken niet met een tijdslimiet
                Panel(pnlTijd) = "0s"
                intCurrentQuestionTime = 0
            Else
                Panel(pnlTijd) = CStr(bytTimeLimit) & "s"
                intCurrentQuestionTime = CInt(bytTimeLimit)
            End If
        End If
        '
        If blnPM_DebugLog Then Call LogDebug("Current Question Number: " & intCurrentQuestionNumber)
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Property
    '
    ' CurrentQuestionTime
    Public Property Get CurrentQuestionTime() As Integer
        CurrentQuestionTime = intCurrentQuestionTime
    End Property
    '
    Public Property Let CurrentQuestionTime(intNewCurrentQuestionTime As Integer)
        intCurrentQuestionTime = intNewCurrentQuestionTime
    End Property
    '
    ' CurrentAverageQuestionTime
    Public Property Get CurrentAverageQuestionTime() As Integer
        CurrentAverageQuestionTime = intCurrentAverageQuestionTime
    End Property
    '
    ' CurrentTotalSessionTime
    Public Property Get CurrentTotalSessionTime() As Date
        CurrentTotalSessionTime = dtmCurrentTotalSessionTime
    End Property
    '
    Public Property Let CurrentTotalSessionTime(dtmNewCurrentTotalSessionTime As Date)
        dtmCurrentTotalSessionTime = dtmNewCurrentTotalSessionTime
    End Property
    '
    ' TotalQuestions
    Public Property Get TotalQuestions() As Integer
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.TotalQuestions [PropertyGet]"
        '
        If Incremental Then
            'Todo: klopt dit?
            TotalQuestions = UBound(myQuestions)
        Else
            If Not (myQuestions(1) Is Nothing) Then
                TotalQuestions = UBound(myQuestions)
            Else
                TotalQuestions = 0
            End If
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Property
    '
    ' SessionFile
    Public Property Get SessionFile() As String
        SessionFile = strSessionFile
    End Property
    '
    Public Property Let SessionFile(strNewSessionFile As String)
        strSessionFile = strNewSessionFile
        If blnPM_DebugLog Then Call LogDebug("Sessionfile: " & strNewSessionFile)
    End Property
    '
    ' Panel
    Private Property Get Panel(bytPanel As StatusBarPanels) As String
        Panel = sbrSessionInfo.Panels(bytPanel).Text
    End Property
    '
    Private Property Let Panel(bytPanel As StatusBarPanels, strNewValue As String)
        sbrSessionInfo.Panels(bytPanel).Text = strNewValue
    End Property
    '
    ' TimeLimitMode
    Public Property Get TimeLimitMode() As Boolean
        TimeLimitMode = blnTimeLimitMode
    End Property
    '
    Public Property Let TimeLimitMode(blnNewTimeLimitMode As Boolean)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.TimeLimitMode [PropertyLet]"
        '
        blnTimeLimitMode = blnNewTimeLimitMode
        '
        If blnNewTimeLimitMode = True Then
            intCurrentQuestionTime = bytTimeLimit
            '
            If bytTimeLimit = 0 Then
                bytTimeLimit = CByte(GetDBSetting("TimeLimit", 0))
            End If
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Property
    '
    ' TimeLimit
    Public Property Get TimeLimit() As Byte
        TimeLimit = bytTimeLimit
    End Property
    '
    Public Property Let TimeLimit(bytNewTimeLimit As Byte)
        bytTimeLimit = bytNewTimeLimit
    End Property
    '
    ' Stack
    Public Property Get Stack(intIndex As Integer) As clsQuestion
        Set Stack = myQuestions(intIndex)
    End Property
    '
    ' IsIncremental
    Public Property Get Incremental() As Boolean
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.Incremental [PropertyGet]"
        '
        Dim i As Integer
        '
        For i = LBound(myQuestions) To UBound(myQuestions)
            If myQuestions(i) Is Nothing Then
                Incremental = True
                GoTo ExitProc
            End If
        Next i
        '
        Incremental = False
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Property
    '
    ' ValidSessionFile
    Public Property Get ValidSessionFile() As Boolean
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsSession.ValidSessionFile [PropertyGet]"
        '
        If strSessionFile <> "" Then
            If Dir(strSessionFile) = StrReverse(Left(StrReverse(strSessionFile), InStr(StrReverse(strSessionFile), "\") - 1)) Then
                ValidSessionFile = (ReadINI(SFP_General, SID_TotalQuestions, strSessionFile, False) <> False)
            End If
        Else
            ValidSessionFile = False
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Property
    '
    ' EmptySession
    Public Property Get EmptySession() As Boolean
        Dim clsThisQuestion As Variant
        '
        For Each clsThisQuestion In myQuestions()
            If Not (clsThisQuestion Is Nothing) Then
                EmptySession = False
                Exit Property
            End If
        Next
        '
        EmptySession = True
    End Property
