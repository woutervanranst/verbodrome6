Attribute VB_Name = "modQuestions"
Option Explicit
'
' SUBS
'
    Public Sub GenerateQuestion(intSlot As Integer)
        'Deze functie vult het slot intSlot met een nieuwe vraag
        On Error GoTo ErrorOccured
        errProcStack errPush, "modQuestions.GenerateQuestion"
        '
        Dim blnGenerateNewQuestion As Boolean
        '
        If mySession.QuestionSlotEmpty(intSlot) Then
            blnGenerateNewQuestion = True
        Else
            If mySession.CurrentQuestion.FromSessionFile Then
                blnGenerateNewQuestion = False
            ElseIf Not (frmMain.mnuFileSettingsQuestionPosementReposeWrongQuestions.Checked) Or Not (mySession.QuestionAllowed(intSlot)) Then
                blnGenerateNewQuestion = True
            End If
        End If
        '
        If blnGenerateNewQuestion Then
            '1. Het slot is leeg
            '2. Fout beantwoorde vragen mogen opnieuw gesteld worden, of de vraag komt niet uit een sessie
            '3. De vraag is niet meer toegelaten
            '
            ' Eventuele vorige antwoorden resetten
            Call mySession.ResetQuestion(intSlot)
            '
            Dim strProbabilityChain() As String: strProbabilityChain = mySession.myScoreBoard.ProbabilityChain
            '
            Dim intVerbe As Integer
            Dim bytTemps As Byte
            Dim bytPersonne As Byte
            '
            Randomize
            '
            Do
                '1. Welk werkwoord?
                intVerbe = intAllowedVerbes(Int(Rnd * (UBound(intAllowedVerbes) + 1)))
                '
                '2. Welke tijd?
                bytTemps = strProbabilityChain(Int(Rnd * (UBound(strProbabilityChain) + 1)))
                '
                '3. Converteren naar relatieve variabelewaarde (zie lijst)
                Select Case NumberOfPersons(bytTemps)
                    Case 6 'Deze tijd biedt de keuze aan 6 personen
                        bytPersonne = bytAllowedPersonnes(Int(Rnd * (UBound(bytAllowedPersonnes) + 1)))
                    Case 3 'Deze tijd biedt de keuze aan 3 personen
                        bytPersonne = Int(Rnd * 3) + 1
                    Case 2 'Deze tijd biedt de keuze aan 2 personen
                        bytPersonne = Int(Rnd * 2) + 1
                    Case 1 'Deze tijd heeft geen personen
                        bytPersonne = 1
                End Select
                '
                Call mySession.AddQuestion(intSlot, intVerbe, bytTemps, bytPersonne)
                '
                Dim bytLoopCount As Byte
                If bytLoopCount = gcLoopTreshold Then
                    If Not (CheckIfAllowedFormsExist) Then
                        If MsgBox(GetTxt("AllQuestionsAsked"), vbQuestion + vbYesNo, AppName) = vbYes Then
                            Call frmMain.mnuFileSettingsCallSelectionWindow_Click
                        Else
                            frmMain.mnuFileSettingsQuestionPosementOnlyAskOnce.Checked = True
                            Call frmMain.mnuFileSettingsQuestionPosementOnlyAskOnce_Click
                            bytLoopCount = 0
                        End If
                    Else
                        bytLoopCount = 0
                    End If
                Else
                    bytLoopCount = bytLoopCount + 1
                End If
            Loop Until mySession.QuestionAllowed(intSlot) And _
                Not (frmMain.mnuFileSettingsQuestionPosementOnlyAskOnce.Checked And _
                mySession.QuestionAlreadyPosed(intSlot))
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Public Sub PrintQuestion(intSlot As Integer)
        On Error GoTo ErrorOccured
        errProcStack errPush, "modQuestions.PrintQuestion"
        '
        Dim strSplitted() As String
        Dim strQuestionText As String
        Dim i As Integer
        '
        'mySession.CurrentQuestionNumber = intSlot
        strQuestionText = mySession.CurrentQuestion.QuestionText
        '
        If Len(strQuestionText) > 35 Then
            strSplitted = Split(strQuestionText, " ")
            strQuestionText = ""
            '
            For i = LBound(strSplitted) To UBound(strSplitted)
                If Len(strQuestionText & " " & strSplitted(i)) > 35 And InStr(strQuestionText, vbCrLf) = False Then
                    strQuestionText = strQuestionText & vbCrLf & strSplitted(i)
                Else
                    strQuestionText = Trim(strQuestionText & " " & strSplitted(i))
                End If
            Next i
        End If
        '
        frmMain.lblVraag.Caption = strQuestionText
        DoEvents
        Call LogEvent(AskedQuestion, mySession.CurrentQuestion.Verbe & "$" & mySession.CurrentQuestion.TempsIndex & "$" & mySession.CurrentQuestion.PersonneIndex)
        If blnPM_Debug Then frmMain.txtInput.Text = mySession.CurrentQuestion.CorrectAnswer
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Public Sub CheckAnswer()
        On Error GoTo ErrorOccured
        errProcStack errPush, "modQuestions.CheckAnswer"
        '
        Dim strUserAnswer As String: strUserAnswer = Replace(Trim(frmMain.txtInput.Text), "  ", " ")
        'Dim strCorrectAnswer As String: strCorrectAnswer = udtQuestions(intOefNr).Antwoord
        Dim strMsgbox As String
        '
        With mySession.CurrentQuestion
            If strUserAnswer = "" Then
                'Niets ingegeven
                GoTo ExitProc
            ElseIf Left(strUserAnswer, 1) = "/" Then
                'Commando
                If frmMain.mnuFileSettingsAcceptCommands.Checked Then
                    Select Case Mid(strUserAnswer, 2)
                        Case "showquestionstack"
                            frmQuestionStack.Show 1
                        Case "wouterisgrootwouterismachtigenhijheeftnepietvannemetertachtig"
                            blnPM_Debug = Not (blnPM_Debug)
                            If blnPM_Debug Then
                                frmMain.txtInput.Text = mySession.CurrentQuestion.CorrectAnswer
                                GoTo ExitProc
                            End If
                        Case "quit"
                            Call frmMain.mnuFileExit_Click
                        Case "settings"
                            Call frmMain.mnuFileSettingsCallSelectionWindow_Click
                        Case "score"
                            Call frmMain.mnuFileShowScore_Click
                        Case "history"
                            Call frmMain.mnuFileHistoryView_Click
                        Case "info"
                            Call frmMain.mnuInfo_Click
                        Case "nederlands", "english", "français"
                            Select Case Mid(strUserAnswer, 2)
                                Case "nederlands": Call frmMain.mnuFileSettingsLanguage_Click(0)
                                Case "français": Call frmMain.mnuFileSettingsLanguage_Click(1)
                                Case "english": Call frmMain.mnuFileSettingsLanguage_Click(2)
                            End Select
                        Case "lookup"
                            Call frmMain.mnuViewLookup_Click
                        Case "register"
                            If blnPM_Debug Then
                                LetDBSetting "RegistrationKey", UCase(GenerateRegKey)
                                MsgBox GetTxt("RegistrationKey_Correct"), vbInformation, AppName
                            Else
                                GoTo NoCommand
                            End If
                        Case Else
NoCommand::
                            'Command not recognised
                            MsgBox Replace(GetTxt("CommandUnknown"), "$command$", strUserAnswer), vbExclamation, AppName
                    End Select
                    '
                    Call LogEvent(Command, strUserAnswer)
                End If
                '
                If blnPM_Debug Then
                    frmMain.txtInput.Text = mySession.CurrentQuestion.CorrectAnswer
                Else
                    frmMain.txtInput.Text = ""
                End If
                '
                GoTo ExitProc
            ElseIf IsCorrectAnswer(strUserAnswer, .CorrectAnswer) Then
                'Juist antwoord
                If Join(.PreviousAnswers) = "" Then
                    'Van de eerste keer juist
                    strMsgbox = GetTxt("AnswerStatistic_FirstTimeCorrect")
                    Call LogEvent(CorrectAnswer, "1" & "$" & CStr(mySession.CurrentQuestionTime) & "s")
                Else
                    'Meerdere pogingen
                    strMsgbox = GetTxt("AnswerStatistic_NotFirstTimeCorrect")
                    strMsgbox = Replace(strMsgbox, "$tries$", UBound(.PreviousAnswers) + 2)
                    strMsgbox = Replace(strMsgbox, "$previousanswers$", Join(.PreviousAnswers, vbCrLf))
                    Call LogEvent(CorrectAnswer, CStr(UBound(.PreviousAnswers) + 2) & "$" & CStr(mySession.CurrentQuestionTime) & "s" & "$" & Join(.PreviousAnswers, "$"))
                End If
                '
                'Score bijwerken
                Call mySession.myScoreBoard.AddEntry(True, UBound(.PreviousAnswers) + 2, .TempsIndex)
                '
                'Aangeven dat dit gevraagd is (en goed beantwoord)
                myFormes.Posed(.VerbeIndex, .TempsIndex, .PersonneIndex) = True
                '
                'Melding weergeven
                MsgBox strMsgbox, vbInformation, AppName
            Else
                'Fout antwoord
                If Join(.PreviousAnswers) = "" Then
                    'De eerste keer fout
                    strMsgbox = GetTxt("AnswerStatistic_FirstTimeWrong")
                    strMsgbox = Replace(strMsgbox, "$correctanswer$", .CorrectAnswer)
                    strMsgbox = Replace(strMsgbox, "$wronganswer$", strUserAnswer)
                    Call LogEvent(WrongAnswer, "1" & "$" & CStr(mySession.CurrentQuestionTime) & "s" & "$" & strUserAnswer)
                Else
                    'Meerdere pogingen
                    'Is dit antwoord al eens gegeven?
                    strMsgbox = GetTxt("AnswerStatistic_NotFirstTimeWrong")
                    If Join(Filter(.PreviousAnswers, strUserAnswer)) <> "" Then
                        strMsgbox = Replace(strMsgbox, "$alreadygiven$", GetTxt("AnswerStatistic_AlreadyGiven"))
                    Else
                        strMsgbox = Replace(strMsgbox, "$alreadygiven$", "")
                    End If
                    strMsgbox = Replace(strMsgbox, "$tries$", UBound(.PreviousAnswers) + 1)
                    strMsgbox = Replace(strMsgbox, "$correctanswer$", .CorrectAnswer)
                    strMsgbox = Replace(strMsgbox, "$wronganswer$", strUserAnswer)
                    strMsgbox = Replace(strMsgbox, "$previousanswers$", Join(.PreviousAnswers, vbCrLf))
                    Call LogEvent(WrongAnswer, CStr(UBound(.PreviousAnswers) + 2) & "$" & CStr(mySession.CurrentQuestionTime) & "s" & "$" & Join(.PreviousAnswers, "$"))
                End If
                '
                'Vraag opnieuw stellen?
                If frmMain.mnuFileSettingsQuestionPosementReposeWrongQuestions.Checked Then
                    'De slot kopieeren, en het foute antwoord stockeren
                    Call mySession.CopySlot
                    Call .AddPreviousAnswer(strUserAnswer)
                End If
                '
                'Score bijwerken
                Call mySession.myScoreBoard.AddEntry(False, UBound(.PreviousAnswers) + 1, .TempsIndex)
                '
                'Melding weergeven
                MsgBox strMsgbox, vbExclamation, AppName
            End If
        End With
        '
        'Volgende vraag
        frmMain.txtInput.Text = ""
        frmMain.txtInput.SetFocus
        mySession.CurrentQuestionNumber = mySession.CurrentQuestionNumber + 1
        Call GenerateQuestion(mySession.CurrentQuestionNumber)
        Call PrintQuestion(mySession.CurrentQuestionNumber)
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
'
' PROCEDURE HULP FUNCTIES
'
    Private Function CheckIfAllowedFormsExist() As Boolean
        'TODO: Deze functie vervangen (vééél te traag)
'        On Error Resume Next
'        '
'        Dim i As Integer, j As Byte, k As Byte
'        '
'        Screen.MousePointer = vbHourglass
'        '
'        For i = LBound(udtVerbes) To UBound(udtVerbes)
'            For j = 1 To 17
'                For k = 1 To 6
'                    'If udtVerbes(i, j, k).Toegelaten Then
'                        If frmMain.mnuFileSettingsQuestionPosementOnlyAskOnce.Checked And Not (myFormes.Posed(i, j, k)) Then ' udtVerbes(i, j, k).Gevraagd = False Then
'                            CheckIfAllowedFormsExist = True
'                            Exit Function
'                        ElseIf Not (frmMain.mnuFileSettingsQuestionPosementOnlyAskOnce.Checked) Then
'                            CheckIfAllowedFormsExist = True
'                            Exit Function
'                        End If
'                    'End If
'                Next k
'            Next j
'        Next i
'        '
'        Screen.MousePointer = vbDefault
    End Function
    '
    Private Function IsCorrectAnswer(strUserAnswer As String, strCorrectAnswer As String) As Boolean
        On Error GoTo ErrorOccured
        errProcStack errPush, "modQuestions.IsCorrectAnswer"
        '
        'Verandering v6.4:
        strUserAnswer = LCase(strUserAnswer)
        '
        If InStr(strCorrectAnswer, "/") Then
            'Meerdere mogelijkheden
            Dim strSingleEntities() As String, strMultipleEntities() As String
            '
            If InStr(strCorrectAnswer, " ") Then
                strSingleEntities = Split(strCorrectAnswer, " ")
                strMultipleEntities = Split(strSingleEntities(UBound(strSingleEntities)), "/")
                strSingleEntities(UBound(strSingleEntities)) = ""
            Else
                strMultipleEntities = Split(strCorrectAnswer, "/")
            End If
            '
            If Join(strSingleEntities, " ") & strMultipleEntities(0) = strUserAnswer Or _
               Join(strSingleEntities, " ") & strMultipleEntities(1) = strUserAnswer Then
                IsCorrectAnswer = True
            Else
                IsCorrectAnswer = False
            End If
        Else
            'Slechts 1 mogelijkheid
            IsCorrectAnswer = (LCase(strUserAnswer) = LCase(strCorrectAnswer))
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
