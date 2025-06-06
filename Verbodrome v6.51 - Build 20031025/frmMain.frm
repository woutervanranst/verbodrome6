VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMain"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrActiveForm 
      Interval        =   1000
      Left            =   5400
      Top             =   1200
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   5775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2085
            MinWidth        =   2085
            Text            =   "1"
            TextSave        =   "1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2085
            MinWidth        =   2085
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2085
            MinWidth        =   2085
            Text            =   "0s"
            TextSave        =   "0s"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2085
            MinWidth        =   2085
            Text            =   "00:00:00"
            TextSave        =   "00:00:00"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2085
            MinWidth        =   2085
         EndProperty
      EndProperty
   End
   Begin VB.Line linStatusbar 
      X1              =   0
      X2              =   6000
      Y1              =   1250
      Y2              =   1250
   End
   Begin VB.Label lblVraag 
      BackStyle       =   0  'Transparent
      Caption         =   "lblVraag"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Score"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Tijd"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   4
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Totale Tijd"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Gem. Tijd"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   4
      Left            =   4800
      TabIndex        =   2
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Oef n°"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Menu mnuFile 
      Caption         =   "mnuFile"
      Begin VB.Menu mnuFileOpenSession 
         Caption         =   "mnuFileOpenSession"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveSession 
         Caption         =   "mnuFileSaveSession"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuStreep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSettings 
         Caption         =   "mnuFileSettings"
         Begin VB.Menu mnuFileSettingsCallSelectionWindow 
            Caption         =   "mnuFileSettingsCallSelectionWindow"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuStreep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileSettingsSetLanguage 
            Caption         =   "mnuFileSettingsSetLanguage"
            Begin VB.Menu mnuFileSettingsLanguage 
               Caption         =   "<<DUTCH>>"
               Index           =   0
            End
            Begin VB.Menu mnuFileSettingsLanguage 
               Caption         =   "<<FRENCH>>"
               Index           =   1
            End
            Begin VB.Menu mnuFileSettingsLanguage 
               Caption         =   "<<ENGLISH>>"
               Index           =   2
            End
         End
         Begin VB.Menu mnuFileSettingsQuestionPosement 
            Caption         =   "mnuFileSettingsQuestionPosement"
            Begin VB.Menu mnuFileSettingsQuestionPosementReposeWrongQuestions 
               Caption         =   "mnuFileSettingsQuestionPosementReposeWrongQuestions"
            End
            Begin VB.Menu mnuFileSettingsQuestionPosementOnlyAskOnce 
               Caption         =   "mnuFileSettingsQuestionPosementOnlyAskOnce"
            End
         End
         Begin VB.Menu mnuFileSettingsReset 
            Caption         =   "mnuFileSettingsReset"
            Begin VB.Menu mnuFileSettingsResetAskedForms 
               Caption         =   "mnuFileSettingsResetAskedForms"
            End
            Begin VB.Menu mnuFileSettingsResetPoints 
               Caption         =   "mnuFileSettingsResetPoints"
            End
         End
         Begin VB.Menu mnuFileSettingsTimeLimit 
            Caption         =   "mnuFileSettingsTimeLimit"
            Begin VB.Menu mnuFileSettingsTimeLimitSet 
               Caption         =   "mnuFileSettingsTimeLimitSet"
            End
            Begin VB.Menu mnuFileSettingsTimeLimitStatus 
               Caption         =   "mnuFileSettingsTimeLimitStatus"
               Begin VB.Menu mnuFileSettingsTimeLimitStatusSet 
                  Caption         =   "mnuFileSettingsTimeLimitStatusSet(0)"
                  Index           =   0
               End
               Begin VB.Menu mnuFileSettingsTimeLimitStatusSet 
                  Caption         =   "mnuFileSettingsTimeLimitStatusSet(1)"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu mnuFileSettingsSessions 
            Caption         =   "mnuFileSettingsSessions"
            Begin VB.Menu mnuFileSettingsSessionsSaveAllQuestions 
               Caption         =   "mnuFileSettingsSessionsSaveAllQuestions"
            End
         End
         Begin VB.Menu mnuStreep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileSettingsAcceptCommands 
            Caption         =   "mnuFileSettingsAcceptCommands"
         End
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "mnuFileHistory"
         Begin VB.Menu mnuFileHistoryView 
            Caption         =   "mnuFileHistoryView"
         End
         Begin VB.Menu mnuStreep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileHistoryEnabled 
            Caption         =   "mnuFileHistoryEnabled"
         End
         Begin VB.Menu mnuFileHistoryEmpty 
            Caption         =   "mnuFileHistoryEmpty"
         End
      End
      Begin VB.Menu mnuStreep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileShowScore 
         Caption         =   "mnuFileShowScore"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFilePause 
         Caption         =   "mnuFilePause"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "mnuFileExit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "mnuView"
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "mnuViewStatusbar"
         Begin VB.Menu mnuViewStatusbarShow 
            Caption         =   "mnuViewStatusbarShow"
         End
         Begin VB.Menu mnuViewStatusbarShowTime 
            Caption         =   "mnuViewStatusbarShowTime"
         End
      End
      Begin VB.Menu mnuViewBackground 
         Caption         =   "mnuViewBackground"
         Begin VB.Menu mnuViewHideBackGround 
            Caption         =   "mnuViewHideBackGround"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuStreep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewBackgroundType 
            Caption         =   "mnuViewBackgroundType"
            Index           =   0
         End
         Begin VB.Menu mnuViewBackgroundType 
            Caption         =   "mnuViewBackgroundType"
            Index           =   1
         End
      End
      Begin VB.Menu mnuStreep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLookup 
         Caption         =   "mnuViewLookup"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "mnuInfo"
   End
   Begin VB.Menu mnuAccents 
      Caption         =   "mnuAccents"
      Begin VB.Menu mnuAccentsAccent 
         Caption         =   "é"
         Index           =   0
      End
      Begin VB.Menu mnuAccentsAccent 
         Caption         =   "è"
         Index           =   1
      End
      Begin VB.Menu mnuAccentsAccent 
         Caption         =   "ê"
         Index           =   2
      End
      Begin VB.Menu mnuAccentsAccent 
         Caption         =   "â"
         Index           =   3
      End
      Begin VB.Menu mnuAccentsAccent 
         Caption         =   "ô"
         Index           =   4
      End
      Begin VB.Menu mnuAccentsAccent 
         Caption         =   "î"
         Index           =   5
      End
      Begin VB.Menu mnuAccentsAccent 
         Caption         =   "ï"
         Index           =   6
      End
   End
   Begin VB.Menu mnuTimeLimitStop 
      Caption         =   "mnuTimeLimitStop"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' FORM
'
    Public Sub cmdOK_Click()
        If blnPM_DebugLog Then Call LogDebug("User input: " & Me.txtInput.Text)
        Call CheckAnswer
    End Sub
    '
    Private Sub Form_Activate()
        txtInput.SetFocus
    End Sub
    '
    Private Sub Form_GotFocus()
        Call Form_Activate
    End Sub
    '
    Private Sub Form_Load()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmMain.Form_Load"
        '
        Dim strResumeQuestion As String
        '
        ' Taalelementen inladen
        Me.Icon = LoadResPicture(107, vbResIcon) 'frmLoad.Icon
        Call LoadLanguage
        '
        'Instellingen inladen
        mnuAccents.Visible = False
        mnuTimeLimitStop.Visible = False
        mnuFileSettingsAcceptCommands.Checked = GetDBSetting("AcceptCommands", True)
        mnuFileSettingsLanguage(GetDBSetting("Language", 0)).Checked = True
        mnuFileSettingsQuestionPosementReposeWrongQuestions.Checked = GetDBSetting("ReposeWrongQuestions", True)
        mnuFileSettingsQuestionPosementOnlyAskOnce.Checked = GetDBSetting("OnlyAskOnce", True)
        mnuFileHistoryEnabled.Checked = (GetDBSetting("HistoryStatus", "Enabled") = "Enabled")
        mnuFileSettingsTimeLimitStatusSet(0).Checked = (GetDBSetting("TimeLimitStatus", "Enabled") = "Enabled")
        mnuFileSettingsTimeLimitStatusSet(1).Checked = (GetDBSetting("TimeLimitStatus", "Enabled") = "Disabled")
        mnuFileSettingsSessionsSaveAllQuestions.Checked = (GetDBSetting("SaveAllQuestions", False) = True)
        mnuViewBackgroundType(0).Checked = (GetDBSetting("BackgroundType", "Wallpaper") = "Wallpaper")
        mnuViewBackgroundType(1).Checked = (GetDBSetting("BackgroundType", "Black") = "Black")
        mnuViewHideBackGround.Checked = (GetDBSetting("BlankBackground", "True") = "True")
        '
        mnuViewStatusbarShow.Checked = True
        If GetDBSetting("ShowStatusBar", True) = False Then Call mnuViewStatusbarShow_Click
        mnuViewStatusbarShowTime.Checked = True
        If GetDBSetting("ShowStatusBarTime", True) = False Then Call mnuViewStatusbarShowTime_Click
        If GetDBSetting("TimeLimitStatus") = "Enabled" Then Call mnuFileSettingsTimeLimitStatusSet_Click(0)
        '
        If mySession.EmptySession Then
            Call GenerateQuestion(mySession.CurrentQuestionNumber)
            Call PrintQuestion(mySession.CurrentQuestionNumber)
        End If
        '
        'Trial
        If RegisteredVersion = False And (GetFirstUseDate("0") = "0") Then
            Call mnuInfo_Click
        End If
        '
        'Geen achtergrond-instelling controleren
        If mnuViewHideBackGround.Checked Then Call frmBlank.Show
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub Form_Unload(Cancel As Integer)
        If MsgBox(GetTxt("QuitProgram"), vbQuestion + vbYesNo, AppName) = vbYes Then
            'Procedure aanroepen om het programma af te sluiten
            Call ExitProgram
        Else
            Cancel = vbCancel
        End If
    End Sub
    '
    Private Sub tmrActiveForm_Timer()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmMain.tmrActiveForm_Timer"
        '
        If frmMain.Visible Then
            If Screen.ActiveForm.Caption = frmMain.Caption Then
                Call mySession.Timer
            End If
        End If
        '
        GoTo ExitProc
ErrorOccured:: 'Om de een of andere mysterieuze reden is frmMain soms geen geldig object, dit veroorzaakt dan een fout. We negeren dit
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    ' Snellere manier om de stack te tonen
    Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim blnShiftDown As Boolean: blnShiftDown = (Shift And vbShiftMask) > 0
        Dim blnCtrlDown As Boolean: blnCtrlDown = (Shift And vbCtrlMask) > 0
        '
        If KeyCode = vbKeyReturn And blnShiftDown And blnCtrlDown And blnPM_Debug Then
            'Stack tonen
            Me.txtInput.Text = "/showquestionstack"
            Call cmdOK_Click
        End If
    End Sub
    '
    Private Sub txtInput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbRightButton Then Call PopupMenu(mnuAccents)
        If blnPM_DebugLog Then Call LogDebug("Accent menu called")
    End Sub
'
' MENUS
'
    Public Sub mnuFileHistoryEnabled_Click()
        mnuFileHistoryEnabled.Checked = Not (mnuFileHistoryEnabled.Checked)
        '
        blnHistoryEnabled = mnuFileHistoryEnabled.Checked
        '
        LetDBSetting "HistoryStatus", IIf(mnuFileHistoryEnabled.Checked, "Enabled", "Disabled")
    End Sub
    '
    Public Sub mnuFileOpenSession_Click()
        Call mySession.ShowOpenSessionDialog
    End Sub
    '
    Public Sub mnuFileSaveSession_Click()
        Call mySession.ShowSaveSessionDialog
    End Sub
    '
    Private Sub mnuFilePause_Click()
        If mnuViewHideBackGround.Checked Then frmBlank.Visible = False
        If blnPM_DebugLog Then Call LogDebug("Pausing")
        '
        With Me
            .Visible = False
            .tmrActiveForm.Enabled = False
            '.tmrSession.Enabled = False
            DoEvents
            '
            MsgBox GetTxt("Pauseer"), vbInformation, AppName
            '
            .tmrActiveForm.Enabled = True
            '.tmrSession.Enabled = True
            .Visible = True
        End With
        'Eventueel het zwart scherm herstellen
        If mnuViewHideBackGround.Checked Then frmBlank.Visible = True
    End Sub
    '
    Public Sub mnuFileSettingsAcceptCommands_Click()
        'De gebruiker veranderde de instelling: Commando's accepteren in de invoerbox
        LetDBSetting "AcceptCommands", Not (mnuFileSettingsAcceptCommands.Checked)
        mnuFileSettingsAcceptCommands.Checked = Not (mnuFileSettingsAcceptCommands.Checked)
    End Sub
    '
    Public Sub mnuFileSettingsCallSelectionWindow_Click()
        Call LogEvent(SelectionWindow)
        '
        frmSelection.Show 1
        '
        'Nieuwe instellingen doorvoeren
        DoEvents
        ReDim intAllowedVerbes(0): intAllowedVerbes(0) = -1
        Call LoadPermissions
        Call GenerateQuestion(mySession.CurrentQuestionNumber)
        Call PrintQuestion(mySession.CurrentQuestionNumber)
    End Sub
    '
    Public Sub mnuFileHistoryEmpty_Click()
        'Historiek ledigen
        If MsgBox(GetTxt("EmptyHistory"), vbYesNo + vbQuestion, AppName) = vbYes Then
            Call EmptyHistory
        End If
    End Sub
    '
    Public Sub mnuFileHistoryView_Click()
        frmHistory.Show 1
    End Sub
    '
    Public Sub mnuFileSettingsLanguage_Click(Index As Integer)
        If bytLanguage <> Index Then
            Dim intSlot As Integer
            '
            'De gebruiker stelde een andere taal in
            mnuFileSettingsLanguage(0).Checked = (Index = 0)
            mnuFileSettingsLanguage(1).Checked = (Index = 1)
            mnuFileSettingsLanguage(2).Checked = (Index = 2)
            '
            bytLanguage = Index
            LetDBSetting "Language", CStr(Index)
            '
            Call LoadLanguage
            '
            lblVraag.Caption = mySession.CurrentQuestion.QuestionText
        End If
    End Sub
    '
    Public Sub mnuFileSettingsQuestionPosementOnlyAskOnce_Click()
        'De gebruiker veranderde de instelling: Goed beantwoorde vragen niet meer stellen
        LetDBSetting "OnlyAskOnce", Not (mnuFileSettingsQuestionPosementOnlyAskOnce.Checked)
        mnuFileSettingsQuestionPosementOnlyAskOnce.Checked = Not (mnuFileSettingsQuestionPosementOnlyAskOnce.Checked)
    End Sub
    '
    Public Sub mnuFileSettingsQuestionPosementReposeWrongQuestions_Click()
        'De gebruiker veranderde de instelling: Fout beantwoorde vragen opnieuw stellen
        LetDBSetting "ReposeWrongQuestions", Not (mnuFileSettingsQuestionPosementReposeWrongQuestions.Checked)
        mnuFileSettingsQuestionPosementReposeWrongQuestions.Checked = Not (mnuFileSettingsQuestionPosementReposeWrongQuestions.Checked)
    End Sub
    
    Public Sub mnuFileSettingsResetAskedForms_Click()
        Call LogEvent(ResetAskedForms)
        If blnPM_DebugLog Then Call LogDebug("Asked forms reset")
        '
        If MsgBox(GetTxt("ResetAskedForms"), vbQuestion + vbYesNo, AppName) = vbYes Then
            Call myFormes.ResetPosedDB
        End If
    End Sub
    '
    Public Sub mnuFileSettingsResetPoints_Click()
        'Puntentelling resetten (niet van deze sessie)
        If MsgBox(GetTxt("ScoreBoardReset"), vbQuestion + vbYesNo, AppName) = vbYes Then
            Dim i As Byte
            '
            strSQL = "SELECT * FROM Score"
            recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
            '
            With recVerbodrome
                '.MoveFirst
                While Not .EOF
                    !BehSc = 0
                    !MaxSc = 0
                    .Update
                    .MoveNext
                Wend
                .Close
            End With
            '
            If blnPM_DebugLog Then Call LogDebug("Scoreboard resetted")
        End If
    End Sub
    '
    Public Sub mnuFileSettingsTimeLimitSet_Click()
        'Tijdslimiet instellen
        Dim strInput As String
        '
        If mySession.TimeLimit = 0 Then
            mySession.TimeLimit = mySession.CurrentAverageQuestionTime
            If mySession.TimeLimit = 0 Then
                mySession.TimeLimit = 10
            End If
        End If
        '
        strInput = InputBox(GetTxt("SetTimeLimit"), AppName, mySession.TimeLimit)
        '
        If IsNumeric(strInput) And Val(strInput) < 256 And Val(strInput) > 3 Then
            mySession.TimeLimit = Val(strInput)
            LetDBSetting "TimeLimit", Val(strInput)
        Else
            MsgBox GetTxt("ErrorSettingTimeLimit"), vbExclamation, AppName
            Exit Sub
        End If
    End Sub
    '
    Private Sub mnuFileSettingsTimeLimitStatusSet_Click(Index As Integer)
        'De gebruiker veranderde de instelling: Tijdslimiet aan/uit
        mnuFileSettingsTimeLimitStatusSet(0).Checked = (Index = 0)
        mnuFileSettingsTimeLimitStatusSet(1).Checked = (Index = 1)
        '
        mnuViewStatusbarShow.Enabled = (Index <> 0)
        mnuViewStatusbarShowTime.Enabled = (Index <> 0)
        mnuTimeLimitStop.Visible = (Index = 0)
        mySession.TimeLimitMode = (Index = 0)
        '
        If Index = 0 Then
            'Tijdslimiet aan
            mnuViewStatusbarShow.Checked = True
            mnuViewStatusbarShowTime.Checked = True
            '
            sbrInfo.Visible = True
            linStatusbar.Visible = True
            Me.Height = 2595
            '
            Call LogEvent(TimeLimitOn, CStr(mySession.TimeLimit))
        Else
            mySession.CurrentQuestionTime = mySession.TimeLimit - mySession.CurrentQuestionTime
            '
            If Not CBool(GetDBSetting("ShowStatusBar", True)) Then Call mnuViewStatusbarShow_Click
            If Not CBool(GetDBSetting("ShowStatusBarTime", True)) Then Call mnuViewStatusbarShowTime_Click
            '
            Call LogEvent(TimeLimitOff)
        End If
        '
        LetDBSetting "TimeLimitStatus", IIf(Index = 0, "Enabled", "Disabled")
    End Sub
    '
    Private Sub mnuFileSettingsSessionsSaveAllQuestions_Click()
        'De gebruiker veranderde de instelling: Alle vragen opslaan in sessiebestand
        LetDBSetting "SaveAllQuestions", Not (mnuFileSettingsSessionsSaveAllQuestions.Checked)
        mnuFileSettingsSessionsSaveAllQuestions.Checked = Not (mnuFileSettingsSessionsSaveAllQuestions.Checked)
    End Sub
    '
    Public Sub mnuFileShowScore_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmMain.mnuFileShowScore_Click"
        '
        If mySession.CurrentQuestionNumber <> 1 Then
            frmScoreReport.Show 1
        Else
            MsgBox GetTxt("Score_NoAvailable"), vbInformation, AppName
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Public Sub mnuInfo_Click()
        frmAbout.Show 1
    End Sub
    '
    Private Sub mnuTimeLimitStop_Click()
        Call mnuFileSettingsTimeLimitStatusSet_Click(1)
    End Sub
    '
    Private Sub mnuViewBackgroundType_Click(Index As Integer)
        If mnuViewBackgroundType(Index).Checked = False Then
            'De gebruiker veranderde de instelling: Achtergrondstype (Black/Wallpaper)
            LetDBSetting "BackgroundType", IIf(Index = 0, "Wallpaper", "Black")
            mnuViewBackgroundType(0).Checked = Not (mnuViewBackgroundType(0).Checked)
            mnuViewBackgroundType(1).Checked = Not (mnuViewBackgroundType(1).Checked)
            '
            If mnuViewHideBackGround.Checked Then
                Call frmBlank.InitWallpaper
            End If
        End If
    End Sub
    '
    Private Sub mnuViewHideBackGround_Click()
        'De gebruiker veranderde de instelling: Geen achtergrond tonen
        LetDBSetting "BlankBackground", Not (mnuViewHideBackGround.Checked)
        mnuViewHideBackGround.Checked = Not (mnuViewHideBackGround.Checked)
        '
        If mnuViewHideBackGround.Checked Then
            'Achtergrond wegdoen
            frmBlank.Show
        Else
            'Achtergrond terugplaatsen
            Unload frmBlank
        End If
    End Sub
    '
    Public Sub mnuViewLookup_Click()
        frmLookup.Show 1
    End Sub
    '
    Public Sub mnuViewStatusbarShow_Click()
        'De gebruiker veranderde de instelling: Statusbar tonen
        LetDBSetting "ShowStatusBar", Not (mnuViewStatusbarShow.Checked)
        sbrInfo.Visible = Not (mnuViewStatusbarShow.Checked)
        linStatusbar.Visible = Not (mnuViewStatusbarShow.Checked)
        Me.Height = IIf(mnuViewStatusbarShow.Checked, 1950, 2595)
        mnuViewStatusbarShow.Checked = Not (mnuViewStatusbarShow.Checked)
    End Sub
    '
    Public Sub mnuViewStatusbarShowTime_Click()
        'De gebruiker veranderde de instelling: Tijd tonen op de statusbar
        LetDBSetting "ShowStatusBarTime", Not (mnuViewStatusbarShowTime.Checked)
        sbrInfo.Panels(pnlTijd).Text = ""
        sbrInfo.Panels(pnlTotTijd).Text = ""
        sbrInfo.Panels(pnlGemTijd).Text = ""
        mnuViewStatusbarShowTime.Checked = Not (mnuViewStatusbarShowTime.Checked)
    End Sub
    '
    Private Sub mnuAccentsAccent_Click(Index As Integer)
        Dim intCurrentPosition As Integer: intCurrentPosition = txtInput.SelStart
        '
        txtInput.Text = Left(txtInput.Text, txtInput.SelStart) & mnuAccentsAccent(Index).Caption & Mid(txtInput.Text, txtInput.SelStart + 1)
        '
        txtInput.SelStart = intCurrentPosition + 1
        '
        If blnPM_DebugLog Then Call LogDebug("Accent inserted " & mnuAccentsAccent(Index).Caption)
    End Sub
    '
    Public Sub mnuFileExit_Click()
        'Programma afsluiten
        Unload frmMain
    End Sub
'
' PROCEDURE HULP FUNCTIES
'
    Public Sub LoadLanguage()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmMain.LoadLanguage"
        '
        'Procedure die de taalelementen van frmMain inlaad(t)
        Dim strArray() As String
        Dim i As Integer
        '
        Me.Caption = AppName & IIf(RegisteredVersion, "", " - " & GetTxt("frmMain_Evaluation"))
        mnuFile.Caption = GetTxt("mnuFile")
        mnuFileOpenSession.Caption = GetTxt("mnuFileOpen")
        mnuFileSaveSession.Caption = GetTxt("mnuFileSave")
        mnuFilePause.Caption = GetTxt("mnuFilePause")
        mnuFileSettings.Caption = GetTxt("mnuFileSettings")
        mnuFileSettingsCallSelectionWindow.Caption = GetTxt("mnuFileSettingsCallSelectionWindow")
        mnuFileSettingsAcceptCommands.Caption = GetTxt("mnuFileSettingsAcceptCommands")
        mnuFileSettingsSetLanguage.Caption = GetTxt("mnuFileSettingsSetLanguage")
        mnuFileSettingsLanguage(0).Caption = GetTxt("mnuFileSettingsLanguage(0)")
        mnuFileSettingsLanguage(1).Caption = GetTxt("mnuFileSettingsLanguage(1)")
        mnuFileSettingsLanguage(2).Caption = GetTxt("mnuFileSettingsLanguage(2)")
        mnuFileSettingsQuestionPosement.Caption = GetTxt("mnuFileSettingsQuestionPosement")
        mnuFileSettingsQuestionPosementReposeWrongQuestions.Caption = GetTxt("mnuFileSettingsQuestionPosementReposeWrongQuestions")
        mnuFileSettingsQuestionPosementOnlyAskOnce.Caption = GetTxt("mnuFileSettingsQuestionPosementOnlyAskOnce")
        mnuFileSettingsReset.Caption = GetTxt("mnuFileSettingsReset")
        mnuFileSettingsResetAskedForms.Caption = GetTxt("mnuFileSettingsResetAskedForms")
        mnuFileSettingsResetPoints.Caption = GetTxt("mnuFileSettingsResetPoints")
        mnuFileSettingsTimeLimit.Caption = GetTxt("mnuFileSettingsTimeLimit")
        mnuFileSettingsTimeLimitSet.Caption = GetTxt("mnuFileSettingsTimeLimitSet")
        mnuFileSettingsTimeLimitStatus.Caption = GetTxt("mnuStatus")
        mnuFileSettingsTimeLimitStatusSet(0).Caption = GetTxt("mnuStatusOn")
        mnuFileSettingsTimeLimitStatusSet(1).Caption = GetTxt("mnuStatusOff")
        mnuFileSettingsSessions.Caption = GetTxt("mnuFileSettingsSessions")
        mnuFileSettingsSessionsSaveAllQuestions.Caption = GetTxt("mnuFileSettingsSessionsSaveAllQuestions")
        mnuFileHistory.Caption = GetTxt("mnuFileHistory")
        mnuFileHistoryView.Caption = GetTxt("mnuFileHistoryView")
        mnuFileHistoryEnabled.Caption = GetTxt("mnuFileHistoryEnabled")
        mnuFileHistoryView.Caption = GetTxt("mnuFileHistoryView")
        mnuFileHistoryEmpty.Caption = GetTxt("mnuFileHistoryEmpty")
        mnuFileShowScore.Caption = GetTxt("mnuFileShowScore")
        mnuFileExit.Caption = GetTxt("mnuFileExit")
        mnuView.Caption = GetTxt("mnuView")
        mnuViewStatusbar.Caption = GetTxt("mnuViewStatusbar")
        mnuViewBackground.Caption = GetTxt("mnuViewBackground")
        strArray = Split(GetTxt("mnuViewBackgroundType"), ";")
        mnuViewBackgroundType(0).Caption = strArray(0)
        mnuViewBackgroundType(1).Caption = strArray(1)
        mnuViewHideBackGround.Caption = GetTxt("mnuViewHideBackGround")
        mnuViewStatusbarShow.Caption = GetTxt("mnuViewStatusbarShow")
        mnuViewStatusbarShowTime.Caption = GetTxt("mnuViewStatusbarShowTime")
        mnuViewLookup.Caption = GetTxt("mnuViewLookup")
        mnuInfo.Caption = GetTxt("mnuInfo")
        mnuTimeLimitStop.Caption = GetTxt("mnuTimeLimitStop")
        '
        strArray = Split(GetTxt("lblInfo"), ";")
        For i = LBound(strArray) To UBound(strArray)
            lblInfo(i) = strArray(i)
        Next i
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
