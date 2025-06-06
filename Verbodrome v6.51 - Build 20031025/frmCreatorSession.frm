VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreatorSession 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmCreatorSession"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9600
   Icon            =   "frmCreatorSession.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "cmdDelete"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlQuestionList 
      Left            =   8760
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreatorSession.frx":08CA
            Key             =   "question"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmSelectQuestion 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      Begin Verbodrome6.uclLookUp uclLookUp 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4048
      End
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "cmdInsert"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwQuestionList 
      Height          =   3285
      Left            =   4920
      TabIndex        =   2
      Top             =   80
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlQuestionList"
      SmallIcons      =   "imlQuestionList"
      ColHdrIcons     =   "imlQuestionList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
      Begin VB.Menu mnuStreep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "mnuFileExit"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "mnuInfo"
   End
End
Attribute VB_Name = "frmCreatorSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim editSession As New clsSession
'
' FORM
'
    Private Sub Form_Load()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorSession.Form_Load"
        '
        Dim strColumnHeaders() As String: strColumnHeaders = Split(GetTxt("tabSessionCreator"), ";")
        '
        'Taalelementen inladen
        Me.Caption = AppName & " - " & GetTxt("SessionCreator")
        mnuFile.Caption = GetTxt("mnuFile")
        mnuFileExit.Caption = GetTxt("mnuFileExit")
        mnuFileOpenSession.Caption = GetTxt("mnuFileOpen")
        mnuFileSaveSession.Caption = GetTxt("mnuFileSave")
        cmdInsert.Caption = GetTxt("QuestionAdd")
        cmdDelete.Caption = GetTxt("QuestionDelete")
        mnuInfo.Caption = GetTxt("mnuInfo")
        '
        Call uclLookUp.Initialize
        '
        'Kolommen initialiseren
        With Me.lvwQuestionList
            .ColumnHeaders.Add , , "", 300
            .ColumnHeaders.Add , , strColumnHeaders(0), 360
            .ColumnHeaders.Add , , strColumnHeaders(1), 1235 - 80
            .ColumnHeaders.Add , , strColumnHeaders(2), 1800 - 80
            .ColumnHeaders.Add , , strColumnHeaders(3), 800 - 80
            '
            .SortKey = 1
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub Form_Unload(Cancel As Integer)
        Cancel = vbCancel
        Call mnuFileExit_Click
    End Sub
    '
    Private Sub cmdInsert_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorSession.cmdInsert_Click"
        '
        Call editSession.AddQuestion( _
            editSession.TotalQuestions + 1, _
            uclLookUp.Werkwoord, _
            uclLookUp.Tijd, _
            uclLookUp.Persoon)
        '
        Call PrintQuestions
        '
        Me.lvwQuestionList.ListItems(Me.lvwQuestionList.ListItems.Count).Selected = True
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub lvwQuestionList_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorSession.lvwQuestionList_Click"
        '
        With editSession.Stack(CInt(Mid(Me.lvwQuestionList.SelectedItem.ListSubItems(1).Key, 3)))
            uclLookUp.Werkwoord = .VerbeIndex
            uclLookUp.Tijd = .TempsIndex
            uclLookUp.Persoon = .PersonneIndex
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub cmdDelete_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorSession.cmdDelete_Click"
        '
        Dim i As Integer
        Dim intCurrentSelectedItem As Integer
        '
        If Not (Me.lvwQuestionList.SelectedItem Is Nothing) Then
            intCurrentSelectedItem = CInt(Mid(Me.lvwQuestionList.SelectedItem.ListSubItems(1).Key, 3))
            '
            Call editSession.DeleteQuestion(CInt(Mid(Me.lvwQuestionList.SelectedItem.ListSubItems(1).Key, 3)))
            '
            If editSession.TotalQuestions = 0 Then
                Me.lvwQuestionList.ListItems.Clear
            Else
                Call PrintQuestions
            End If
            '
            If intCurrentSelectedItem < lvwQuestionList.ListItems.Count Then
                Me.lvwQuestionList.ListItems(intCurrentSelectedItem).Selected = True
            Else
                If lvwQuestionList.ListItems.Count > 0 Then
                    Me.lvwQuestionList.ListItems(lvwQuestionList.ListItems.Count).Selected = True
                End If
            End If
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
'
' MENUS
'
    Private Sub mnuFileExit_Click()
        If MsgBox(GetTxt("QuitProgram"), vbQuestion + vbYesNo, AppName) = vbYes Then
            End
        End If
    End Sub
    '
    Private Sub mnuFileOpenSession_Click()
        If editSession.ShowOpenSessionDialog Then
            If editSession.Incremental Then
                MsgBox GetTxt("SessionCreator_IncrementalSessionFile"), vbInformation, AppName
            End If
            '
            mnuFileSaveSession.Enabled = Not (editSession.Incremental)
            cmdInsert.Enabled = Not (editSession.Incremental)
            cmdDelete.Enabled = Not (editSession.Incremental)
            '
            lvwQuestionList.ListItems.Clear
            '
            Call PrintQuestions
        End If
    End Sub
    '
    Private Sub mnuFileSaveSession_Click()
        editSession.ShowSaveSessionDialog
    End Sub
    '
    Private Sub mnuInfo_Click()
        frmAbout.Show 1
    End Sub
'
' PROCEDURE HULP FUNCTIES
'
    Private Sub PrintQuestions()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorSession.PrintQuestions"
        '
        Dim i As Integer
        Dim lEntry As ListItem
        Dim sEntry As ListSubItem
        '
        Me.lvwQuestionList.ListItems.Clear
        '
        For i = 1 To editSession.TotalQuestions
            If Not (editSession.QuestionSlotEmpty(i)) Then
                With Me.lvwQuestionList
                    Set lEntry = .ListItems.Add(, , , , "question")
                    '
                    Set sEntry = lEntry.ListSubItems.Add(, "Q_" & CStr(i), Right("0" & CStr(i), 2))
                    Set sEntry = lEntry.ListSubItems.Add(, , editSession.Stack(i).Verbe)
                    Set sEntry = lEntry.ListSubItems.Add(, , Convert(AbsNumber, Full, editSession.Stack(i).TempsIndex))
                    Set sEntry = lEntry.ListSubItems.Add(, , Choose(editSession.Stack(i).PersonneIndex, "Je", "Tu", "Il/Elle", "Nous", "Vous", "Ils/Elles"))
                End With
            End If
        Next i
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
