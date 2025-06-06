VERSION 5.00
Begin VB.Form frmScoreReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmScoreReport"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin Verbodrome6.uclScoreReport uclScoreReport 
      Height          =   3015
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Image imgScore 
      DragIcon        =   "frmScoreReport.frx":0000
      Height          =   480
      Left            =   240
      Picture         =   "frmScoreReport.frx":1CCA
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmScoreReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' FORM
'
    Private Sub Form_Load()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmScoreReport.Form_Load"
        '
        Screen.MousePointer = vbHourglass
        '
        Dim i As Integer
        '
        'Taalelementen inladen
        Me.Caption = AppName & " - " & GetTxt("Score")
        Me.Icon = LoadResPicture(107, vbResIcon) 'frmLoad.Icon
        Me.cmdOK.Caption = GetTxt("cmdOK")
        '
        uclScoreReport.Resizable = True
        '
        For i = 1 To 22
            If mySession.myScoreBoard.NumberOfEntries(i) <> 0 Then
                'Er zijn oefeningen opgelost van deze tijd
                Call uclScoreReport.AddScore(Convert(AbsNumber, Full, i), mySession.myScoreBoard.Score(i), mySession.myScoreBoard.NumberOfEntries(i), mySession.CurrentQuestionNumber - 1)
            End If
        Next i
        '
        Call LogEvent(Score, CStr(uclScoreReport.intBehSc) & "/" & CStr(uclScoreReport.intMaxSc))
        '
        'Form resizen
        cmdOK.Top = uclScoreReport.Top + uclScoreReport.Height + 225
        Me.Height = cmdOK.Top + cmdOK.Height + 540
        '
        Screen.MousePointer = vbDefault
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub cmdOK_Click()
        Unload Me
    End Sub
    '
    Private Sub uclScoreReport_GotFocus()
        cmdOK.SetFocus
    End Sub
