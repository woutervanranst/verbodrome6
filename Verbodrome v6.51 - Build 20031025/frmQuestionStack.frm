VERSION 5.00
Begin VB.Form frmQuestionStack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmQuestionStack"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboQuestionStack 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtVorigeInvoer 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtAntwoord 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtVraagzin 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Questions in stack:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vorige invoer:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Antwoord:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vraagzin:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   660
   End
End
Attribute VB_Name = "frmQuestionStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' FORM
'
    Private Sub Form_Load()
        Dim i As Integer
        '
        cboQuestionStack.Clear
        '
        For i = 1 To mySession.TotalQuestions
            cboQuestionStack.AddItem i
        Next i
        '
        cboQuestionStack.ListIndex = mySession.CurrentQuestionNumber - 1
        '
        'Taalelementen
        Me.Caption = AppName & " - Questionstack"
        Me.Icon = LoadResPicture(107, vbResIcon) 'frmLoad.Icon
    End Sub
    '
    Private Sub cboQuestionStack_Click()
        If mySession.Stack(cboQuestionStack.Text) Is Nothing Then
            txtVraagzin.Text = ""
            txtAntwoord.Text = ""
            txtVorigeInvoer.Text = ""
        Else
            txtVraagzin.Text = mySession.Stack(cboQuestionStack.Text).QuestionText  'udtQuestions(cboQuestionStack.Text).Vraagzin
            txtAntwoord.Text = mySession.Stack(cboQuestionStack.Text).CorrectAnswer  'udtQuestions(cboQuestionStack.Text).Antwoord
            txtVorigeInvoer.Text = Join(mySession.Stack(cboQuestionStack.Text).PreviousAnswers, ";")   'Join(udtQuestions(cboQuestionStack.Text).Invoer, ";")
        End If
    End Sub
    
