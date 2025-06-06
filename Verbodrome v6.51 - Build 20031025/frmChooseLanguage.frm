VERSION 5.00
Begin VB.Form frmChooseLanguage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmChooseLanguage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      ItemData        =   "frmChooseLanguage.frx":08CA
      Left            =   120
      List            =   "frmChooseLanguage.frx":08D7
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblEngels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Choose your language"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   3500
   End
   Begin VB.Label lblFrans 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Choisissez votre langue"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3500
   End
   Begin VB.Label lblNederlands 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Kies uw taal"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3500
   End
End
Attribute VB_Name = "frmChooseLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private blnUserClickedOK As Boolean
'
' FORM
'
    Private Sub cmdOK_Click()
        blnUserClickedOK = True
        bytLanguage = cboLanguage.ListIndex
        Unload Me
    End Sub
    '
    Private Sub Form_Load()
        cboLanguage.ListIndex = 0
        Me.Caption = AppName
    End Sub
    '
    Private Sub Form_Unload(Cancel As Integer)
        If blnUserClickedOK = False Then
            End
        End If
    End Sub
