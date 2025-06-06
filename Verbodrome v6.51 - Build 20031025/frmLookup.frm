VERSION 5.00
Begin VB.Form frmLookup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmLookup"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin Verbodrome6.uclLookUp uclLookUp 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4048
   End
End
Attribute VB_Name = "frmLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Sub Form_Load()
    'Taalelementen inladen
    Me.Caption = AppName & " - " & Replace(GetTxt("mnuViewLookup"), "&", "")
    Me.Icon = LoadResPicture(107, vbResIcon) 'frmLoad.Icon
    '
    Call uclLookUp.Initialize
End Sub
