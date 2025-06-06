VERSION 5.00
Begin VB.UserControl uclProgressBar 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   255
   ScaleWidth      =   4800
   Begin VB.Shape shpFrame 
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4800
   End
   Begin VB.Label lblForeGround 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "lblForeGround"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label lblBackground 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "uclProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Const pcRedBackGround = &HFF&
Const pcGreenBackGround = &HC000&
'
Const intMinimum = 0
Const intMaximum = 100
Dim intCurrentValue As Integer
'
Public Property Let Value(intValue As Integer)
    On Error Resume Next
    '
    intCurrentValue = intValue
    '
    lblBackground.BackColor = IIf(intCurrentValue = intMaximum, pcGreenBackGround, pcRedBackGround)
    
    If intCurrentValue <= intMaximum Then
        'Binnen de grezen
        lblBackground.Width = UserControl.Width * (intCurrentValue / intMaximum)
    End If
    '
    lblForeGround.Caption = Round(intCurrentValue / intMaximum * 100, 0) & "%"
End Property
'
Public Property Get Value() As Integer
    Value = intCurrentValue
End Property
'
Public Property Let Max(intNewMaximum As Integer)
    Max = intNewMaximum
End Property
'
Public Property Let Min(intNewMinimum As Integer)
    Min = intNewMinimum
End Property
'
Private Sub UserControl_Resize()
    UserControl.Height = 255
    shpFrame.Height = 255
    shpFrame.Width = UserControl.Width
    lblForeGround.Width = UserControl.Width
End Sub
