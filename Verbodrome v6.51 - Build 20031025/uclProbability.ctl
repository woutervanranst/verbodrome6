VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl uclProbability 
   BackStyle       =   0  'Transparent
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   PropertyPages   =   "uclProbability.ctx":0000
   ScaleHeight     =   615
   ScaleWidth      =   2385
   Begin MSComctlLib.ProgressBar prgProbability 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtPercentage 
      Height          =   285
      Left            =   1365
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   375
   End
   Begin VB.VScrollBar vsbProbability 
      Height          =   285
      Left            =   2040
      Max             =   2
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblProbability 
      Caption         =   "lblProbability"
      Height          =   195
      Left            =   15
      TabIndex        =   3
      Top             =   0
      Width           =   2070
   End
   Begin VB.Label lblPercentSymbol 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   285
      Width           =   120
   End
End
Attribute VB_Name = "uclProbability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'
Public Event ProbabilityChange(bytCurrentValue As Byte)
'
Private bytCurrentPercentage As Byte
'
Public Property Let FontBold(blnBold As Boolean)
    lblProbability.FontBold = blnBold
End Property
'
Public Property Let Caption(strProbabilityName As String)
    lblProbability.Caption = strProbabilityName
End Property
'
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
   Caption = lblProbability.Caption
End Property
'
Public Property Let Probability(intNewProbability As Integer)
    txtPercentage.Text = CStr(intNewProbability)
End Property
'
Public Property Get Probability() As Integer
Attribute Probability.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Probability = CInt(txtPercentage.Text)
End Property
'
Private Sub txtPercentage_Change()
    On Error Resume Next
    '
    If txtPercentage.Text <> "" Then
        If txtPercentage >= 0 And txtPercentage <= 100 Then
            prgProbability.Value = txtPercentage.Text
        Else
            Select Case txtPercentage
                Case Is < 0: txtPercentage.Text = 0
                Case Is > 100: txtPercentage.Text = 100
            End Select
        End If
    Else
        txtPercentage.Text = bytCurrentPercentage
    End If
    '
    bytCurrentPercentage = CByte(txtPercentage.Text)
    prgProbability.Value = bytCurrentPercentage
    '
    RaiseEvent ProbabilityChange(bytCurrentPercentage)
End Sub
'
Private Sub txtPercentage_GotFocus()
    With txtPercentage
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
'
Private Sub txtPercentage_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack And IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub
'
Private Sub UserControl_Initialize()
    prgProbability.Min = 0
    prgProbability.Max = 100
    vsbProbability.Value = 1
End Sub
'
Private Sub UserControl_Resize()
    UserControl.Width = 2385
    UserControl.Height = 615
End Sub
'
Private Sub vsbProbability_Change()
    If Not (vsbProbability.Value = 2 And txtPercentage.Text = "100") Or Not (vsbProbability.Value = 0 And txtPercentage.Text = "0") Then
        txtPercentage.Text = Replace(txtPercentage.Text, "%", "") - (vsbProbability - 1)
        vsbProbability.Value = 1
    End If
End Sub
