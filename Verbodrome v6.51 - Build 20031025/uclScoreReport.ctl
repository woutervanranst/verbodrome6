VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl uclScoreReport 
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   ScaleHeight     =   3060
   ScaleWidth      =   6480
   Begin MSComctlLib.ListView lvwScore 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   435
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlHistory"
      SmallIcons      =   "imlHistory"
      ColHdrIcons     =   "imlHistory"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblScoreFooter 
      Caption         =   "lblScoreFooter"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   2640
      Width           =   6420
   End
   Begin VB.Label lblScoreHeader 
      Caption         =   "lblScoreHeader"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   6420
   End
End
Attribute VB_Name = "uclScoreReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Dim intTotalMaxSc As Integer
Dim intTotalBehSc As Integer
Dim strLanguage() As String
Dim blnResizable As Boolean
'
Private Sub UserControl_Initialize()
    'Kolommen initialiseren
    With UserControl.lvwScore
        .ColumnHeaders.Add , , "", 2160
        .ColumnHeaders.Add , , "", 720
        .ColumnHeaders.Add , , "", 1440
        .ColumnHeaders.Add , , "", 2160
        '
        .BackColor = UserControl.BackColor
    End With
    '
    'Taalelementen
    strLanguage = Split(GetTxt("ScoreEntry"), ";")
End Sub
'
Public Sub AddScore(strTijd As String, intBehSc As Integer, intMaxSc As Integer, intTotalEx As Integer)
    On Error Resume Next
    '
    Dim i As Byte
    '
    Dim lEntry As ListItem
    Dim sEntry As ListSubItem
    '
    With UserControl.lvwScore
        Set lEntry = .ListItems.Add(, , strTijd & ":")
        '
        Set sEntry = lEntry.ListSubItems.Add(, , intBehSc & "/" & intMaxSc)
        Set sEntry = lEntry.ListSubItems.Add(, , strLanguage(0) & Round((intBehSc / intMaxSc * 100), 0) & "%")
        Set sEntry = lEntry.ListSubItems.Add(, , "(" & Round(intMaxSc / CInt(intTotalEx) * 100, 0) & "%" & strLanguage(1))
    End With
    '
    'Header en footer
    intTotalBehSc = intTotalBehSc + intBehSc
    intTotalMaxSc = intTotalMaxSc + intMaxSc
    '
    lblScoreHeader.Caption = Replace(GetTxt("ScoreHeader"), "$aantaloefeningen$", intTotalEx)
    lblScoreFooter.Caption = GetTxt("ScoreFooter")
    lblScoreFooter.Caption = Replace(lblScoreFooter.Caption, "$behsc$", intTotalBehSc)
    lblScoreFooter.Caption = Replace(lblScoreFooter.Caption, "$maxsc$", intTotalMaxSc)
    lblScoreFooter.Caption = Replace(lblScoreFooter.Caption, "$percentage$", Round(intTotalBehSc / intTotalMaxSc * 100, 1))
    lblScoreFooter.Caption = Replace(lblScoreFooter.Caption, "$score20$", Round(intTotalBehSc / intTotalMaxSc * 20, 1))
    '
    'Control resizen
    If blnResizable Then
        lvwScore.Height = lvwScore.ListItems.Count * 210
        lblScoreFooter.Top = lvwScore.Top + lvwScore.Height + 240
        UserControl.Height = lblScoreFooter.Top + lblScoreFooter.Height
    Else
        lvwScore.Top = lblScoreHeader.Height + 120
        lvwScore.Height = 3615 - (lblScoreHeader.Height + 120) - 120 - 195
        lblScoreFooter.Top = lvwScore.Top + lvwScore.Height + 120
        UserControl.Height = 3615
        lvwScore.Width = 5175
    End If
End Sub
'
Public Property Get intBehSc() As Integer
    intBehSc = intTotalBehSc
End Property
'
Public Property Get intMaxSc() As Integer
    intMaxSc = intTotalMaxSc
End Property
'
Public Property Let Resizable(blnResize As Boolean)
    blnResizable = blnResize
End Property
'
Public Sub NoScores()
    lblScoreHeader.Caption = GetTxt("Score_NoAvailable")
    lvwScore.Visible = False
    lblScoreFooter.Visible = False
End Sub
