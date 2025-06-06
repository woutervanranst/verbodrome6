VERSION 5.00
Begin VB.UserControl uclAbout 
   BackColor       =   &H00000000&
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3000
   ScaleWidth      =   4800
   Begin VB.Timer tmrScroll 
      Interval        =   80
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "An Onraedt"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   2
      Left            =   405
      TabIndex        =   26
      Top             =   1920
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblPerson(15)"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4545
      Index           =   15
      Left            =   360
      TabIndex        =   25
      Top             =   6720
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Chris Hanssens"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   10
      Left            =   405
      TabIndex        =   24
      Top             =   4920
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wouter Van Ranst"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   7
      Left            =   405
      TabIndex        =   23
      Top             =   3840
      Width           =   4005
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   0
      Left            =   360
      TabIndex        =   22
      Top             =   120
      Width           =   4005
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "© 1999-2003"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   -38
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Software Corporation Program"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   -38
      TabIndex        =   18
      Top             =   2535
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   75
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Image imgBlueLogo 
      Height          =   1995
      Left            =   1365
      Picture         =   "uclAbout.ctx":0000
      Stretch         =   -1  'True
      Top             =   375
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "voor hun antwoord op (al) mijn vragen en onoplosbare bugs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Index           =   7
      Left            =   405
      TabIndex        =   16
      Top             =   6120
      Width           =   4005
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Met dank aan het VANAS team:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   6
      Left            =   405
      TabIndex        =   15
      Top             =   4440
      Width           =   4005
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Werkwoordsdatabase:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   5
      Left            =   405
      TabIndex        =   14
      Top             =   3600
      Width           =   4005
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Vertalingen:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   4
      Left            =   405
      TabIndex        =   13
      Top             =   2760
      Width           =   4005
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Debuggers:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   3
      Left            =   405
      TabIndex        =   12
      Top             =   1440
      Width           =   4005
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmeur:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   1
      Left            =   405
      TabIndex        =   11
      Top             =   480
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Philip Debaere"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   11
      Left            =   405
      TabIndex        =   10
      Top             =   5160
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jo Cap"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   14
      Left            =   405
      TabIndex        =   9
      Top             =   5880
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gunter Troch"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   12
      Left            =   405
      TabIndex        =   8
      Top             =   5400
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dimitri Geudens"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   13
      Left            =   405
      TabIndex        =   7
      Top             =   5640
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Danny Depecker"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   9
      Left            =   405
      TabIndex        =   6
      Top             =   4680
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Brewalan Le Drû"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   8
      Left            =   405
      TabIndex        =   5
      Top             =   4080
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Monique Van Tilborgh"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   5
      Left            =   405
      TabIndex        =   4
      Top             =   3000
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wouter Van Ranst"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   6
      Left            =   405
      TabIndex        =   3
      Top             =   3240
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Steven Peeters"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   1
      Left            =   405
      TabIndex        =   2
      Top             =   1680
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Matthieu Roels"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   4
      Left            =   405
      TabIndex        =   1
      Top             =   2400
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wouter Van Ranst"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   0
      Left            =   405
      TabIndex        =   0
      Top             =   720
      Width           =   4005
   End
   Begin VB.Label lblPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wouter Van Ranst"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   3
      Left            =   405
      TabIndex        =   20
      Top             =   2160
      Width           =   4005
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Met dank aan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   2
      Left            =   360
      TabIndex        =   21
      Top             =   1080
      Width           =   4005
   End
End
Attribute VB_Name = "uclAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Public Event ShowDone()
'
Private Sub tmrScroll_Timer()
    On Error Resume Next
    '
    UserControl.Height = 200 * 15
    UserControl.Width = 320 * 15
    '
    Dim i As Byte
    '
    If lblPerson(lblPerson.UBound).Top + lblPerson(lblPerson.UBound).Height > 0 Then
    'If lblInfo(lblInfo.UBound).Top + lblInfo(lblInfo.UBound).Height > 0 Then
        For i = lblInfo.LBound To lblInfo.UBound
            lblInfo(i).Top = lblInfo(i).Top - 25
        Next i
        '
        For i = lblPerson.LBound To lblPerson.UBound
            lblPerson(i).Top = lblPerson(i).Top - 25
        Next i
    Else
        lblEnd(0).Visible = True
        lblEnd(1).Visible = True
        lblEnd(2).Visible = True
        imgBlueLogo.Visible = True
        '
        tmrScroll.Enabled = False
        '
        'RaiseEvent ShowDone
    End If
End Sub
'
Private Sub UserControl_Initialize()
    On Error Resume Next
    '
    Dim strInfo() As String: strInfo = Split(GetTxt("Info"), ";")
    Dim i As Byte
    '
    For i = LBound(strInfo) To UBound(strInfo)
        lblInfo(i).Caption = strInfo(i)
    Next i
    '
    For i = lblInfo.LBound To lblInfo.UBound
        lblInfo(i).Top = lblInfo(i).Top + UserControl.Height
    Next i
    '
    For i = lblPerson.LBound To lblPerson.UBound
        lblPerson(i).Top = lblPerson(i).Top + UserControl.Height
    Next i
    '
    'Andere groetjes ;)
    lblPerson(15).Caption = "Lode De Jonck" & vbCrLf & Choose(bytLanguage + 1, "mijn vader", "mon père", "my dad") & vbCrLf & "Eric Van Assche" & vbCrLf & "Eric Van Hool" & vbCrLf & "Margueritte Raemaekers" & vbCrLf & "Mark Borremans" & vbCrLf & "Stefan Borremans" & vbCrLf & "Kristel De Leeuw" & vbCrLf & "Annemie Simkens" & vbCrLf & "Nathalie Renard" & vbCrLf & vbCrLf & "Julia De Bie, Jeanninne De Wiest, Gaëlle Derrien, Hubert Derrien, Yves Derrien, Nicole Duval, Dirk Hermans, Annick Laurijssen, Harry Laurijssen, Jean Obenans, Marc Roels, Rudi Rotthier, Olmo Van Assche, Antoon Van Coillie, Dennis Van Landeghem, Emma Van Ranst, Eugène van Ranst, Kristel Van Ranst, Ria Van Ranst, Stefan Van Ranst, Alice Van Tilborgh, Michel Van Tilborgh, Christian Wildiers"
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 200 * 15
    UserControl.Width = 320 * 15
End Sub
