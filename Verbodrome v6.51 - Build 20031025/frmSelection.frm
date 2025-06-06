VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmSelection"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlVerbes 
      Left            =   240
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelection.frx":0000
            Key             =   "Leraar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelection.frx":03DC
            Key             =   "Pluriel"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelection.frx":07F7
            Key             =   "Singulier"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelection.frx":0BF8
            Key             =   "TempsCat"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelection.frx":0FCA
            Key             =   "Temps"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelection.frx":13C1
            Key             =   "Bescherelle"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelection.frx":176F
            Key             =   "Category"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabSettings 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "<Werkwoorden>"
      TabPicture(0)   =   "frmSelection.frx":1B27
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkDisplayChecked"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "treCategories"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmSortMethod"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmSearch"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lstVerbes"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSelect(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSelect(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "<Tijden>"
      TabPicture(1)   =   "frmSelection.frx":1B43
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblTotal"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblProbability"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "treTemps"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "uclProbability(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "uclProbability(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "uclProbability(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "uclProbability(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "uclProbability(0)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkAutomaticProbability"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "prgTotalProbability"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "vsbProbabilities"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "<Personen>"
      TabPicture(2)   =   "frmSelection.frx":1B5F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "trePersonnes"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "<Score>"
      TabPicture(3)   =   "frmSelection.frx":1B7B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "uclScoreReport"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.CommandButton cmdSelect 
         Caption         =   "cmdSelect(1)"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   21
         Top             =   1140
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "cmdSelect(0)"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   22
         Top             =   1140
         Width           =   1335
      End
      Begin VB.VScrollBar vsbProbabilities 
         Height          =   3045
         LargeChange     =   2
         Left            =   -69960
         TabIndex        =   17
         Top             =   550
         Visible         =   0   'False
         Width           =   255
      End
      Begin Verbodrome6.uclProgressBar prgTotalProbability 
         Height          =   255
         Left            =   -71760
         TabIndex        =   16
         Top             =   3720
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   450
      End
      Begin VB.CheckBox chkAutomaticProbability 
         Caption         =   "chkAutomaticProbability"
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   3735
         Width           =   2385
      End
      Begin Verbodrome6.uclProbability uclProbability 
         Height          =   615
         Index           =   0
         Left            =   -72320
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   1085
      End
      Begin VB.ListBox lstVerbes 
         Height          =   2310
         Left            =   2640
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   1680
         Width           =   2670
      End
      Begin VB.Frame frmSearch 
         Caption         =   "frmSearch"
         Height          =   650
         Left            =   2640
         TabIndex        =   4
         Top             =   420
         Width           =   2655
         Begin VB.TextBox txtSearchString 
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Text            =   "txtSearchString"
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame frmSortMethod 
         Caption         =   "frmSortMethod"
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   2295
         Begin VB.OptionButton optSortMethod 
            Caption         =   "SortByCat"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optSortMethod 
            Caption         =   "SortByName"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   1935
         End
      End
      Begin MSComctlLib.TreeView trePersonnes 
         Height          =   1770
         Left            =   -74880
         TabIndex        =   6
         Top             =   540
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3122
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   5
         Checkboxes      =   -1  'True
         SingleSel       =   -1  'True
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView treCategories 
         Height          =   1935
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3413
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   5
         HotTracking     =   -1  'True
         SingleSel       =   -1  'True
         Appearance      =   1
      End
      Begin Verbodrome6.uclProbability uclProbability 
         Height          =   615
         Index           =   1
         Left            =   -72315
         TabIndex        =   10
         Top             =   1290
         Visible         =   0   'False
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   1085
      End
      Begin Verbodrome6.uclProbability uclProbability 
         Height          =   615
         Index           =   2
         Left            =   -72315
         TabIndex        =   11
         Top             =   1860
         Visible         =   0   'False
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   1085
      End
      Begin Verbodrome6.uclProbability uclProbability 
         Height          =   615
         Index           =   3
         Left            =   -72315
         TabIndex        =   12
         Top             =   2430
         Visible         =   0   'False
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   1085
      End
      Begin Verbodrome6.uclProbability uclProbability 
         Height          =   615
         Index           =   4
         Left            =   -72315
         TabIndex        =   13
         Top             =   3000
         Visible         =   0   'False
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   1085
      End
      Begin MSComctlLib.TreeView treTemps 
         Height          =   3075
         Left            =   -74880
         TabIndex        =   5
         Top             =   540
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5424
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   5
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   1
      End
      Begin VB.CheckBox chkDisplayChecked 
         Caption         =   "chkDisplayChecked"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3735
         Width           =   2535
      End
      Begin Verbodrome6.uclScoreReport uclScoreReport 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   6376
      End
      Begin VB.Label lblProbability 
         AutoSize        =   -1  'True
         Caption         =   "lblProbability"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72255
         TabIndex        =   18
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         Caption         =   "Totaal:"
         Height          =   195
         Left            =   -72360
         TabIndex        =   15
         Top             =   3740
         Width           =   495
      End
      Begin VB.Shape Shape1 
         Height          =   3075
         Left            =   -72360
         Top             =   540
         Width           =   2670
      End
   End
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private strCurrentCat                   As String
Private strCaptionList()                As String
Private intProbabilityRatings(1 To 22)  As Integer
Private blnUserEdit                     As Boolean
'
' FORM
'
    Private Sub chkDisplayChecked_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.chkDisplayChecked_Click"
        '
        Dim i As Integer
        '
        'If lstVerbes.ListCount = 0 Then Exit Sub
        '
        lstVerbes.Visible = False
        blnUserEdit = False
        '
        If chkDisplayChecked.Value = vbChecked Then
            'v6.5: Geoptimaliseerde code
            With lstVerbes
                For i = .ListCount - 1 To 0 Step -1
                    If Not (.Selected(i)) Then
                        .RemoveItem i
                    End If
                Next i
            End With
        Else
            If optSortMethod(0).Value Then
                'Sorteren op categorie
                strSQL = "SELECT InfPr, TempAllowed, Category FROM Verbes WHERE instr(Category,'" & strCurrentCat & "')<>False"
            Else
                'Sorteren op naam
                strSQL = "SELECT InfPr, TempAllowed FROM Verbes"
            End If
            '
            recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockOptimistic
            '
            lstVerbes.Clear
            With recVerbodrome
                If Not (.EOF And .BOF) Then
                    Do While Not .EOF
                        lstVerbes.AddItem Replace(!InfPr, "*", "")
                        lstVerbes.Selected(lstVerbes.ListCount - 1) = !TempAllowed
                        .MoveNext
                    Loop
                End If
                .Close
            End With
        End If
        '
        lstVerbes.ListIndex = 0
        lstVerbes.Visible = True
        blnUserEdit = True
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub cmdSelect_Click(Index As Integer)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.cmdSelect_Click"
        If blnPM_DebugLog Then Call LogDebug(IIf(Index = 0, "Alles geselecteerd in deze categorie", "Niets geselecteerd in deze categorie"))
        '
        Dim i As Integer
        '
        lstVerbes.Visible = False
        blnUserEdit = False
        '
        If optSortMethod(0).Value Then
            'Sorteren op categorie
            Select Case Index
                Case 0
                    'Alles selecteren
                    If chkDisplayChecked.Value = vbChecked Then
                        'Enkel de werkwoorden die aangevinkt zijn mogen getoond worden
                        strSQL = "SELECT InfPr, TempAllowed, Category FROM Verbes WHERE instr(Category,'" & strCurrentCat & "')<>False"
                        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockOptimistic
                        '
                        lstVerbes.Clear
                        With recVerbodrome
                            If Not (.EOF And .BOF) Then
                                Do While Not .EOF
                                    lstVerbes.AddItem Replace(!InfPr, "*", "")
                                    lstVerbes.Selected(lstVerbes.ListCount - 1) = True
                                    !TempAllowed = True
                                    .Update
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    Else
                        'Alle werkwoorden tonen
                        For i = 0 To lstVerbes.ListCount - 1
                            lstVerbes.Selected(i) = True
                        Next i
                        '
                        strSQL = "SELECT InfPr, TempAllowed, Category FROM Verbes WHERE instr(Category,'" & strCurrentCat & "')<>False AND TempAllowed=False"
                        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockOptimistic
                        '
                        With recVerbodrome
                            If Not (.EOF And .BOF) Then
                                Do While Not .EOF
                                    !TempAllowed = True
                                    .Update
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    End If
                Case 1
                    'Niets selecteren
                    If chkDisplayChecked.Value = vbChecked Then
                        'Enkel de werkwoorden die aangevinkt zijn mogen getoond worden
                        strSQL = "SELECT InfPr, TempAllowed, Category FROM Verbes WHERE instr(Category,'" & strCurrentCat & "')<>False AND TempAllowed=True"
                        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockOptimistic
                        '
                        lstVerbes.Clear
                        With recVerbodrome
                            If Not (.EOF And .BOF) Then
                                Do While Not .EOF
                                    !TempAllowed = False
                                    .Update
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    Else
                        'Alle werkwoorden tonen
                        For i = 0 To lstVerbes.ListCount - 1
                            lstVerbes.Selected(i) = False
                        Next i
                        '
                        strSQL = "SELECT InfPr, TempAllowed, Category FROM Verbes WHERE instr(Category,'" & strCurrentCat & "')<>False AND TempAllowed=True"
                        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockOptimistic
                        '
                        With recVerbodrome
                            If Not (.EOF And .BOF) Then
                                Do While Not .EOF
                                    !TempAllowed = False
                                    .Update
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    End If
            End Select
        Else
            Select Case Index
                Case 0
                    'Alles selecteren
                    If chkDisplayChecked.Value = vbChecked Then
                        'Enkel de werkwoorden die aangevinkt zijn mogen getoond worden
                        strSQL = "SELECT InfPr, TempAllowed, Category FROM Verbes"
                        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockOptimistic
                        '
                        lstVerbes.Clear
                        With recVerbodrome
                            If Not (.EOF And .BOF) Then
                                Do While Not .EOF
                                    lstVerbes.AddItem Replace(!InfPr, "*", "")
                                    lstVerbes.Selected(lstVerbes.ListCount - 1) = True
                                    !TempAllowed = True
                                    .Update
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    Else
                        'Alle werkwoorden tonen
                        For i = 0 To lstVerbes.ListCount - 1
                            lstVerbes.Selected(i) = True
                        Next i
                        '
                        strSQL = "SELECT InfPr, TempAllowed FROM Verbes WHERE TempAllowed=False"
                        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockOptimistic
                        '
                        With recVerbodrome
                            If Not (.EOF And .BOF) Then
                                Do While Not .EOF
                                    !TempAllowed = True
                                    .Update
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    End If
                Case 1
                    'Niets selecteren
                    If chkDisplayChecked.Value = vbChecked Then
                        'Enkel de werkwoorden die aangevinkt zijn mogen getoond worden
                        strSQL = "SELECT InfPr, TempAllowed FROM Verbes WHERE TempAllowed=True"
                        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockOptimistic
                        '
                        lstVerbes.Clear
                        With recVerbodrome
                            If Not (.EOF And .BOF) Then
                                Do While Not .EOF
                                    !TempAllowed = False
                                    .Update
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    Else
                        'Alle werkwoorden tonen
                        For i = 0 To lstVerbes.ListCount - 1
                            lstVerbes.Selected(i) = False
                        Next i
                        '
                        strSQL = "SELECT InfPr, TempAllowed FROM Verbes WHERE TempAllowed=True"
                        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockOptimistic
                        '
                        With recVerbodrome
                            If Not (.EOF And .BOF) Then
                                Do While Not .EOF
                                    !TempAllowed = False
                                    .Update
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    End If
            End Select
        End If
        '
        lstVerbes.ListIndex = 0
        lstVerbes.Visible = True
        blnUserEdit = True
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub Form_Load()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.Form_Load"
        '
        Dim strTabs() As String
        Dim i As Integer
        '
        Screen.MousePointer = vbHourglass
        '
        'Taalelementen inladen
        Me.Caption = AppName & " - " & GetTxt("frmSelection")
        Me.Icon = LoadResPicture(107, vbResIcon)
        Me.optSortMethod(0).Caption = GetTxt("SortByCat")
        Me.optSortMethod(1).Caption = GetTxt("SortByName")
        Me.frmSortMethod.Caption = GetTxt("frmSortMethod")
        Me.frmSearch.Caption = GetTxt("cmdSearch")
        Me.cmdSelect(0).Caption = GetTxt("cmdSelectAll")
        Me.cmdSelect(1).Caption = GetTxt("cmdSelectNone")
        Me.lblProbability.Caption = GetTxt("lblProbability")
        Me.txtSearchString.Text = ""
        Me.chkDisplayChecked.Caption = GetTxt("chkDisplayChecked")
        Me.chkAutomaticProbability.Caption = GetTxt("chkAutomaticProbability")
        Me.chkAutomaticProbability.Value = IIf(GetDBSetting("AutomaticProbability", True), vbChecked, vbUnchecked)
        '
        strTabs = Split(GetTxt("tabSelection"), ";")
        For i = LBound(strTabs) To UBound(strTabs)
            tabSettings.TabCaption(i) = strTabs(i)
        Next i
        '
        treCategories.ImageList = imlVerbes
        treTemps.ImageList = imlVerbes
        trePersonnes.ImageList = imlVerbes
        '
        'Probabiliteitsgetallen inladen
        Call LoadStaticProbabilities
        prgTotalProbability.Value = 100
        '
        'Totaal Scoreboard
        Dim TotalScoreBoard As New clsScoreBoard
        '
        strSQL = "SELECT * FROM Score"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockReadOnly
        '
        uclScoreReport.Resizable = False
        '
        With recVerbodrome
            If Not (.BOF And .EOF) Then
                'Totaal aantal oefeningen uitvissen
                While Not .EOF
                    TotalScoreBoard.BehSc(!Id) = !BehSc + mySession.myScoreBoard.Score(!Id)
                    TotalScoreBoard.MaxSc(!Id) = !MaxSc + mySession.myScoreBoard.NumberOfEntries(!Id)
                    .MoveNext
                Wend
                '
                If TotalScoreBoard.TotalNumberOfEntries <> 0 Then
                    .MoveFirst
                    While Not .EOF
                        If TotalScoreBoard.NumberOfEntries(!Id) <> 0 Then
                            Call uclScoreReport.AddScore(Convert(AbsNumber, Full, !Id), TotalScoreBoard.Score(!Id), TotalScoreBoard.NumberOfEntries(!Id), TotalScoreBoard.TotalNumberOfEntries)
                        End If
                        .MoveNext
                    Wend
                Else
                    Call uclScoreReport.NoScores
                End If
                '
                .Close
            Else
                .Close
            End If
        End With
        '
        'Werkwoordstoelatingen kopieren naar temp
        strSQL = "SELECT InfPr, Allowed, TempAllowed FROM Verbes WHERE Allowed<>TempAllowed"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
        '
        With recVerbodrome
            If Not (.EOF And .BOF) Then
                While Not .EOF
                    !TempAllowed = !Allowed
                    .Update
                    .MoveNext
                Wend
            End If
            .Close
        End With
        '
        'Treeviews klaarmaken
        Call LoadTreeViewItems
        '
        'Microsoft Bugfix
        treCategories.SingleSel = False
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
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.Form_Unload"
        '
        Dim blnAskable As Boolean: blnAskable = True
        '
        'Controleren of er vraagbare vormen zijn overgebleven
        '1. Werkwoorden
        strSQL = "SELECT InfPr, TempAllowed FROM Verbes WHERE TempAllowed=True"
        '
        With recVerbodrome
            .Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
            If .RecordCount = 0 Then blnAskable = False
            .Close
        End With
        '
        '2. Tijden
        strSQL = "SELECT Name, Allowed FROM Allowance WHERE len(Name)>2 AND Allowed=True"
        '
        With recVerbodrome
            .Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
            If .RecordCount = 0 Then blnAskable = False
            .Close
        End With
        '
        '3. Personen
        With trePersonnes
            If .Nodes("1E").Checked = False And .Nodes("2E").Checked = False And .Nodes("3E").Checked = False And .Nodes("1M").Checked = False And .Nodes("2M").Checked = False And .Nodes("3M").Checked = False Then
                blnAskable = False
            End If
        End With
        '
        If blnAskable = False Then
            MsgBox GetTxt("NoAskable"), vbExclamation, AppName
            Cancel = vbCancel
            GoTo ExitProc
        Else
            'Vragen wat de gebruiker wil doen
            Select Case MsgBox(GetTxt("SaveChanges"), vbYesNoCancel + vbQuestion, AppName)
                Case vbYes
                    'Form afsluiten en instellingen opslaan
                    If blnPM_DebugLog Then Call LogDebug("Settings SAVED")
                    Call SaveSettings
                Case vbNo
                    If blnPM_DebugLog Then Call LogDebug("Settings NOT SAVED")
                Case vbCancel
                    'Form niet afsluiten
                    Cancel = vbCancel
            End Select
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub lstVerbes_ItemCheck(Item As Integer)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.lstVerbes_ItemCheck"
        '
        If blnUserEdit Then
            If blnPM_DebugLog Then Call LogDebug("Item " & lstVerbes.List(Item) & " " & IIf(lstVerbes.Selected(Item), "geselecteerd", "gedeselecteerd"))
            '
            strSQL = "SELECT InfPr, TempAllowed FROM Verbes WHERE Left(InfPr, " & Len(lstVerbes.List(Item)) & ")='" & Replace(lstVerbes.List(Item), "'", "''") & "'"
            recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
            '
            With recVerbodrome
                !TempAllowed = lstVerbes.Selected(Item)
                .Update
                .Close
            End With
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub optSortMethod_Click(Index As Integer)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.optSortMethod_Click"
        '
        If Index = 0 Then
            'Op categorie
            With treCategories
                .Visible = True
                '
                .Height = 1935
                .Left = 120
                .Top = 1680
                .Width = 2295
            End With
            '
            With lstVerbes
                .Height = 2310
                .Left = 2640
                .Top = 1680
                .Width = 2670
            End With
            '
            If strCurrentCat <> "" Then
                Call treCategories_NodeClick(treCategories.Nodes(strCurrentCat))
            Else
                lstVerbes.Clear
            End If
        Else
            'Op infinitief
            treCategories.Visible = False
            '
            With lstVerbes
                .Top = 1680
                .Left = 120
                .Height = 1935
                .Width = 5190
            End With
            '
            strSQL = "SELECT InfPr, TempAllowed FROM Verbes" & IIf(chkDisplayChecked.Value = vbChecked, " WHERE TempAllowed=True", "")
            recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
            '
            blnUserEdit = False
            '
            With recVerbodrome
                lstVerbes.Visible = False
                lstVerbes.Clear
                '
                If Not (.EOF And .BOF) Then
                    Do While Not .EOF
                        lstVerbes.AddItem Replace(!InfPr, "*", "")
                        lstVerbes.Selected(lstVerbes.ListCount - 1) = !TempAllowed
                        .MoveNext
                    Loop
                End If
                '
                lstVerbes.Visible = True
                lstVerbes.ListIndex = 0
                .Close
            End With
            '
            blnUserEdit = True
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub treTemps_NodeCheck(ByVal Node As MSComctlLib.Node)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.treTemps_NodeCheck"
        '
        '1. Aanklikken doorvoeren naar de root- en childnodes
        Dim nodChild As Node
        '
        If Node.Children <> 0 Then
            'Root
            '1. Checkboxes controleren
            Set nodChild = Node.Child
            '
            Do
                nodChild.Checked = Node.Checked
                '
                '2. Gevolgen op probabiliteitsmanagement
                Call ProcessProbabilityChanges(nodChild)
                '
                If nodChild = nodChild.LastSibling Then
                    Exit Do
                Else
                    Set nodChild = nodChild.Next
                End If
            Loop
        Else
            'Child
            '1. Checkboxes controleren
            If Not (Node.Checked) Then
                Set nodChild = Node.FirstSibling
                '
                Do
                    If nodChild.Checked Then Exit Do
                    '
                    If nodChild = nodChild.LastSibling Then
                        nodChild.Parent.Checked = False
                        Exit Do
                    Else
                        Set nodChild = nodChild.Next
                    End If
                Loop
            Else
                Node.Parent.Checked = True
            End If
            '
            '2. Gevolgen op probabiliteitsmanagement
            Call ProcessProbabilityChanges(Node)
        End If
        '
        '3. Totaalstatus berekenen
        Call ReCalculateProbabilities
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub treTemps_NodeClick(ByVal Node As MSComctlLib.Node)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.treTemps_NodeClick"
        '
        'Lijst opstellen van de tijdnamen
        Dim i As Byte
        If Node.Children <> 0 Then
            'Hoofdcategorie aangeklikt
            i = 0
            Dim nodTemps As Node
            '
            Set nodTemps = Node.Child
            ReDim strCaptionList(0 To Node.Children)
            '
            Do
                strCaptionList(i) = nodTemps.Parent.Text & " " & nodTemps.Text
                '
                If nodTemps = nodTemps.LastSibling Then
                    Exit Do
                Else
                    i = i + 1
                    Set nodTemps = nodTemps.Next
                End If
            Loop
            '
            With vsbProbabilities
                .Visible = UBound(strCaptionList) > 4
                .Min = 0
                .Max = UBound(strCaptionList) - 5
                .Value = 0
            End With
            '
            Call LoadProbabilityCaptions
            DoEvents
        Else
            'Subcategorie aangeklikt
            vsbProbabilities.Value = 0
            Do
                For i = 0 To 4
                    If (uclProbability(i).Visible) And (Node.Parent.Text & " " & Node.Text = uclProbability(i).Caption) Then
                        uclProbability(i).SetFocus
                        Exit Do
                    End If
                Next i
                '
                vsbProbabilities.Value = vsbProbabilities.Value + 1
            Loop
        End If
        '
        Call SimulateSingleSel(Node, treTemps)
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub txtSearchString_Change()
        Dim i As Integer
        '
        For i = 0 To lstVerbes.ListCount
            If LCase(Left(lstVerbes.List(i), Len(txtSearchString.Text))) = LCase(txtSearchString.Text) Then
                lstVerbes.ListIndex = i
                Exit For
            End If
        Next i
    End Sub
    '
    Private Sub uclProbability_ProbabilityChange(Index As Integer, bytCurrentValue As Byte)
        'Wegschrijven naar variabele
        intProbabilityRatings(Convert(Full, AbsNumber, uclProbability(Index).Caption)) = bytCurrentValue
        'Totaalstatus berekenen
        Call ReCalculateProbabilities
        'Als het percentage 0 is, unchecken
        treTemps.Nodes(Convert(Full, Abreviation, uclProbability(Index).Caption)).Checked = (bytCurrentValue <> 0)
    End Sub
    '
    Private Sub treCategories_NodeClick(ByVal Node As MSComctlLib.Node)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.treCategories_NodeClick"
        If blnPM_DebugLog Then Call LogDebug("Category activated: " & Node.Text)
        '
        strSQL = "SELECT InfPr, TempAllowed, Category FROM Verbes WHERE instr(Category,'" & Node.Key & "')<>False" & IIf(chkDisplayChecked.Value = vbChecked, " AND TempAllowed=True", "")
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
        '
        blnUserEdit = False
        strCurrentCat = Node.Key
        '
        With recVerbodrome
            lstVerbes.Visible = False
            lstVerbes.Clear
            '
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    lstVerbes.AddItem Replace(!InfPr, "*", "")
                    lstVerbes.Selected(lstVerbes.ListCount - 1) = !TempAllowed
                    .MoveNext
                Loop
            End If
            '
            lstVerbes.Visible = True
            lstVerbes.ListIndex = 0
            .Close
        End With
        '
        'If chkDisplayChecked.Value = vbChecked Then Call chkDisplayChecked_Click
        blnUserEdit = True
        '
        Call SimulateSingleSel(Node, treCategories)
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub vsbProbabilities_Change()
        Call LoadProbabilityCaptions
    End Sub
'
' PROCEDURE HULP FUNCTIES
'
    Private Sub LoadProbabilityCaptions()
        Dim i As Byte
        '
        For i = 0 To 4
            If i <= UBound(strCaptionList) - 1 Then
                'Tonen
                uclProbability(i).Visible = True
                'Label van de tijd weergeven
                uclProbability(i).Caption = strCaptionList(i + vsbProbabilities.Value)
                'Percentage inladen
                uclProbability(i).Probability = intProbabilityRatings(Convert(Full, AbsNumber, strCaptionList(i + vsbProbabilities.Value)))
            Else
                uclProbability(i).Visible = False
            End If
        Next i
    End Sub
    '
    Private Sub LoadStaticProbabilities()
        'Kopie van procedure in clsScoreBoard
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.LoadStaticProbabilities"
        '
        'Probabiliteit inladen van de database
        strSQL = "SELECT Name, Probability FROM Allowance WHERE Len(Name)>2"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockReadOnly
        '
        With recVerbodrome
            While Not .EOF
                intProbabilityRatings(Convert(Abreviation, AbsNumber, !Name)) = !Probability
                .MoveNext
            Wend
            .Close
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub LoadTreeViewItems()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.LoadTreeViewItems"
        '
        'Treeviewitems inladen
        '
        '1. treCategories: Hoofdcategoriën toevoegen
        strSQL = "SELECT * FROM Category WHERE Len(CatName)=1"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
        With recVerbodrome
            If Not (.EOF And .BOF) Then
                .MoveFirst
                While Not .EOF
                    treCategories.Nodes.Add , , !CatName, !CatDescription, IIf(!CatName = "B", "Bescherelle", "Leraar")
                    .MoveNext
                Wend
            Else
                .Close
                GoTo ExitProc
            End If
            .Close
        End With
        '
        '2. treCategories: Childcategoriën toevoegen
        strSQL = "SELECT * FROM Category WHERE Len(CatName)>1"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
        With recVerbodrome
            If Not (.EOF And .BOF) Then
                .MoveFirst
                While Not .EOF
                    Me.treCategories.Nodes.Add Left(!CatName, 1), tvwChild, !CatName, !CatDescription, "Category"
                    .MoveNext
                Wend
            Else
                .Close
                GoTo ExitProc
            End If
            .Close
        End With
        '
        '3. treTemps: Tijden inladen
        With treTemps.Nodes
            .Add , , "Infinitif", "Infinitif", "Temps"
            .Add , , "Indicatif", "Indicatif", "Temps"
            .Add , , "Conditionnel", "Conditionnel", "Temps"
            .Add , , "Subjonctif", "Subjonctif", "Temps"
            .Add , , "Impératif", "Impératif", "Temps"
            .Add , , "Participe", "Participe", "Temps"
            .Add , , "Gérondif", "Gérondif", "Temps"
            '
            .Add "Infinitif", tvwChild, "InfPr", "Présent", "TempsCat"
            .Add "Infinitif", tvwChild, "InfPa", "Passé", "TempsCat"
            .Add "Indicatif", tvwChild, "IndPr", "Présent", "TempsCat"
            .Add "Indicatif", tvwChild, "IndPC", "Passé Composé", "TempsCat"
            .Add "Indicatif", tvwChild, "IndIpf", "Imparfait", "TempsCat"
            .Add "Indicatif", tvwChild, "IndPQP", "Plus-que-parfait", "TempsCat"
            .Add "Indicatif", tvwChild, "IndPS", "Passé Simple", "TempsCat"
            .Add "Indicatif", tvwChild, "IndPA", "Passé Antérieur", "TempsCat"
            .Add "Indicatif", tvwChild, "IndFS", "Futur Simple", "TempsCat"
            .Add "Indicatif", tvwChild, "IndFA", "Futur Antérieur", "TempsCat"
            .Add "Conditionnel", tvwChild, "CondPr", "Présent", "TempsCat"
            .Add "Conditionnel", tvwChild, "CondPa", "Passé", "TempsCat"
            .Add "Subjonctif", tvwChild, "SubjPr", "Présent", "TempsCat"
            .Add "Subjonctif", tvwChild, "SubjPa", "Passé", "TempsCat"
            .Add "Subjonctif", tvwChild, "SubjIpf", "Imparfait", "TempsCat"
            .Add "Subjonctif", tvwChild, "SubjPQP", "Plus-que-parfait", "TempsCat"
            .Add "Impératif", tvwChild, "ImpPr", "Présent", "TempsCat"
            .Add "Impératif", tvwChild, "ImpPa", "Passé", "TempsCat"
            .Add "Participe", tvwChild, "PartPr", "Présent", "TempsCat"
            .Add "Participe", tvwChild, "PartPa", "Passé", "TempsCat"
            .Add "Gérondif", tvwChild, "GerPr", "Présent", "TempsCat"
            .Add "Gérondif", tvwChild, "GerPa", "Passé", "TempsCat"
        End With
        '
        '4. trePersons: Personen inladen
        With trePersonnes
            .Nodes.Add , , "1E", "Je", "Singulier"
            .Nodes.Add , , "2E", "Tu", "Singulier"
            .Nodes.Add , , "3E", "Il/Elle", "Singulier"
            .Nodes.Add , , "1M", "Nous", "Pluriel"
            .Nodes.Add , , "2M", "Vous", "Pluriel"
            .Nodes.Add , , "3M", "Ils/Elles", "Pluriel"
        End With
        '
        '5. Toelatingen inladen
        'Tijden
        strSQL = "SELECT * FROM Allowance WHERE Allowed=True"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
        With recVerbodrome
            If Not (.EOF And .BOF) Then
                .MoveFirst
                While Not .EOF
                    If Len(!Name) = 2 Then
                        'Persoon
                        trePersonnes.Nodes(CStr(!Name)).Checked = True
                    Else
                        'Tijd
                        treTemps.Nodes(CStr(!Name)).Checked = True
                        treTemps.Nodes(CStr(!Name)).Parent.Checked = True
                    End If
                    .MoveNext
                Wend
            End If
            .Close
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub SaveSettings()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.SaveSettings"
        '
        'Instellingen opslaan
        '1. Werkwoorden opslaan
        strSQL = "SELECT InfPr, Allowed, TempAllowed FROM Verbes WHERE Allowed<>TempAllowed"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
        '
        With recVerbodrome
            If Not (.EOF And .BOF) Then
                While Not .EOF
                    !Allowed = !TempAllowed
                    .Update
                    .MoveNext
                Wend
            End If
            .Close
        End With
        '
        '2. Tijden, personen en probabiliteit opslaan
        strSQL = "SELECT Name, Allowed, Probability FROM Allowance"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
        With recVerbodrome
            If Not (.EOF And .BOF) Then
                While Not .EOF
                    If Len(!Name) > 2 Then
                        'Tijd
                        !Allowed = CBool(treTemps.Nodes(CStr(!Name)).Checked)
                        !Probability = intProbabilityRatings(Convert(Abreviation, AbsNumber, !Name))
                    Else
                        'Persoon
                        !Allowed = CBool(trePersonnes.Nodes(CStr(!Name)).Checked)
                    End If
                    .Update
                    .MoveNext
                Wend
            Else
                .Close
                GoTo ExitProc
            End If
            .Close
        End With
        '
        '3. Automatische probabiliteitsinstelling
        If GetDBSetting("AutomaticProbability") <> CBool(chkAutomaticProbability.Value) Then
            Call LetDBSetting("AutomaticProbability", (chkAutomaticProbability.Value = vbChecked))
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub ProcessProbabilityChanges(Node As MSComctlLib.Node)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmSelection.ProcessProbabilityChanges"
        '
        Dim i As Byte
        '
        For i = 0 To 4
            If Convert(Full, Abreviation, uclProbability(i).Caption) = Node.Key Then
                'Zichtbaar
                uclProbability(i).Probability = IIf(Node.Checked, intProbabilityRatings(Convert(Full, AbsNumber, uclProbability(i).Caption)), 0)
                Exit For
            End If
            '
            If i = 4 Then
                'Onzichtbaar
                intProbabilityRatings(Convert(Abreviation, AbsNumber, Node.Key)) = IIf(Node.Checked, intProbabilityRatings(Convert(Abreviation, AbsNumber, Node.Key)), 0)
            End If
        Next i
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub ReCalculateProbabilities()
        Dim bytTotalProbability As Byte
        Dim i As Integer
        '
        For i = 1 To 22
            bytTotalProbability = bytTotalProbability + intProbabilityRatings(i)
        Next i
        '
        prgTotalProbability.Value = bytTotalProbability
    End Sub
'
' BUGFIX MICROSOFT
'
    Private Sub SimulateSingleSel(ByVal Node As MSComctlLib.Node, Control As TreeView)
       ' expand current node
       Node.Expanded = True
    
       ' collapse all siblings and parents' siblings
       Dim nodeParent As MSComctlLib.Node
       Set nodeParent = Node
       Do Until nodeParent Is Nothing
          Call CollapseSiblings(nodeParent, Control)
          Set nodeParent = nodeParent.Parent
       Loop
    
       ' collapse child's siblings (collapse all children)
       Dim nodeChild As MSComctlLib.Node
    
       Set nodeChild = Node.Child
       If nodeChild Is Nothing Then Exit Sub
       nodeChild.Expanded = False
       Call CollapseSiblings(nodeChild, Control)
    End Sub
    '
    Private Sub CollapseSiblings(ByVal Node As MSComctlLib.Node, Control As TreeView)
       Dim n As Integer
    
       n = Node.FirstSibling.Index
       If n <> Node.Index Then Node.FirstSibling.Expanded = False
       
       Do While n <> Node.LastSibling.Index
          n = Control.Nodes(n).Next.Index
          If n <> Node.Index Then Control.Nodes(n).Expanded = False
       Loop
    End Sub
