VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmHistory"
   ClientHeight    =   6510
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgHistoryItems 
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   6120
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog dlgExport 
      Left            =   8040
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmSelection 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
      Begin VB.CheckBox chkHistoryItem 
         Caption         =   "chkHistoryItem"
         Height          =   240
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkHistoryItem 
         Caption         =   "chkHistoryItem"
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkHistoryItem 
         Caption         =   "chkHistoryItem"
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkHistoryItem 
         Caption         =   "chkHistoryItem"
         Height          =   240
         Index           =   3
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkHistoryItem 
         Caption         =   "chkHistoryItem"
         Height          =   240
         Index           =   4
         Left            =   480
         TabIndex        =   3
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmHistory.frx":0000
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "frmHistory.frx":0428
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   2
         Left            =   120
         Picture         =   "frmHistory.frx":082E
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   3
         Left            =   120
         Picture         =   "frmHistory.frx":08A9
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   4
         Left            =   120
         Picture         =   "frmHistory.frx":0932
         Top             =   1680
         Width           =   240
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1815
         Left            =   60
         Top             =   180
         Width           =   375
      End
   End
   Begin MSComctlLib.TreeView treSessions 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   7223
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   617
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "imlHistory"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlHistory 
      Left            =   8640
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":0D2B
            Key             =   "command"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":10D5
            Key             =   "session"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":14E8
            Key             =   "daysession"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":191D
            Key             =   "correct"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":19A8
            Key             =   "error"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":1A8E
            Key             =   "system"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":1EC6
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":1F6B
            Key             =   "question"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":2381
            Key             =   "score"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":278A
            Key             =   "timelimit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":2B58
            Key             =   "wrong"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistory.frx":2BF1
            Key             =   "update"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwHistory 
      Height          =   5895
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "mnuFile"
      Begin VB.Menu mnuFileExportToFile 
         Caption         =   "mnuFileExportToFile"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "mnuFileExit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "mnuView"
      Begin VB.Menu mnuViewItem 
         Caption         =   "mnuViewIcon"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "mnuViewDate"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "mnuViewTime"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "mnuViewDescription"
         Checked         =   -1  'True
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private strCurrentHistoryItem       As String
Private strHistoryItems(0 To 15)    As String
'
Private Enum HistoryItems
    HistoryItem_AskedFormsReset = 0
    HistoryItem_CommandInput = 1
    HistoryItem_DUFSuccesful = 2
    HistoryItem_Empty = 3
    HistoryItem_Error = 4
    HistoryItem_FirstTimeCorrect = 5
    HistoryItem_FirstTimeWrong = 6
    HistoryItem_NotFirstTimeCorrect = 7
    HistoryItem_NotFirstTimeWrong = 8
    HistoryItem_ProgramShutDown = 9
    HistoryItem_ProgramStartup = 10
    HistoryItem_ScoreAsked = 11
    HistoryItem_SelectionWindowCall = 12
    HistoryItem_SettingChanged = 13
    HistoryItem_TimeLimitOff = 14
    HistoryItem_TimeLimitOn = 15
End Enum
'
' FORM
'
    Private Sub chkHistoryItem_Click(Index As Integer)
        If strCurrentHistoryItem <> "" Then Call treSessions_NodeClick(treSessions.Nodes(strCurrentHistoryItem))
        If blnPM_DebugLog Then Call LogDebug("chkHistoryItem_Click: " & Index & ".Value=" & chkHistoryItem(Index).Value)
    End Sub
    '
    Private Sub Form_Load()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmHistory.Form_Load"
        '
        Dim strText() As String
        Dim i As Byte
        '
        'Taalelementen inladen
        Me.Caption = AppName & " - " & Replace(GetTxt("mnuFileHistory"), "&", "")
        Me.Icon = LoadResPicture(107, vbResIcon) 'frmLoad.Icon
        mnuFile.Caption = GetTxt("mnuFile")
        mnuFileExit.Caption = GetTxt("mnuFileExit")
        mnuView.Caption = GetTxt("mnuView")
        mnuFileExportToFile.Caption = GetTxt("mnuFileExportToFile")
        '
        strText = Split(GetTxt("HistoryItems"), ";")
        For i = LBound(strText) To UBound(strText)
            chkHistoryItem(i).Caption = strText(i)
        Next i
        '
        strText = Split(GetTxt("mnuViewItem"), ";")
        For i = LBound(strText) To UBound(strText)
            mnuViewItem(i).Caption = strText(i)
        Next i
        '
        'Uit de database de taalelementen inladen (snelheid)
        strHistoryItems(HistoryItem_AskedFormsReset) = GetTxt("HistoryItem_AskedFormsReset")
        strHistoryItems(HistoryItem_CommandInput) = GetTxt("HistoryItem_CommandInput")
        strHistoryItems(HistoryItem_DUFSuccesful) = GetTxt("HistoryItem_DUFSuccesful")
        strHistoryItems(HistoryItem_Empty) = GetTxt("HistoryItem_Empty")
        strHistoryItems(HistoryItem_Error) = GetTxt("HistoryItem_Error")
        strHistoryItems(HistoryItem_FirstTimeCorrect) = GetTxt("HistoryItem_FirstTimeCorrect")
        strHistoryItems(HistoryItem_FirstTimeWrong) = GetTxt("HistoryItem_FirstTimeWrong")
        strHistoryItems(HistoryItem_NotFirstTimeCorrect) = GetTxt("HistoryItem_NotFirstTimeCorrect")
        strHistoryItems(HistoryItem_NotFirstTimeWrong) = GetTxt("HistoryItem_NotFirstTimeWrong")
        strHistoryItems(HistoryItem_ProgramShutDown) = GetTxt("HistoryItem_ProgramShutDown")
        strHistoryItems(HistoryItem_ProgramStartup) = GetTxt("HistoryItem_ProgramStartup")
        strHistoryItems(HistoryItem_ScoreAsked) = GetTxt("HistoryItem_ScoreAsked")
        strHistoryItems(HistoryItem_SelectionWindowCall) = GetTxt("HistoryItem_SelectionWindowCall")
        strHistoryItems(HistoryItem_SettingChanged) = GetTxt("HistoryItem_SettingChanged")
        strHistoryItems(HistoryItem_TimeLimitOff) = GetTxt("HistoryItem_TimeLimitOff")
        strHistoryItems(HistoryItem_TimeLimitOn) = GetTxt("HistoryItem_TimeLimitOn")
        '
        'Verschillende sessies inladen
        '1. Datums van de verschillende sessies
        Dim dtmPreviousDate As Date
        strSQL = "SELECT Date FROM History ORDER BY Date"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockReadOnly
        '
        With recVerbodrome
            If Not (.BOF And .EOF) Then
                While Not .EOF
                    If dtmPreviousDate <> !Date Then
                        treSessions.Nodes.Add , , CStr(!Date), !Date, "daysession"
                        dtmPreviousDate = !Date
                    End If
                    .MoveNext
                Wend
            End If
            .Close
        End With
        '
        '2. Verschillende tijdstippen
        strSQL = "SELECT Id, Date, Time FROM History WHERE EventType=0"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
        '
        With recVerbodrome
            If Not (.BOF And .EOF) Then
                While Not .EOF
                    treSessions.Nodes.Add CStr(!Date), tvwChild, "N_" & CStr(!Id), CStr(!Time), "session"
                    .MoveNext
                Wend
            End If
            .Close
        End With
        '
        '3. Kolommen initialiseren
        With Me.lvwHistory
            .ColumnHeaders.Add , , "", 0
            .ColumnHeaders.Add , , "", 360
            .ColumnHeaders.Add , , "", 1050
            .ColumnHeaders.Add , , "", 810
            .ColumnHeaders.Add , , "", 7000
            '
            .SortKey = 3
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub lvwHistory_KeyDown(KeyCode As Integer, Shift As Integer)
        If Not (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight) Then
            KeyCode = 0
            Shift = 0
        End If
    End Sub
    '
    Private Sub treSessions_NodeClick(ByVal Node As MSComctlLib.Node)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmHistory.treSessions_NodeClick"
        '
        Dim strParameters As String
        Dim strParamList() As String
        Dim Question As New clsQuestion
        '
        Screen.MousePointer = vbHourglass
        DoEvents
        '
        strCurrentHistoryItem = Node.Key
        mnuFileExportToFile.Enabled = True
        '
        If Node.Children <> 0 Then
            'Root
            'strSQL = "SELECT * FROM History WHERE Date=#" & ConvertToUSNotation(Node.Text) & "#"
            'strSQL = "SELECT * FROM History WHERE Date=#" & Format(Node.Text, "dd/mm/yyyy") & "#"
            strSQL = "SELECT * FROM History WHERE CStr(Date)='" & Format(Node.Text, "dd/mm/yyyy") & "'"
            'Exit Sub
        Else
            'Child
            'Voorste limiet opleggen
            strSQL = "SELECT * FROM History WHERE Id>" & (CInt(Replace(Node.Key, "N_", "")) - 1)
            '
            'Achterste limiet opleggen
            If Node <> Node.LastSibling Then
                'Vóór het vorige tijdstip
                strSQL = strSQL & " AND Id<" & Replace(Node.Next.Key, "N_", "")
            Else
                'Laatste node -> laatste oefensessie van de dag
                strSQL = strSQL & " AND CStr(Date)='" & Format(Node.Parent.Text, "dd/mm/yyyy") & "'"
            End If
        End If
        Me.lvwHistory.ListItems.Clear
        '
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
        If blnPM_DebugLog Then Call LogDebug("Query: " & strSQL)
        '
        With recVerbodrome
            If Not (.BOF And .EOF) Then
                prgHistoryItems.Min = 0
                prgHistoryItems.Max = .RecordCount
                prgHistoryItems = 0
                '
                While Not .EOF
                    strParameters = ""
                    Select Case !EventType
                        Case 0, 1, 5, 6, 7, 9, 12, 13
                            'Systeemevent
                            If chkHistoryItem(0).Value = vbChecked Then
                                Select Case !EventType
                                    Case 0
                                        'Programma opstart
                                        strParameters = strHistoryItems(HistoryItem_ProgramStartup)
                                    Case 1
                                        'Programma afgesloten
                                        strParameters = strHistoryItems(HistoryItem_ProgramShutDown)
                                    Case 5
                                        'Commando
                                        strParameters = strHistoryItems(HistoryItem_CommandInput)
                                        strParameters = Replace(strParameters, "$command$", !Parameters)
                                    Case 6
                                        'Selectiescherm opgeroepen
                                        strParameters = strHistoryItems(HistoryItem_SelectionWindowCall)
                                    Case 7
                                        'Fout, of Historiekdatabase geleegd
                                        If !Parameters = "HistoryEmpty" Then
                                            strParameters = strHistoryItems(HistoryItem_Empty)
                                        Else
                                            strParamList = Split(!Parameters, "$")
                                            strParameters = strHistoryItems(HistoryItem_Error)
                                            strParameters = Replace(strParameters, "$procedurename$", strParamList(0))
                                            strParameters = Replace(strParameters, "$errorcode$", strParamList(1))
                                            strParameters = Replace(strParameters, "$description$", strParamList(2))
                                            strParameters = Replace(strParameters, "$userresponse$", strParamList(3))
                                        End If
                                    Case 9
                                        'Gevraagde vormen reset
                                        strParameters = strHistoryItems(HistoryItem_AskedFormsReset)
                                    Case 12
                                        'Veranderde instelling
                                        strParameters = strHistoryItems(HistoryItem_SettingChanged)
                                        strParameters = Replace(strParameters, "$setting$", Left(!Parameters, InStr(!Parameters, "$") - 1))
                                        strParameters = Replace(strParameters, "$value$", Mid(!Parameters, InStr(!Parameters, "$") + 1))
                                    Case 13
                                        'Updatefile met succes ingevoerd
                                        strParameters = strHistoryItems(HistoryItem_DUFSuccesful)
                                End Select
                                Call AddHistoryItem(!EventType, Format(!Date, "dd/mm/yyyy"), CStr(!Time), strParameters)
                            End If
                        Case 2
                            'Vraag
                            If chkHistoryItem(1).Value = vbChecked Then
                                strParamList = Split(!Parameters, "$")
                                '
                                Call AddHistoryItem(!EventType, Format(!Date, "dd/mm/yyyy"), CStr(!Time), QuestionText(strParamList(0), CByte(strParamList(1)), CByte(strParamList(2))))
                            End If
                       Case 3, 4
                            strParamList = Split(!Parameters, "$")
                            '
                            If !EventType = 3 Then
                                'Correct antwoord
                                strParameters = IIf(CByte(strParamList(0)) = 1, strHistoryItems(HistoryItem_FirstTimeCorrect), strHistoryItems(HistoryItem_NotFirstTimeCorrect))
                            Else
                                'Fout antwoord
                                strParameters = IIf(CByte(strParamList(0)) = 1, strHistoryItems(HistoryItem_FirstTimeWrong), strHistoryItems(HistoryItem_NotFirstTimeWrong))
                                strParameters = Replace(strParameters, "$answer$", strParamList(UBound(strParamList)))
                            End If
                            '
                            strParameters = Replace(strParameters, "$time$", strParamList(1))
                            If CByte(strParamList(0)) > 1 Then strParameters = Replace(strParameters, "$tries$", strParamList(0))
                            '
                            If chkHistoryItem(!EventType - 1).Value = vbChecked Then
                                Call AddHistoryItem(!EventType, Format(!Date, "dd/mm/yyyy"), CStr(!Time), strParameters)
                            End If
                        Case 8, 10, 11
                            'Programma
                            If chkHistoryItem(4).Value = vbChecked Then
                                Select Case !EventType
                                    Case 8
                                        'Score opgevraagd
                                        strParameters = strHistoryItems(HistoryItem_ScoreAsked)
                                        strParameters = Replace(strParameters, "$score$", !Parameters)
                                    Case 10
                                        'Tijdslimiet aan
                                        strParameters = strHistoryItems(HistoryItem_TimeLimitOn)
                                        strParameters = Replace(strParameters, "$timelimit$", !Parameters)
                                    Case 11
                                        'Tijdslimiet af
                                        strParameters = strHistoryItems(HistoryItem_TimeLimitOff)
                                End Select
                               Call AddHistoryItem(!EventType, Format(!Date, "dd/mm/yyyy"), CStr(!Time), strParameters)
                            End If
                    End Select
                    '
                    prgHistoryItems = prgHistoryItems + 1
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
        Screen.MousePointer = vbDefault
        errProcStack errPop, ""
    End Sub
'
' MENUS
'
    Private Sub mnuFileExit_Click()
        Unload Me
    End Sub
    '
    Private Sub mnuFileExportToFile_Click()
        On Error Resume Next
        '
        If Me.lvwHistory.ListItems.Count = 0 Then Exit Sub
        '
ChooseOtherFileName::
        With dlgExport
            .DialogTitle = GetTxt("HistoryExportSelectFile")
            .Filter = GetTxt("TextFile") & " (*.txt)|*.txt"
            .ShowSave
            '
            If .FileName = "" Then
                Exit Sub
            ElseIf LCase(Dir(.FileName)) = LCase(.FileTitle) Then
                Select Case MsgBox(GetTxt("WarningFileExists"), vbYesNoCancel + vbQuestion, AppName)
                    Case vbNo
                        GoTo ChooseOtherFileName
                    Case vbCancel
                        Exit Sub
                End Select
            End If
        End With
        '
        Dim strHistoryEntry As String
        Dim BNr: BNr = FreeFile
        Open dlgExport.FileName For Output As #BNr
        '
        Dim i As Integer, j As Byte
        '
        prgHistoryItems.Min = 0
        prgHistoryItems.Max = Me.lvwHistory.ListItems.Count
        '
        Screen.MousePointer = vbHourglass
        '
        For i = 1 To Me.lvwHistory.ListItems.Count
            Select Case CByte(Me.lvwHistory.ListItems(i).Text)
                Case 0, 1, 5, 6, 7, 9, 12, 13
                    'Systeem
                    strHistoryEntry = "S    - "
                Case 2
                    'Vraag
                    strHistoryEntry = " Q   - "
                Case 3, 4
                    'Antwoord
                    strHistoryEntry = "  A  - "
                Case 8, 10, 11
                    'Programma
                    strHistoryEntry = "   P - "
            End Select
            '
            For j = 2 To 4
                strHistoryEntry = strHistoryEntry & Me.lvwHistory.ListItems(i).ListSubItems(j).Text
                strHistoryEntry = strHistoryEntry & Choose(j - 1, "@", " : ", "")
            Next j
            '
            Print #BNr, strHistoryEntry
            prgHistoryItems = i
        Next i
        '
        Close #BNr
        '
        Screen.MousePointer = vbDefault
    End Sub
    '
    Private Sub mnuViewItem_Click(Index As Integer)
        mnuViewItem(Index).Checked = Not (mnuViewItem(Index).Checked)
        '
        Select Case Index
            Case 0: Me.lvwHistory.ColumnHeaders(2).Width = IIf(mnuViewItem(Index).Checked, 360, 0)
            Case 1: Me.lvwHistory.ColumnHeaders(3).Width = IIf(mnuViewItem(Index).Checked, 1050, 0)
            Case 2: Me.lvwHistory.ColumnHeaders(4).Width = IIf(mnuViewItem(Index).Checked, 810, 0)
            Case 3: Me.lvwHistory.ColumnHeaders(5).Width = IIf(mnuViewItem(Index).Checked, 7000, 0)
        End Select
    End Sub
'
' PROCEDURE HULP FUNCTIES
'
    Private Sub AddHistoryItem(EventType As Byte, strDate As String, strTime As String, Optional strParameters As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmHistory.AddHistoryItem"
        '
        If blnPM_DebugLog Then Call LogDebug("Adding item: " & EventType & ", Parameters: " & strParameters)
        '
        Dim strParamList() As String: strParamList = Split(strDate & "$" & strTime & "$" & strParameters, "$")
        Dim strIcon As String
        Dim i As Byte
        '
        Dim lEntry As ListItem
        Dim sEntry As ListSubItem
        '
        Select Case EventType
            Case 0, 1, 9: strIcon = "system"
            Case 2: strIcon = "question"
            Case 3: strIcon = "correct"
            Case 4: strIcon = "wrong"
            Case 5, 6: strIcon = "command"
            Case 7: strIcon = "error"
            Case 8: strIcon = "score"
            Case 10, 11: strIcon = "timelimit"
            Case 12: strIcon = "properties"
            Case 13: strIcon = "update"
        End Select
        '
        With Me.lvwHistory
            Set lEntry = .ListItems.Add(, , EventType)
            '
            Set sEntry = lEntry.ListSubItems.Add(, , , strIcon)
            For i = LBound(strParamList) To UBound(strParamList)
                Set sEntry = lEntry.ListSubItems.Add(, , strParamList(i))
            Next i
        End With
        '
        GoTo ExitProc
ErrorOccured::
        Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Function QuestionText(strVerbe As String, bytTemps As Byte, bytPersonne As Byte) As String
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmHistory.QuestionText"
        '
        If bytTemps >= gcIndPr And bytTemps <= gcSubjPQP Then
            'Met persoon
            QuestionText = GetTxt("QuestionTime")
            QuestionText = Replace(QuestionText, "$persoon$", PersonneText(bytPersonne))
        Else
            'Zonder persoon
            QuestionText = GetTxt("QuestionTimeLess")
        End If
        '
        'Infinitief implementeren
        QuestionText = Replace(QuestionText, "$infinitief$", strVerbe)
        '
        'Tijd implementeren
        QuestionText = Replace(QuestionText, "$tijd$", TempsText(bytTemps, bytPersonne))
        '
        'In het Frans: niet 'donnez le infinitif' maar 'donnez l'infinitif'
        If bytLanguage = 1 Then
            QuestionText = Replace(QuestionText, "le i", "l'i")
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Private Function TempsText(bytTemps As Byte, bytPersonne As Byte) As String
        If bytTemps = gcImpPr Or bytTemps = gcImpPa Then
            'Imperatieven
            TempsText = LCase(Convert(AbsNumber, ShortFull, bytTemps) & " " & Choose(bytPersonne, "Sing.(-)", "1e Pl.(+)", "2e Pl(+)"))
        ElseIf bytTemps = gcPartPa Then
            'Participe Passé
            TempsText = LCase(Convert(AbsNumber, ShortFull, bytTemps) & IIf(bytPersonne = 1, "", " Pass."))
        Else
            'Alle andere tijden
            TempsText = LCase(Convert(AbsNumber, ShortFull, bytTemps))
        End If
    End Function
    '
    Private Function PersonneText(bytPersonne As Byte) As String
        Dim strPersonneCaptions() As String
        strPersonneCaptions = Split(GetTxt("QuestionPersons"), ";")
        PersonneText = strPersonneCaptions(bytPersonne - 1)
    End Function
