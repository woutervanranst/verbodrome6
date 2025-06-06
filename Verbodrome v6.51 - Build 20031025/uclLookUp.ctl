VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl uclLookUp 
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   ScaleHeight     =   2220
   ScaleWidth      =   4335
   Begin VB.CommandButton cmdSearch 
      Caption         =   "cmdSearch"
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtAnswer 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.ComboBox cboPersonne 
      Height          =   315
      ItemData        =   "uclLookUp.ctx":0000
      Left            =   1200
      List            =   "uclLookUp.ctx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.ComboBox cboTemps 
      Height          =   315
      ItemData        =   "uclLookUp.ctx":0042
      Left            =   1200
      List            =   "uclLookUp.ctx":0088
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ComboBox cboVerbe 
      Height          =   315
      ItemData        =   "uclLookUp.ctx":0260
      Left            =   1200
      List            =   "uclLookUp.ctx":0262
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin MSComctlLib.ProgressBar prgSearch 
      Height          =   150
      Left            =   120
      TabIndex        =   8
      Top             =   2000
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   265
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblPersoon 
      AutoSize        =   -1  'True
      Caption         =   "lblPersoon"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label lblTijd 
      AutoSize        =   -1  'True
      Caption         =   "lblTijd"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1140
      Width           =   405
   End
   Begin VB.Label lblWerkwoord 
      AutoSize        =   -1  'True
      Caption         =   "lblWerkwoord"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   975
   End
End
Attribute VB_Name = "uclLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Public Sub Initialize()
    On Error GoTo ErrorOccured
    errProcStack errPush, "uclLookup.Initialize"
    '
    Dim strTabs() As String
    Dim strVerbes() As String: strVerbes = myFormes.GetEnumerator
    Dim i As Integer
    '
    UserControl.cmdSearch.Caption = GetTxt("cmdSearch")
    '
    strTabs = Split(GetTxt("LookupCaptions"), ";")
    lblWerkwoord.Caption = strTabs(0)
    lblTijd.Caption = strTabs(1)
    lblPersoon.Caption = strTabs(2)
    '
    'Comboboxen inladen
    '1. Werkwoorden inladen
    For i = LBound(strVerbes) To UBound(strVerbes)
        cboVerbe.AddItem strVerbes(i)
    Next i
    '
    '2. Tijden -> Al ingeladen
    '3. Personen -> Al ingeladen
    '
    cboVerbe.ListIndex = 0
    cboTemps.ListIndex = 0
    cboPersonne.ListIndex = 0
    '
    GoTo ExitProc
ErrorOccured::
    If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
    errProcStack errPop, ""
End Sub
'
Private Sub cboVerbe_Click()
    Call LookUpParam
End Sub
'
Private Sub cboTemps_Click()
    Call ConfigureComboBoxes
    Call LookUpParam
End Sub
'
Private Sub cboPersonne_Click()
    Call LookUpParam
End Sub
'
Private Sub LookUpParam()
    On Error Resume Next
    '
    If Not (cboVerbe.ListIndex = -1 Or cboTemps.ListIndex = -1 Or cboPersonne.ListIndex = -1) Then
        txtAnswer.Text = myFormes.Form(cboVerbe.ListIndex + 1, cboTemps.ListIndex + 1, IIf(NumberOfPersons(cboTemps.ListIndex + 1) <> 1, cboPersonne.ListIndex + 1, 1))
        cmdSearch.Enabled = False
    End If
End Sub
'
Private Sub ConfigureComboBoxes()
    On Error GoTo ErrorOccured
    errProcStack errPush, "uclLookUp.ConfigureComboBoxes"
    '
    Select Case cboTemps.ListIndex + 1
        Case 1, 2, 19, 21, 22
            cboPersonne.Enabled = False
            '
        Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
            cboPersonne.Enabled = True
            '
            If cboPersonne.ListCount <> 6 Then
                With cboPersonne
                    .Clear
                    .AddItem "Je"
                    .AddItem "Tu"
                    .AddItem "Il/Elle"
                    .AddItem "Nous"
                    .AddItem "Vous"
                    .AddItem "Ils/Elles"
                    .ListIndex = 0
                End With
            End If
        Case 17, 18
            'Imperatif
            cboPersonne.Enabled = True
            '
            If cboPersonne.ListCount <> 3 Then
                With cboPersonne
                    .Clear
                    .AddItem "Tu"
                    .AddItem "Nous"
                    .AddItem "Vous"
                    .ListIndex = 0
                End With
            End If
        Case 20
            'Participe passé actif/passif
            cboPersonne.Enabled = True
            '
            If cboPersonne.ListCount <> 2 Then
                With cboPersonne
                    .Clear
                    .AddItem "Actif"
                    .AddItem "Passif"
                    .ListIndex = 0
                End With
            End If
    End Select
    '
    GoTo ExitProc
ErrorOccured::
    If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
    errProcStack errPop, ""
End Sub
'
Private Sub cmdSearch_Click()
    On Error GoTo ErrorOccured
    errProcStack errPush, "uclLookup.cmdSearch_Click"
    '
    Dim recLookup As New ADODB.Recordset
    '
    If txtAnswer.Text <> "" Then
        Dim i As Integer, j As Byte, k As Byte
        '
        prgSearch.Min = 0
        prgSearch.Max = myFormes.Count
        '
        Screen.MousePointer = vbHourglass
        DoEvents
        '
        'TODO: je suis venu komt er ook uit: niet ok
        '
        strSQL = "SELECT * FROM Verbes WHERE " & _
            "Instr(InfPr, '" & LCase(txtAnswer.Text) & "') or " & "Instr(InfPa, '" & LCase(txtAnswer.Text) & "') or " & _
            "Instr(IndPr, '" & LCase(txtAnswer.Text) & "') or " & "Instr(IndPC, '" & LCase(txtAnswer.Text) & "') or " & _
            "Instr(IndIpf, '" & LCase(txtAnswer.Text) & "') or " & "Instr(IndPQP, '" & LCase(txtAnswer.Text) & "') or " & _
            "Instr(IndPS, '" & LCase(txtAnswer.Text) & "') or " & "Instr(IndPA, '" & LCase(txtAnswer.Text) & "') or " & _
            "Instr(IndFS, '" & LCase(txtAnswer.Text) & "') or " & "Instr(IndFA, '" & LCase(txtAnswer.Text) & "') or " & _
            "Instr(CondPr, '" & LCase(txtAnswer.Text) & "') or " & "Instr(CondPa, '" & LCase(txtAnswer.Text) & "') or " & _
            "Instr(SubjPr, '" & LCase(txtAnswer.Text) & "') or " & "Instr(SubjPa, '" & LCase(txtAnswer.Text) & "') or " & _
            "Instr(SubjIpf, '" & LCase(txtAnswer.Text) & "') or " & "Instr(SubjPQP, '" & LCase(txtAnswer.Text) & "') or " & _
            "Instr(ImpPr, '" & LCase(txtAnswer.Text) & "') or " & "Instr(ImpPa, '" & LCase(txtAnswer.Text) & "') or " & _
            "Instr(PartPr, '" & LCase(txtAnswer.Text) & "') or " & "Instr(PartPa, '" & LCase(txtAnswer.Text) & "') or " & _
            "Instr(GerPr, '" & LCase(txtAnswer.Text) & "') or " & "Instr(GerPa, '" & LCase(txtAnswer.Text) & "')"
        
        recLookup.Open strSQL, gcnConnection, adUseClient, adLockReadOnly
        
        Debug.Print recLookup.RecordCount
        
        cboVerbe.ListIndex = myFormes.GetIndexNumber(recLookup!InfPr) - 1
        cboTemps.ListIndex = recLookup.AbsolutePosition - 1
        'cboPersonne.ListIndex = k - 1
        
        recLookup.Close
        
        'TODO: Veel te traag vervangen door query
        'IS DIT NUTTIG???? Query opbouwen aan de hand van de benodigde coordinaten
'                strSQL = "SELECT " & IIf(strFieldName = "InfPr", "", "InfPr,") & _
'                    strFieldName & " FROM Verbes WHERE InfPr='" & _
'                    Replace(udtVerbes(udtDatabaseUsage(i).Werkwoord, 1, 1).Text, "'", "''") & "'" & _
'                    " OR InfPr='" & _
'                    Replace(udtVerbes(udtDatabaseUsage(i).Werkwoord, 1, 1).Text, "'", "''") & "*'"
'                'strSQL = Replace(strSQL, "'", Chr(34))
'                recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic

        
'        For i = LBound(udtVerbes) To UBound(udtVerbes)
'            prgSearch = i
'            For j = 1 To 22
'                For k = 1 To 6
'                    If LCase(txtAnswer.Text) = LCase(udtVerbes(i, j, k).Text) Then
'                        'Resultaat gevonden
'                        cboVerbe.ListIndex = i - 1
'                        cboTemps.ListIndex = j - 1
'                        cboPersonne.ListIndex = k - 1
'                        '
'                        prgSearch = prgSearch.Max
'                        Screen.MousePointer = vbDefault
'                        GoTo ExitProc
'                    End If
'                Next k
'            Next j
'        Next i
    End If
    '
    Screen.MousePointer = vbDefault
    MsgBox GetTxt("NoMatches"), vbInformation, AppName
    '
    GoTo ExitProc
ErrorOccured::
    If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
    errProcStack errPop, ""
End Sub
'
Private Sub txtAnswer_Change()
    cmdSearch.Enabled = True
End Sub
'
Public Property Get Werkwoord() As Integer
    Werkwoord = cboVerbe.ListIndex + 1
End Property
'
Public Property Get Tijd() As Byte
    Tijd = CByte(cboTemps.ListIndex + 1)
End Property
'
Public Property Get Persoon() As Byte
    If cboTemps.Enabled Then
        Persoon = CByte(IIf(NumberOfPersons(cboTemps.ListIndex + 1) <> 1, cboPersonne.ListIndex + 1, 1))
    Else
        Persoon = 1
    End If
End Property
'
Public Property Let Werkwoord(intWerkwoord As Integer)
    cboVerbe.ListIndex = intWerkwoord - 1
End Property
'
Public Property Let Tijd(bytTijd As Byte)
    cboTemps.ListIndex = bytTijd - 1
End Property
'
Public Property Let Persoon(bytPersoon As Byte)
    cboPersonne.ListIndex = bytPersoon - 1
End Property
