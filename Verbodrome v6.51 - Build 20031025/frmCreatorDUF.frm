VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreatorDUF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmCreatorDUF"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9480
   Icon            =   "frmCreatorDUF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCommands 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   4980
      Width           =   9255
   End
   Begin VB.CheckBox chkDisplayChecked 
      Caption         =   "chkDisplayChecked"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   4560
      Width           =   4575
   End
   Begin MSComctlLib.ImageList imlVerbes 
      Left            =   3960
      Top             =   4080
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
            Picture         =   "frmCreatorDUF.frx":08CA
            Key             =   "Leraar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreatorDUF.frx":0CA6
            Key             =   "Pluriel"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreatorDUF.frx":10C1
            Key             =   "Singulier"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreatorDUF.frx":14C2
            Key             =   "TempsCat"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreatorDUF.frx":1894
            Key             =   "Temps"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreatorDUF.frx":1C8B
            Key             =   "Bescherelle"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreatorDUF.frx":2039
            Key             =   "Category"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstVerbes 
      Height          =   4335
      Left            =   4800
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin MSComctlLib.TreeView treCategories 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8281
      _Version        =   393217
      Indentation     =   1411
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imlVerbes"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "mnuFile"
      Begin VB.Menu mnuFileNew 
         Caption         =   "mnuFileNew"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "mnuFileOpen"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "mnuFileSave"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuStreep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "mnuFileExit"
      End
   End
   Begin VB.Menu mnuCategory 
      Caption         =   "mnuCategory"
      Begin VB.Menu mnuCategoryAdd 
         Caption         =   "mnuCategoryAdd"
         Begin VB.Menu mnuCategoryAddItem 
            Caption         =   "&Hoofdcategorie"
            Index           =   0
            Shortcut        =   ^{INSERT}
         End
         Begin VB.Menu mnuCategoryAddItem 
            Caption         =   "&Categorie"
            Index           =   1
            Shortcut        =   +{INSERT}
         End
      End
      Begin VB.Menu mnuCategoryEdit 
         Caption         =   "&Hernoemen"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuCategoryDelete 
         Caption         =   "&Verwijderen"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "mnuInfo"
   End
End
Attribute VB_Name = "frmCreatorDUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private blnUserEdit                     As Boolean
Private udtDUFCommands()                As DUFCommand
Private udtAllowanceChanges()           As AllowanceChange
Private WithEvents clsMyDUFInterpreter  As clsDUFInterpreter
Attribute clsMyDUFInterpreter.VB_VarHelpID = -1
'
Private strText                         As String
Private strDUFCommand                   As String
'
Private Type DUFCommand
    DUFCommand                          As String
    Text                                As String
End Type
'
Private Type AllowanceChange
    Category                            As String
    Infinitive                          As String
    Allowed                             As Boolean
End Type
'
' FORM
'
    Private Sub Form_Load()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.Form_Load"
        '
        'Taalelementen inladen
        Me.Caption = AppName & " - " & GetTxt("DUFCreator")
        Me.chkDisplayChecked.Caption = GetTxt("chkDisplayChecked")
        '
        Me.mnuFile.Caption = GetTxt("mnuFile")
        Me.mnuFileNew.Caption = GetTxt("mnuFileNew")
        Me.mnuFileOpen.Caption = GetTxt("mnuFileOpen")
        Me.mnuFileSave.Caption = GetTxt("mnuFileSave")
        Me.mnuFileExit.Caption = GetTxt("mnuFileExit")
        Me.mnuCategory.Caption = GetTxt("mnuCategory")
        Me.mnuCategoryAdd.Caption = GetTxt("mnuCategoryAdd")
        Me.mnuCategoryAddItem(0).Caption = GetTxt("mnuCategoryAddItemParentCat")
        Me.mnuCategoryAddItem(1).Caption = GetTxt("mnuCategoryAddItemChildCat")
        Me.mnuCategoryDelete.Caption = GetTxt("mnuCategoryDelete")
        Me.mnuCategoryEdit.Caption = GetTxt("mnuCategoryEdit")
        Me.mnuInfo.Caption = GetTxt("mnuInfo")
        '
        ReDim udtDUFCommands(0)
        ReDim udtAllowanceChanges(0)
        Set clsMyDUFInterpreter = New clsDUFInterpreter
        '
        ' Treeviews (opnieuw) inladen
        Call LoadTreeviews
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub treCategories_NodeClick(ByVal Node As MSComctlLib.Node)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.treCategories_NodeClick"
        '
        'De juiste werkwoorden checken
        '
        strSQL = "SELECT InfPr, Category FROM Verbes"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
        '
        With recVerbodrome
            blnUserEdit = False
            lstVerbes.Visible = False
            lstVerbes.Clear
            '
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    If chkDisplayChecked.Value = vbChecked Then
                        If InStr(!Category, Node.Key) Then
                            lstVerbes.AddItem Replace(!InfPr, "*", "")
                            lstVerbes.Selected(lstVerbes.ListCount - 1) = True
                        End If
                    Else
                        lstVerbes.AddItem Replace(!InfPr, "*", "")
                        lstVerbes.Selected(lstVerbes.ListCount - 1) = (InStr(!Category, Node.Key))
                    End If
                    .MoveNext
                Loop
            End If
            '
            Call DoAdditionalChanges
            lstVerbes.ListIndex = 0
            lstVerbes.Visible = True
            blnUserEdit = True
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
    Private Sub chkDisplayChecked_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.chkDisplayChecked_Click"
        '
        If treCategories.SelectedItem.Key <> "" Then
            If chkDisplayChecked.Value = vbChecked Then
                If lstVerbes.ListCount = 0 Then
                    GoTo ExitProc
                ElseIf lstVerbes.SelCount = 0 Then
                    lstVerbes.Clear
                    GoTo ExitProc
                Else
                    Dim i As Integer
                    '
                    For i = lstVerbes.ListCount - 1 To 0 Step -1
                        If Not (lstVerbes.Selected(i)) Then
                            lstVerbes.RemoveItem (i)
                        End If
                    Next i
                End If
            Else
                Call treCategories_NodeClick(treCategories.SelectedItem)
            End If
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
'        If chkDisplayChecked.Value = vbChecked Then
'            With lstVerbes
'                If lstVerbes.ListCount = 0 Then
'                    Exit Sub
'                ElseIf lstVerbes.SelCount = 0 Then
'                    lstVerbes.Clear
'                    Exit Sub
'                End If
'                '
'                lstVerbes.Visible = False
'                blnUserEdit = False
'                Call DoAdditionalChanges
'                .ListIndex = 0
'                '
'                Do
'                    If .Selected(.ListIndex) = False Then
'                        .RemoveItem .ListIndex
'                        '
'                        If .ListCount = 0 Then Exit Do
'                    Else
'                        If .ListIndex + 1 <> .ListCount Then
'                            .ListIndex = .ListIndex + 1
'                        Else
'                            Exit Do
'                        End If
'                    End If
'
'                Loop
'            End With
'            '
'            lstVerbes.ListIndex = 0
'            lstVerbes.Visible = True
'            blnUserEdit = True
'        Else
'            If treCategories.SelectedItem.Key <> "" Then
'                Call treCategories_NodeClick(treCategories.SelectedItem)
'            End If
'        End If
    End Sub
    '
    Private Sub lstVerbes_ItemCheck(Item As Integer)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.lstVerbes_ItemCheck"
        '
        If Len(treCategories.SelectedItem.Key) <= 1 And blnUserEdit Then
            MsgBox GetTxt("DUFCreator_NoCatSelected"), vbExclamation, AppName
            lstVerbes.Selected(Item) = Not (lstVerbes.Selected(Item))
        ElseIf Len(treCategories.SelectedItem.Key) > 1 And blnUserEdit Then
            'Commando: Individuele categorie veranderd
            strText = GetTxt(IIf(lstVerbes.Selected(Item), "DUFCreator_VerbCategoryAdd", "DUFCreator_VerbCategoryRemove"))
            strText = Replace(strText, "$infinitive$", lstVerbes.List(Item))
            strText = Replace(strText, "$category$", FullCategoryName)
            strDUFCommand = "Verb.Category." & IIf(lstVerbes.Selected(Item), "Add", "Remove") & "(" & lstVerbes.List(Item) & ")=" & treCategories.SelectedItem.Key
            Call AddDUFCommand(strDUFCommand, strText)
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub Form_Unload(Cancel As Integer)
        Cancel = vbCancel
        Call mnuFileExit_Click
    End Sub
'
' MENUS
'
    Private Sub mnuFileNew_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.mnuFileNew_Click"
        '
        If MsgBox(GetTxt("DUFCreator_NewFileAreYouSure"), vbQuestion + vbYesNo, AppName) = vbYes Then
            treCategories.Nodes.Clear
            lstCommands.Clear
            lstVerbes.Clear
            '
            Call LoadTreeviews
            ReDim udtDUFCommands(0)
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub mnuFileOpen_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.mnuFileOpen_Click"
        '
        With frmCommonDialog.dlgCommonDialog
            .DialogTitle = GetTxt("SelectDUF")
            .Filter = GetTxt("UpdateFile") & " (*" & gcDatabaseUpdateFileExtention & ")|*" & gcDatabaseUpdateFileExtention
            .CancelError = True
            .ShowOpen
            '
            If .FileName = "" Then GoTo ExitProc
            '
            treCategories.Nodes.Clear
            lstCommands.Clear
            lstVerbes.Clear
            '
            Call LoadTreeviews
            '
            Call clsMyDUFInterpreter.OpenFile(.FileName)
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If Err.Number = 32755 Then
            'Op Cancel gedrukt
            GoTo ExitProc
        Else
            If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
        End If
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub mnuFileSave_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.mnuFileSave_Click"
        '
        With frmCommonDialog.dlgCommonDialog
            .DialogTitle = GetTxt("SelectDUF")
            .Filter = GetTxt("UpdateFile") & " (*" & gcDatabaseUpdateFileExtention & ")|*" & gcDatabaseUpdateFileExtention
            .CancelError = True
            .ShowSave
            '
            If .FileName = "" Then GoTo ExitProc
            '
            If LCase(Dir(.FileName)) = LCase(Right(.FileName, InStr(StrReverse(.FileName), "\") - 1)) Then
                If MsgBox(GetTxt("WarningFileExists"), vbQuestion + vbYesNo, AppName) = vbNo Then GoTo ExitProc
            End If
            '
            Dim BNr: BNr = FreeFile
            Dim i As Integer
            '
            Open .FileName For Output As #BNr
            '
            For i = LBound(udtDUFCommands) To UBound(udtDUFCommands)
                Print #BNr, udtDUFCommands(i).DUFCommand
            Next i
            '
            Close #BNr
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If Err.Number = 32755 Then
            'Op Cancel gedrukt
            GoTo ExitProc
        Else
            If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
        End If
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub mnuFileExit_Click()
        If MsgBox(GetTxt("QuitProgram"), vbQuestion + vbYesNo, AppName) = vbYes Then
            End
        End If
    End Sub
    '
    Private Sub mnuCategoryAddItem_Click(Index As Integer)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.mnuCategoryAddItem_Click"
        '
        Dim strNewDescription As String
        '
        strNewDescription = InputBox(GetTxt(IIf(Index = 0, "DUFCreator_PromptNewParentCat", "DUFCreator_PromptNewChildCat")), AppName)
        '
        If strNewDescription <> "" Then
            'Er is iets (zinnigs) ingevuld
            Dim strNewKey As String: strNewKey = FindNewKeyName(Index = 0)
            '
            strText = GetTxt(IIf(Index = 0, "DUFCreator_CategoriesAdd_Parent", "DUFCreator_CategoriesAdd_Child"))
            strText = Replace(strText, "$category$", IIf(Index = 0, strNewDescription, FullCategoryName & "\" & strNewDescription))
            strDUFCommand = "Categories.Add " & strNewKey & "," & strNewDescription
            Call AddDUFCommand(strDUFCommand, strText)
            '
            treCategories.Nodes(strNewKey).EnsureVisible
            treCategories.Nodes(strNewKey).Selected = True
            Call treCategories_NodeClick(treCategories.SelectedItem)
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub mnuCategoryEdit_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.mnuCategoryEdit_Click"
        '
        'Categorie hernoemen
        Dim strNewDescription As String
        '
        If Not (treCategories.SelectedItem Is Nothing) Then
            strNewDescription = InputBox(Replace(GetTxt(IIf(Len(treCategories.SelectedItem.Key) = 1, "DUFCreator_PromptCategoriesDescription_Parent", "DUFCreator_PromptCategoriesDescription_Child")), "$category$", FullCategoryName), AppName, treCategories.SelectedItem.Text)
            '
            'Controleren of dit niet dezelfde categorie was
            If strNewDescription <> "" And strNewDescription <> treCategories.SelectedItem.Text Then
                'Andere naam
                '
                strText = GetTxt(IIf(Len(treCategories.SelectedItem.Key) = 1, "DUFCreator_CategoriesDescription_Parent", "DUFCreator_CategoriesDescription_Child"))
                
                strText = Replace(strText, "$oldname$", FullCategoryName)
                strText = Replace(strText, "$newname$", strNewDescription)
                strDUFCommand = "Categories.Description(" & treCategories.SelectedItem.Key & ")=" & strNewDescription
                Call AddDUFCommand(strDUFCommand, strText)
            End If
        Else
            MsgBox GetTxt("DUFCreator_NoCatSelected"), vbInformation, AppName
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub mnuCategoryDelete_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.mnuCategoryDelete_Click"
        '
        'Categorie verwijderen
        '
        If Not (treCategories.SelectedItem Is Nothing) Then
            If MsgBox(Replace(GetTxt(IIf(Len(treCategories.SelectedItem.Key) = 1, "DUFCreator_PromptCategoriesRemove_Parent", "DUFCreator_PromptCategoriesRemove_Child")), "$category$", FullCategoryName), vbYesNo + vbQuestion, AppName) = vbYes Then
                strText = GetTxt(IIf(Len(treCategories.SelectedItem.Key) = 1, "DUFCreator_CategoriesRemove_Parent", "DUFCreator_CategoriesRemove_Child"))
                strText = Replace(strText, "$category$", FullCategoryName)
                strDUFCommand = "Categories.Remove(" & treCategories.SelectedItem.Key & ")"
                Call AddDUFCommand(strDUFCommand, strText)
            End If
        Else
            MsgBox GetTxt("DUFCreator_NoCatSelected"), vbInformation, AppName
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub mnuInfo_Click()
        frmAbout.Show 1
    End Sub
'
' PROCEDURE HULP-FUNCTIES
'
    Private Sub AddDUFCommand(strDUFCommand As String, strText As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.AddDUFCommand"
        '
        If udtDUFCommands(0).Text <> "" Then ReDim Preserve udtDUFCommands(LBound(udtDUFCommands) To UBound(udtDUFCommands) + 1)
        '
        With udtDUFCommands(UBound(udtDUFCommands))
            .DUFCommand = strDUFCommand
            .Text = strText
            '
            lstCommands.AddItem .Text
            clsMyDUFInterpreter.Interprete (.DUFCommand)
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub AddAdditionalChange(strCategory As String, strInfinitive As String, blnAllowed As Boolean)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.AddAdditionalChange"
        '
        If udtAllowanceChanges(0).Category <> "" Then ReDim Preserve udtAllowanceChanges(LBound(udtAllowanceChanges) To UBound(udtAllowanceChanges) + 1)
        '
        With udtAllowanceChanges(UBound(udtAllowanceChanges))
            .Category = strCategory
            .Infinitive = strInfinitive
            .Allowed = blnAllowed
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub DoAdditionalChanges()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.DoAdditionalChanges"
        '
        Dim i As Integer, j As Integer
        '
        For i = LBound(udtAllowanceChanges) To UBound(udtAllowanceChanges)
            With udtAllowanceChanges(i)
                For j = 0 To lstVerbes.ListCount - 1
                    If treCategories.SelectedItem.Key = .Category And lstVerbes.List(j) = .Infinitive Then
                        lstVerbes.Selected(j) = .Allowed
                    End If
                Next j
            End With
        Next i
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Function FullCategoryName() As String
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.FullCategoryName"
        '
        If Len(treCategories.SelectedItem.Key) = 1 Then
            'Hoofdcategorie
            FullCategoryName = treCategories.SelectedItem.Text
        Else
            'Childcategorie
            FullCategoryName = treCategories.SelectedItem.Parent.Text & "\" & treCategories.SelectedItem.Text
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Private Function FindNewKeyName(blnRoot As Boolean) As String
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.FindNewKeyName"
        '
        Dim nodNode As Node
        Dim strAvailableKeys As String
        '
        If blnRoot Then
            strAvailableKeys = "ABCDEFGHIJKLMNOPQURSTUVWXYZ"
            '
            If Len(treCategories.SelectedItem.Key) = 1 Then
                'Root aangeduid
                Set nodNode = treCategories.SelectedItem.FirstSibling
            Else
                'Child aangeduid
                Set nodNode = treCategories.SelectedItem.Root.FirstSibling
            End If
            '
            Do
                strAvailableKeys = Replace(strAvailableKeys, nodNode.Key, "")
                If nodNode = nodNode.LastSibling Then
                    Exit Do
                Else
                    Set nodNode = nodNode.Next
                End If
            Loop
            '
            If Len(strAvailableKeys) > 0 Then
                FindNewKeyName = Left(strAvailableKeys, 1)
            Else
                MsgBox "Out of parent keys", vbExclamation, AppName
            End If
        Else
            Dim strParentKeyName As String
            '
            strAvailableKeys = "01,02,03,04,05,06,07,08,09,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99"
            '
            If Len(treCategories.SelectedItem.Key) = 1 Then
                'Root aangeduid
                If treCategories.SelectedItem.Children <> 0 Then
                    Set nodNode = treCategories.SelectedItem.Child.FirstSibling
                Else
                    strParentKeyName = treCategories.SelectedItem.Key
                End If
            Else
                'Child aangeduid
                Set nodNode = treCategories.SelectedItem.FirstSibling
            End If
            '
            If Not (nodNode Is Nothing) Then
                Do
                    strAvailableKeys = Replace(strAvailableKeys, Mid(nodNode.Key, 2), "")
                    If nodNode = nodNode.LastSibling Then
                        Exit Do
                    Else
                        Set nodNode = nodNode.Next
                    End If
                Loop
                '
                strParentKeyName = nodNode.Parent.Key
            End If
            '
            strAvailableKeys = Replace(strAvailableKeys, ",", "")
            '
            If Len(strAvailableKeys) > 0 Then
                FindNewKeyName = strParentKeyName & Left(strAvailableKeys, 2)
            Else
                MsgBox "Out of child keys", vbExclamation, AppName
            End If
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Private Sub LoadTreeviews()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.LoadTreeviews"
        '
        '1. treCategories: Hoofdcategoriën toevoegen
        strSQL = "SELECT * FROM Category WHERE Len(CatName)=1"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
        With recVerbodrome
            If Not (.EOF And .BOF) Then
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
                While Not .EOF
                    Me.treCategories.Nodes.Add Left(!CatName, 1), tvwChild, !CatName, !CatDescription, "Category"
                    .MoveNext
                Wend
                .MoveFirst
                treCategories.SelectedItem = treCategories.Nodes(CStr(!CatName)).Root.FirstSibling
            Else
                .Close
                GoTo ExitProc
            End If
            '
            .Close
        End With
        '
        '3. Werkwoordenlijst inladen met de eerste node
        Call treCategories_NodeClick(treCategories.SelectedItem)
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    ' Interpretatiegedeelte van de DUF commandos
    '
    Private Sub clsMyDUFInterpreter_CategoriesAdd(strNewKey As String, strNewDescription As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.clsMyDUFInterpreter_CategoriesAdd"
        '
        'Categorie toevoegen
        If Len(strNewKey) = 1 Then
            'Hoofdcategorie
            treCategories.Nodes.Add , , strNewKey, strNewDescription, IIf(strNewKey = "B", "Bescherelle", "Leraar")
        Else
            'Categorie
            treCategories.Nodes.Add Left(strNewKey, 1), tvwChild, strNewKey, strNewDescription, "Category"
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub clsMyDUFInterpreter_CategoriesDescription(strKey As String, strNewDescription As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.clsMyDUFInterpreter_CategoriesDescription"
        '
        'Categorie hernoemen
        treCategories.Nodes(strKey).Text = strNewDescription
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub clsMyDUFInterpreter_CategoriesRemove(strKey As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.clsMyDUFInterpreter_CategoriesRemove"
        '
        'Categorie verwijderen
        Call treCategories.Nodes.Remove(strKey)
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub clsMyDUFInterpreter_VerbCategoryAdd(strInfinitive As String, strAddedToCategory As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.clsMyDUFInterpreter_VerbCategoryAdd"
        '
        'Werkwoord aan een categorie toevoegen
        Dim i As Integer
        '
        For i = 0 To lstVerbes.ListCount - 1
            If lstVerbes.List(i) = strInfinitive Then
                Call AddAdditionalChange(strAddedToCategory, strInfinitive, True)
                Exit For
            End If
        Next i
        '
        If treCategories.SelectedItem.Key = strAddedToCategory Then
            blnUserEdit = False
            Call DoAdditionalChanges
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
    Private Sub clsMyDUFInterpreter_VerbCategoryRemove(strInfinitive As String, strRemovedFromCategory As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmCreatorDUF.clsMyDUFInterpreter_VerbCategoryRemove"
        '
        'Werkwoord verwijderen uit een categorie
        Dim i As Integer
        '
        For i = 0 To lstVerbes.ListCount - 1
            If lstVerbes.List(i) = strInfinitive Then
                Call AddAdditionalChange(strRemovedFromCategory, strInfinitive, False)
                Exit For
            End If
        Next i
        '
        If treCategories.SelectedItem.Key = strRemovedFromCategory Then
            blnUserEdit = False
            Call DoAdditionalChanges
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
    Private Sub clsMyDUFInterpreter_UnknownCommand(strCommandLine As String)
        'Onbekend commando
        MsgBox Replace(GetTxt("CommandUnknown"), "$command$", strCommandLine), vbExclamation, AppName
    End Sub
