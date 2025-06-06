VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDUFInstaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
' DUF 2.0
'
Private WithEvents clsMyDUF As clsDUFInterpreter
Attribute clsMyDUF.VB_VarHelpID = -1
Private blnErrorOccured As Boolean
'
' SUBS
'
    ' Constructor
    Private Sub Class_Initialize()
        Set clsMyDUF = New clsDUFInterpreter
    End Sub
    '
    Public Sub InstallDUF(strDUFFile As String)
        clsMyDUF.OpenFile (strDUFFile)
        '
        Call LogEvent(DUFSuccesful, "")
        If Not (blnErrorOccured) Then
            'Succesvol
            MsgBox GetTxt("DUFInstaller_Succesfull"), vbInformation, AppName
        Else
            'Een of meerdere fouten
            MsgBox GetTxt("DUFInstaller_Error"), vbExclamation, AppName
        End If
    End Sub
    '
    Private Sub clsMyDUF_CategoriesAdd(strNewKey As String, strNewDescription As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsDUFInstaller.clsMyDUF_CategoriesAdd"
        '
        strSQL = "SELECT CatName, CatDescription FROM Category WHERE CatName='" & strNewKey & "'"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
        '
        With recVerbodrome
            If (.EOF And .BOF) Then
                .AddNew
                !CatName = strNewKey
                !CatDescription = strNewDescription
                .Update
            Else
                .Close
                blnErrorOccured = True
                MsgBox GetTxt("DUFInstaller_ErrorAddingCategory"), vbExclamation, AppName
                GoTo ExitProc
            End If
            .Close
        End With
        '
        GoTo ExitProc
ErrorOccured::
        blnErrorOccured = True
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub clsMyDUF_CategoriesDescription(strKey As String, strNewDescription As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsDUFInstaller.clsMyDUF_CategoriesDescription"
        '
        strSQL = "SELECT CatName, CatDescription FROM Category WHERE CatName='" & strKey & "'"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
        '
        With recVerbodrome
            If Not (.EOF And .BOF) Then
                !CatDescription = strNewDescription
                .Update
            Else
                .Close
                blnErrorOccured = True
                MsgBox GetTxt("DUFInstaller_ErrorEditingCategory"), vbExclamation, AppName
                GoTo ExitProc
            End If
            .Close
        End With
        '
        GoTo ExitProc
ErrorOccured::
        blnErrorOccured = True
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub clsMyDUF_CategoriesRemove(strKey As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsDUFInstaller.clsMyDUF_CategoriesRemove"
        '
        strSQL = "DELETE * FROM Category WHERE Left(CatName, " & Len(strKey) & ")='" & strKey & "'"
        gcnConnection.Execute strSQL
        '
        strSQL = "SELECT InfPr, Category FROM Verbes WHERE Instr(Category, '" & strKey & "')"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
        '
        Dim strOldCat As String
        '
        With recVerbodrome
            If Not (.BOF And .EOF) Then
                While Not .EOF
                    strOldCat = Mid(!Category, InStr(!Category, strKey))
                    If InStr(strOldCat, ",") Then strOldCat = Left(strOldCat, InStr(strOldCat, ",") - 1)
                    !Category = Replace(!Category, strOldCat, "")
                    !Category = Replace(!Category, ",", " ")
                    !Category = Trim(!Category)
                    !Category = Replace(!Category, "  ", ",")
                    !Category = Replace(!Category, " ", ",")
                    '
                    .Update
                    .MoveNext
                Wend
            End If
            .Close
        End With
        '
        GoTo ExitProc
ErrorOccured::
        blnErrorOccured = True
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub clsMyDUF_VerbCategoryAdd(strInfinitive As String, strAddedToCategory As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsDUFInstaller.clsMyDUF_VerbCategoryAdd"
        '
        strSQL = "SELECT InfPr, Category FROM Verbes WHERE Left(InfPr, " & Len(strInfinitive) & ")='" & strInfinitive & "'"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
        '
        With recVerbodrome
            If Not (.BOF And .EOF) Then
                !Category = !Category & " " & strAddedToCategory
                !Category = Trim(!Category)
                !Category = Replace(!Category, "  ", ",")
                !Category = Replace(!Category, " ", ",")
                .Update
            Else
                .Close
                blnErrorOccured = True
                MsgBox GetTxt("DUFInstaller_ErrorAddingVerbToCategory"), vbExclamation, AppName
                GoTo ExitProc
            End If
            .Close
        End With
        '
        GoTo ExitProc
ErrorOccured::
        blnErrorOccured = True
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub clsMyDUF_VerbCategoryRemove(strInfinitive As String, strRemovedFromCategory As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsDUFInstaller.clsMyDUF_VerbCategoryRemove"
        '
        strSQL = "SELECT InfPr, Category FROM Verbes WHERE Left(InfPr, " & Len(strInfinitive) & ")='" & strInfinitive & "'"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
        '
        With recVerbodrome
            If Not (.BOF And .EOF) Then
                !Category = Replace(!Category, strRemovedFromCategory, "")
                !Category = Replace(!Category, ",", " ")
                !Category = Trim(!Category)
                !Category = Replace(!Category, " ", ",")
                .Update
            Else
                .Close
                MsgBox GetTxt("DUFInstaller_ErrorRemovingVerbFromCategory"), vbExclamation, AppName
                GoTo ExitProc
            End If
            .Close
        End With
        '
        GoTo ExitProc
ErrorOccured::
        blnErrorOccured = True
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub clsMyDUF_UnknownCommand(strCommandLine As String)
        'Negeren
    End Sub
