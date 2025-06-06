VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDUFInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
'DUF 2.0 - Interpretatieklasse
'
Public Event VerbCategoryAdd(strInfinitive As String, strAddedToCategory As String)
Public Event VerbCategoryRemove(strInfinitive As String, strRemovedFromCategory As String)
'
Public Event CategoriesAdd(strNewKey As String, strNewDescription As String)
Public Event CategoriesDescription(strKey As String, strNewDescription As String)
Public Event CategoriesRemove(strKey As String)
'
Public Event UnknownCommand(strCommandLine As String)
'
' SUBS
'
    Public Sub OpenFile(strDUFFile As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsDUFInterpreter.OpenFile"
        '
        Dim strInstructions() As String: ReDim strInstructions(0)
        Dim BNr As Integer: BNr = FreeFile
        '
        Open strDUFFile For Input As #BNr
        '
        frmLoad.Show
        Screen.MousePointer = vbHourglass
        '
        frmLoad.lblCurrentJob.Caption = GetTxt("DUFInstaller_InterpretingCommands")
        DoEvents
        '
        While Not EOF(BNr)
            If strInstructions(UBound(strInstructions)) <> "" And Left(strInstructions(UBound(strInstructions)), 1) <> "'" Then
                ReDim Preserve strInstructions(LBound(strInstructions) To UBound(strInstructions) + 1)
            End If
            '
            Line Input #BNr, strInstructions(UBound(strInstructions))
        Wend
        '
        Dim i As Integer
        '
        frmLoad.prgDatabase.Max = IIf(UBound(strInstructions) = 0, 1, UBound(strInstructions))
        '
        For i = LBound(strInstructions) To UBound(strInstructions)
            frmLoad.prgDatabase.Value = i
            Call Interprete(strInstructions(i))
        Next i
        '
        Screen.MousePointer = vbDefault
        Unload frmLoad
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Public Sub Interprete(strInstruction As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsDUFInterpreter.Interprete"
        '
        Dim strInfinitive As String, strAddedToCategory As String, strRemovedFromCategory As String, strNewKey As String, strKey As String, strNewDescription As String
        '
        If Left(strInstruction, Len("Verb.Category.Add")) = "Verb.Category.Add" Then
            'Een categorie aan een werkwoord toevoegen
            '    Verb.Category.Add(strInfinitive) = strAddedToCategory
            strInstruction = Replace(strInstruction, " (", "(")
            strInstruction = Replace(strInstruction, "Verb.Category.Add(", "")
            strInstruction = Replace(strInstruction, " = ", "=")
            strInstruction = Replace(strInstruction, ")", "")
            '
            strInfinitive = Left(strInstruction, InStr(strInstruction, "=") - 1)
            strAddedToCategory = Mid(strInstruction, InStr(strInstruction, "=") + 1)
            '
            RaiseEvent VerbCategoryAdd(strInfinitive, strAddedToCategory)
        ElseIf Left(strInstruction, Len("Verb.Category.Remove")) = "Verb.Category.Remove" Then
            'Een categorie van een werkwoord verwijderen
            '    Verb.Category.Remove(strInfinitive) = strRemovedFromCategory
            strInstruction = Replace(strInstruction, " (", "(")
            strInstruction = Replace(strInstruction, "Verb.Category.Remove(", "")
            strInstruction = Replace(strInstruction, " = ", "=")
            strInstruction = Replace(strInstruction, ")", "")
            '
            strInfinitive = Left(strInstruction, InStr(strInstruction, "=") - 1)
            strRemovedFromCategory = Mid(strInstruction, InStr(strInstruction, "=") + 1)
            '
            RaiseEvent VerbCategoryRemove(strInfinitive, strRemovedFromCategory)
        ElseIf Left(strInstruction, Len("Categories.Add")) = "Categories.Add" Then
            'Categorie toevoegen
            '    Categories.Add strNewKey, strNewDescription
            strInstruction = Replace(strInstruction, "Categories.Add ", "")
            strInstruction = Replace(strInstruction, ", ", ",")
            '
            strNewKey = Left(strInstruction, InStr(strInstruction, ",") - 1)
            strNewDescription = Mid(strInstruction, InStr(strInstruction, ",") + 1)
            '
            RaiseEvent CategoriesAdd(strNewKey, strNewDescription)
        ElseIf Left(strInstruction, Len("Categories.Description")) = "Categories.Description" Then
            'Categorie een nieuwe beschrijving geven
            '    Categories.Description(strKey) = strNewDescription
            strInstruction = Replace(strInstruction, " (", "(")
            strInstruction = Replace(strInstruction, "Categories.Description(", "")
            strInstruction = Replace(strInstruction, " = ", "=")
            strInstruction = Replace(strInstruction, ")", "")
            '
            strKey = Left(strInstruction, InStr(strInstruction, "=") - 1)
            strNewDescription = Mid(strInstruction, InStr(strInstruction, "=") + 1)
            '
            RaiseEvent CategoriesDescription(strKey, strNewDescription)
        ElseIf Left(strInstruction, Len("Categories.Remove")) = "Categories.Remove" Then
            'Categorie verwijderen
            '    Categories.Remove(strKey)
            strInstruction = Replace(strInstruction, " (", "(")
            strInstruction = Replace(strInstruction, "Categories.Remove(", "")
            strInstruction = Replace(strInstruction, ")", "")
            '
            strKey = strInstruction
            '
            RaiseEvent CategoriesRemove(strKey)
        Else
            RaiseEvent UnknownCommand(strInstruction)
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
