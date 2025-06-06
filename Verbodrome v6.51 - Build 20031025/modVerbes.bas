Attribute VB_Name = "modVerbes"
Option Explicit
'
' SUBS
'
    Public Sub LoadVerbs()
        On Error GoTo ErrorOccured
        errProcStack errPush, "modVerbes.LoadVerbs"
        '
        Dim strPersonne() As String
        Dim intVerbIndex As Integer
        Dim i As Byte, j As Byte
        intVerbIndex = 1
        '
        Screen.MousePointer = vbHourglass
        '
        strSQL = "SELECT InfPr FROM Verbes"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
        '
        With recVerbodrome
            frmLoad.prgDatabase.Max = .RecordCount
            ReDim blnIsVerbAllowed(1 To .RecordCount)
            '
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    Call myFormes.AddInfinitive(!InfPr, intVerbIndex)
                    '
                    intVerbIndex = intVerbIndex + 1
                    frmLoad.prgDatabase.Value = frmLoad.prgDatabase.Value + 1
                    .MoveNext
                Loop
            End If
            '
            .Close
        End With
        '
        frmLoad.lblCurrentJob = GetTxt("ApplyingSettings")
        DoEvents
        '
        Screen.MousePointer = vbDefault
        '
        Call LoadPermissions
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Public Sub LoadPermissions()
        On Error GoTo ErrorOccured
        errProcStack errPush, "modVerbes.GetPermissions"
        '
        Dim i As Integer: i = 1
        '
        Screen.MousePointer = vbHourglass
        '
        'Variabelen resizen, en eventueel reeds opgeslagen informatie verwijderen
        ReDim intAllowedVerbes(0): intAllowedVerbes(0) = -1
        ReDim bytAllowedTemps(0): bytAllowedTemps(0) = 0
        ReDim bytAllowedPersonnes(0): bytAllowedPersonnes(0) = 0
        '
        'Toelatingen opzoeken
        'Werkwoordstoelatingen
        strSQL = "SELECT InfPr, Allowed FROM Verbes"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockReadOnly
        '
        With recVerbodrome
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    'Toelatingen wegschrijven
                    If !Allowed = True Then
                        If intAllowedVerbes(0) <> -1 Then ReDim Preserve intAllowedVerbes(0 To UBound(intAllowedVerbes) + 1)
                        intAllowedVerbes(UBound(intAllowedVerbes)) = i
                    End If
                    blnIsVerbAllowed(i) = !Allowed
                    '
                    i = i + 1
                    .MoveNext
                Loop
            End If
            .Close
        End With
        '
        'Tijds- en persoonstoelatingen
        strSQL = "SELECT Name, Allowed FROM Allowance"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockReadOnly
        '
        With recVerbodrome
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    If Len(!Name) = 2 Then
                        'Persoon
                        Select Case !Name
                            Case "1E": i = 1
                            Case "2E": i = 2
                            Case "3E": i = 3
                            Case "1M": i = 4
                            Case "2M": i = 5
                            Case "3M": i = 6
                        End Select
                        '
                        blnIsPersonneAllowed(i) = !Allowed
                        '
                        If !Allowed = True Then
                            'Persoon toegelaten
                            If bytAllowedPersonnes(0) <> 0 Then ReDim Preserve bytAllowedPersonnes(0 To UBound(bytAllowedPersonnes) + 1)
                            bytAllowedPersonnes(UBound(bytAllowedPersonnes)) = i
                        End If
                    Else
                        'Tijd
                        blnIsTempsAllowed(Convert(Abreviation, AbsNumber, !Name)) = !Allowed
                        '
                        If !Allowed = True Then
                            If bytAllowedTemps(0) <> 0 Then ReDim Preserve bytAllowedTemps(0 To UBound(bytAllowedTemps) + 1)
                            bytAllowedTemps(UBound(bytAllowedTemps)) = Convert(Abreviation, AbsNumber, !Name)
                        End If
                    End If
                    '
                    .MoveNext
                Loop
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

