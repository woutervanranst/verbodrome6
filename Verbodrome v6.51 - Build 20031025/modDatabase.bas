Attribute VB_Name = "modDatabase"
Option Explicit
'
' SUBS
'
    Public Sub LetDBSetting(Setting As String, Value As String)
        On Error GoTo ErrorOccured
        errProcStack errPush, "modDatabase.LetDBSetting"
        If blnPM_DebugLog Then Call LogDebug("Setting: " & Setting & ", Value: " & Value)
        '
        Dim recSettings As New ADODB.Recordset
        Dim strSettings As String
        '
        strSettings = "SELECT * FROM Settings WHERE Setting='" & Setting & "'"
        recSettings.Open strSettings, gcConnectionString, adOpenStatic, adLockOptimistic
        '
        With recSettings
            If Not (.EOF And .BOF) Then
                .MoveFirst
                !Value = Value
            Else
                .AddNew
                !Setting = Setting
                !Value = Value
            End If
            .Update
        End With
        '
        Call LogEvent(ChangeSetting, Setting & "$" & Value, True)
        '
        Set recSettings = Nothing
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Public Sub LogEvent(bytEventType As EventType, Optional strParam As String, Optional blnOverride As Boolean)
        On Error Resume Next
        If blnPM_DebugLog Then Call LogDebug("Logging event: " & bytEventType & IIf(strParam <> "", ", Parameterlist: " & strParam, ""))
        '
        If (bytEventType <> DebugLogFile And blnHistoryEnabled) Or blnOverride Then
            Dim recHistory As New ADODB.Recordset
            Dim strSQLHistory As String
            '
            strSQLHistory = "SELECT * FROM History"
            recHistory.Open strSQLHistory, gcConnectionString, adOpenStatic, adLockOptimistic
            '
            With recHistory
                .AddNew
                !EventType = bytEventType
                '
                'Als er met een tijdslimiet wordt gewerkt, dan telt de teller af, en niet op
                If bytEventType = CorrectAnswer Or bytEventType = WrongAnswer Then
                    If mySession.TimeLimitMode Then
                    'If frmMain.mnuFileSettingsTimeLimitStatusOn.Checked Then
                        Dim strParameters() As String
                        strParameters = Split(strParam, "$")
                        strParameters(1) = CStr(mySession.TimeLimit - mySession.CurrentQuestionTime) & "s"
                    End If
                End If
                '
                If strParam <> "" Then !Parameters = strParam
                .Update
                .Close
                '
                Call GetSessionID
            End With
            '
            Set recHistory = Nothing
        End If
    End Sub
    '
    Public Sub EmptyHistory()
        On Error GoTo ErrorOccured
        errProcStack errPush, "modDatabase.EmptyHistory"
        If blnPM_DebugLog Then Call LogDebug("History emptied")
        '
        'Functie die alle historiekgegevens buiten deze van de huidige sessie, wist
        strSQL = "DELETE * FROM History WHERE Id<" & lngSessionID
        gcnConnection.Execute strSQL
        Call LogEvent(Error, "HistoryEmpty")
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Public Sub GetSessionID()
        On Error Resume Next
        '
        If lngSessionID = 0 Then
            Dim recHistory As New ADODB.Recordset
            Dim strSQLHistory As String
            '
            strSQLHistory = "SELECT Id FROM History"
            recHistory.Open strSQLHistory, gcConnectionString, adOpenDynamic, adLockReadOnly
            '
            If Not (recHistory.BOF And recHistory.EOF) Then
                recHistory.MoveLast
                lngSessionID = recHistory!Id
            Else
                lngSessionID = 0
            End If
            '
            recHistory.Close
        End If
    End Sub
'
' FUNCTIONS
'
    Public Function GetTxt(strId As String) As String
        On Error GoTo ErrorOccured
        errProcStack errPush, "modDatabase.GetTxt"
        If blnPM_DebugLog Then Call LogDebug("ID: " & strId)
        '
        Dim recLanguage As New ADODB.Recordset
        Dim strLang As String
        '
        strLang = "SELECT * FROM Language WHERE Id='" & strId & "'"
        recLanguage.Open strLang, gcConnectionString, adOpenStatic, adLockReadOnly
        '
        With recLanguage
            If Not (.EOF And .BOF) Then
                '.MoveFirst
                Select Case bytLanguage
                    Case 0 'Nederlands
                        GetTxt = !Nederlands
                    Case 1 'Frans
                        GetTxt = !Français
                    Case 2 'Engels
                        GetTxt = !English
                End Select
'            Else
'                .Close
'                .Open strLang, gcConnectionString, adOpenDynamic, adLockOptimistic
'                If Not (blnPM_Debug) Then
'                    MsgBox strId
'                    GetTxt = strId
'                Else
'                    .AddNew
'                    !Id = strId
'                    GetTxt = InputBox("Nederlandse omschrijving voor " & strId, AppName)
'                    !Nederlands = GetTxt
'                    .Update
'                End If
            End If
        End With
        '
        GetTxt = Replace(GetTxt, "$enter$", vbCrLf)
        GetTxt = Replace(GetTxt, "$tab$", vbTab)
        GetTxt = Replace(GetTxt, "…", "...")
        '
        Set recLanguage = Nothing
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Public Function GetDBSetting(Setting As String, Optional Default As String) As String
        On Error GoTo ErrorOccured
        errProcStack errPush, "modDatabase.GetDBSetting"
        If blnPM_DebugLog Then Call LogDebug("Setting: " & Setting & IIf(Default <> "", ", Default: " & Default, ""))
        '
        ' TODO: Controleren of voor elke LOGDebug nen if staat
        Dim recSettings As New ADODB.Recordset
        Dim strSettings As String
        '
        strSettings = "SELECT * FROM Settings WHERE Setting='" & Setting & "'"
        recSettings.Open strSettings, gcConnectionString, adOpenStatic, adLockReadOnly
        '
        With recSettings
            If Not (.EOF And .BOF) Then
                .MoveFirst
                GetDBSetting = !Value
            Else
                GetDBSetting = Default
            End If
        End With
        '
        GetDBSetting = Replace(GetDBSetting, "$app.path$", App.Path)
        '
        Set recSettings = Nothing
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Public Function Convert(bytFrom As ConvertType, bytTo As ConvertType, vntConvertWhat As Variant) As Variant
        'On Error GoTo ErrorOccured
        'errProcStack errPush, "modDatabase.Convert"
        '
        Select Case bytFrom
            Case ConvertType.AbsNumber  'Absoluut Nummer
                Select Case bytTo
                    Case ConvertType.Abreviation 'Afkorting
                        Select Case CByte(vntConvertWhat)
                            Case 1:     Convert = "InfPr"
                            Case 2:     Convert = "InfPa"
                            Case 3:     Convert = "IndPr"
                            Case 4:     Convert = "IndPC"
                            Case 5:     Convert = "IndIpf"
                            Case 6:     Convert = "IndPQP"
                            Case 7:     Convert = "IndPS"
                            Case 8:     Convert = "IndPA"
                            Case 9:     Convert = "IndFS"
                            Case 10:    Convert = "IndFA"
                            Case 11:    Convert = "CondPr"
                            Case 12:    Convert = "CondPa"
                            Case 13:    Convert = "SubjPr"
                            Case 14:    Convert = "SubjPa"
                            Case 15:    Convert = "SubjIpf"
                            Case 16:    Convert = "SubjPQP"
                            Case 17:    Convert = "ImpPr"
                            Case 18:    Convert = "ImpPa"
                            Case 19:    Convert = "PartPr"
                            Case 20:    Convert = "PartPa"
                            Case 21:    Convert = "GerPr"
                            Case 22:    Convert = "GerPa"
                        End Select
                    Case ConvertType.Full  'Volledig
                        Select Case CByte(vntConvertWhat)
                            Case 1:     Convert = "Infinitif Présent"
                            Case 2:     Convert = "Infinitif Passé"
                            Case 3:     Convert = "Indicatif Présent"
                            Case 4:     Convert = "Indicatif Passé Composé"
                            Case 5:     Convert = "Indicatif Imparfait"
                            Case 6:     Convert = "Indicatif Plus-que-parfait"
                            Case 7:     Convert = "Indicatif Passé Simple"
                            Case 8:     Convert = "Indicatif Passé Antérieur"
                            Case 9:     Convert = "Indicatif Futur Simple"
                            Case 10:    Convert = "Indicatif Futur Antérieur"
                            Case 11:    Convert = "Conditionnel Présent"
                            Case 12:    Convert = "Conditionnel Passé"
                            Case 13:    Convert = "Subjonctif Présent"
                            Case 14:    Convert = "Subjonctif Passé"
                            Case 15:    Convert = "Subjonctif Imparfait"
                            Case 16:    Convert = "Subjonctif Plus-que-parfait"
                            Case 17:    Convert = "Impératif Présent"
                            Case 18:    Convert = "Impératif Passé"
                            Case 19:    Convert = "Participe Présent"
                            Case 20:    Convert = "Participe Passé"
                            Case 21:    Convert = "Gérondif Présent"
                            Case 22:    Convert = "Gérondif Passé"
                        End Select
                    Case ConvertType.ShortFull  'Afgekort volledig
                        Select Case CByte(vntConvertWhat)
                            Case 1:     Convert = "Inf. Prés."
                            Case 2:     Convert = "Inf. Passé"
                            Case 3:     Convert = "Ind. Prés."
                            Case 4:     Convert = "Ind. Passé Comp."
                            Case 5:     Convert = "Ind. Imparf."
                            Case 6:     Convert = "Ind. Plus-Que-Parfait"
                            Case 7:     Convert = "Ind. Passé Simple"
                            Case 8:     Convert = "Ind. Passé Ant."
                            Case 9:     Convert = "Ind. Fut. Simp."
                            Case 10:    Convert = "Ind. Fut. Ant."
                            Case 11:    Convert = "Cond. Prés."
                            Case 12:    Convert = "Cond. Passé"
                            Case 13:    Convert = "Subj. Prés."
                            Case 14:    Convert = "Subj. Passé"
                            Case 15:    Convert = "Subj. Imparf."
                            Case 16:    Convert = "Subj. Plus-Que-Parfait"
                            Case 17:    Convert = "Impérat. Prés."
                            Case 18:    Convert = "Impérat. Passé"
                            Case 19:    Convert = "Part. Prés."
                            Case 20:    Convert = "Part. Passé"
                            Case 21:    Convert = "Gérondif Présent"
                            Case 22:    Convert = "Gérondif Passé"
                        End Select
                End Select
            Case ConvertType.Abreviation  'Afkorting
                Select Case bytTo
                    Case ConvertType.AbsNumber  'Absoluut Nummer
                        Select Case CStr(vntConvertWhat)
                            Case "InfPr":   Convert = gcInfPr
                            Case "InfPa":   Convert = gcInfPa
                            Case "IndPr":   Convert = gcIndPr
                            Case "IndPC":   Convert = gcIndPC
                            Case "IndIpf":  Convert = gcIndIpf
                            Case "IndPQP":  Convert = gcIndPQP
                            Case "IndPS":   Convert = gcIndPS
                            Case "IndPA":   Convert = gcIndPA
                            Case "IndFS":   Convert = gcIndFS
                            Case "IndFA":   Convert = gcIndFA
                            Case "CondPr":  Convert = gcCondPr
                            Case "CondPa":  Convert = gcCondPa
                            Case "SubjPr":  Convert = gcSubjPr
                            Case "SubjPa":  Convert = gcSubjPa
                            Case "SubjIpf": Convert = gcSubjIpf
                            Case "SubjPQP": Convert = gcSubjPQP
                            Case "ImpPr":   Convert = gcImpPr
                            Case "ImpPa":   Convert = gcImpPa
                            Case "PartPr":  Convert = gcPartPr
                            Case "PartPa":  Convert = gcPartPa
                            Case "GerPr":   Convert = gcGerPr
                            Case "GerPa":   Convert = gcGerPa
                        End Select
                    Case ConvertType.Full  'Volledig
                        Select Case CStr(vntConvertWhat)
                            Case "InfPr":   Convert = "Infinitif Présent"
                            Case "InfPa":   Convert = "Infinitif Passé"
                            Case "IndPr":   Convert = "Indicatif Présent"
                            Case "IndPC":   Convert = "Indicatif Passé Composé"
                            Case "IndIpf":  Convert = "Indicatif Imparfait"
                            Case "IndPQP":  Convert = "Indicatif Plus-que-parfait"
                            Case "IndPS":   Convert = "Indicatif Passé Simple"
                            Case "IndPA":   Convert = "Indicatif Passé Antérieur"
                            Case "IndFS":   Convert = "Indicatif Futur Simple"
                            Case "IndFA":   Convert = "Indicatif Futur Antérieur"
                            Case "CondPr":  Convert = "Conditionnel Présent"
                            Case "CondPa":  Convert = "Conditionnel Passé"
                            Case "SubjPr":  Convert = "Subjonctif Présent"
                            Case "SubjPa":  Convert = "Subjonctif Passé"
                            Case "SubjIpf": Convert = "Subjonctif Imparfait"
                            Case "SubjPQP": Convert = "Subjonctif Plus-que-parfait"
                            Case "ImpPr":   Convert = "Impératif Présent"
                            Case "ImpPa":   Convert = "Impératif Passé"
                            Case "PartPr":  Convert = "Participe Présent"
                            Case "PartPa":  Convert = "Participe Passé"
                            Case "GerPr":   Convert = "Gérondif Présent"
                            Case "GerPa":   Convert = "Gérondif Passé"
                        End Select
                End Select
            Case ConvertType.Full  'Volledig
                Select Case bytTo
                    Case ConvertType.AbsNumber  'Absoluut Nummer
                        Select Case CStr(vntConvertWhat)
                            Case "Infinitif Présent":           Convert = gcInfPr
                            Case "Infinitif Passé":             Convert = gcInfPa
                            Case "Indicatif Présent":           Convert = gcIndPr
                            Case "Indicatif Passé Composé":     Convert = gcIndPC
                            Case "Indicatif Imparfait":         Convert = gcIndIpf
                            Case "Indicatif Plus-que-parfait":  Convert = gcIndPQP
                            Case "Indicatif Passé Simple":      Convert = gcIndPS
                            Case "Indicatif Passé Antérieur":   Convert = gcIndPA
                            Case "Indicatif Futur Simple":      Convert = gcIndFS
                            Case "Indicatif Futur Antérieur":   Convert = gcIndFA
                            Case "Conditionnel Présent":        Convert = gcCondPr
                            Case "Conditionnel Passé":          Convert = gcCondPa
                            Case "Subjonctif Présent":          Convert = gcSubjPr
                            Case "Subjonctif Passé":            Convert = gcSubjPa
                            Case "Subjonctif Imparfait":        Convert = gcSubjIpf
                            Case "Subjonctif Plus-que-parfait": Convert = gcSubjPQP
                            Case "Impératif Présent":           Convert = gcImpPr
                            Case "Impératif Passé":             Convert = gcImpPa
                            Case "Participe Présent":           Convert = gcPartPr
                            Case "Participe Passé":             Convert = gcPartPa
                            Case "Gérondif Présent":            Convert = gcGerPr
                            Case "Gérondif Passé":              Convert = gcGerPa
                        End Select
                    Case ConvertType.Abreviation  'Afkorting
                        Select Case CStr(vntConvertWhat)
                            Case "Infinitif Présent":           Convert = "InfPr"
                            Case "Infinitif Passé":             Convert = "InfPa"
                            Case "Indicatif Présent":           Convert = "IndPr"
                            Case "Indicatif Passé Composé":     Convert = "IndPC"
                            Case "Indicatif Imparfait":         Convert = "IndIpf"
                            Case "Indicatif Plus-que-parfait":  Convert = "IndPQP"
                            Case "Indicatif Passé Simple":      Convert = "IndPS"
                            Case "Indicatif Passé Antérieur":   Convert = "IndPA"
                            Case "Indicatif Futur Simple":      Convert = "IndFS"
                            Case "Indicatif Futur Antérieur":   Convert = "IndFA"
                            Case "Conditionnel Présent":        Convert = "CondPr"
                            Case "Conditionnel Passé":          Convert = "CondPa"
                            Case "Subjonctif Présent":          Convert = "SubjPr"
                            Case "Subjonctif Passé":            Convert = "SubjPa"
                            Case "Subjonctif Imparfait":        Convert = "SubjIpf"
                            Case "Subjonctif Plus-que-parfait": Convert = "SubjPQP"
                            Case "Impératif Présent":           Convert = "ImpPr"
                            Case "Impératif Passé":             Convert = "ImpPa"
                            Case "Participe Présent":           Convert = "PartPr"
                            Case "Participe Passé":             Convert = "PartPa"
                            Case "Gérondif Présent":            Convert = "GerPr"
                            Case "Gérondif Passé":              Convert = "GerPa"
                        End Select
                End Select
        End Select
'        '
'        GoTo ExitProc
'ErrorOccured::
'        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
'ExitProc::
'        errProcStack errPop, ""
    End Function
    '
    Public Function NumberOfPersons(bytAbsoluteTempsIndex As Byte) As Byte
        On Error GoTo ErrorOccured
        errProcStack errPush, "modDatabase.NumberOfPersons"
        '
        Select Case bytAbsoluteTempsIndex
            Case gcIndPr, gcIndPC, gcIndIpf, gcIndPQP, gcIndPS, gcIndPA, gcIndFS, gcIndFA, gcCondPr, gcCondPa, gcSubjPr, gcSubjPa, gcSubjIpf, gcSubjPQP '3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
                '6 personen
                NumberOfPersons = 6
            Case gcImpPr, gcImpPa
                '3 personen
                NumberOfPersons = 3
            Case gcPartPa
                '2 personen
                NumberOfPersons = 2
            Case gcInfPr, gcInfPa, gcPartPr, gcGerPr, gcGerPa
                '1 persoon
                NumberOfPersons = 1
        End Select
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Public Function NumberOfHistoryItems() As Long
        On Error GoTo ErrorOccured
        errProcStack errPush, "modDatabase.NumberOfHistoryItems"
        '
        Dim recHistory As New ADODB.Recordset
        Dim strSQLHistory As String
        '
        strSQLHistory = "SELECT * FROM History"
        With recHistory
            .Open strSQLHistory, gcConnectionString, adOpenStatic, adLockReadOnly
            NumberOfHistoryItems = .RecordCount
            .Close
        End With
        '
        Set recHistory = Nothing
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    
