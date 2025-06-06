Attribute VB_Name = "modTrial"
Option Explicit
'
'Gebruikte registry Keys:
' HKEY_CURRENT_USER\Software\VB and VBA Program Settings\MSBMBM60.DLL\Control   'Installatiedatum
'
Private Declare Function GetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSytemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'
' SUBS
'
    Public Sub ShowRemainingDays()
        On Error GoTo ErrorOccured
        errProcStack errPush, "modTrial.ShowRemainingDays"
        '
        If RegisteredVersion = False Then
            Dim intDaysRemaining As Integer
            Dim strRemainingDays As String
            '
            intDaysRemaining = DaysRemaining
            '
            If intDaysRemaining > 30 Then
                'Iemand  heeft valsgespeeld!!!
                Call WriteFirstUseDate
                intDaysRemaining = 30
            End If
            '
            If intDaysRemaining >= 0 Then
                strRemainingDays = GetTxt("Registration_DaysLeft")
                strRemainingDays = Replace(strRemainingDays, "$daysleft$", intDaysRemaining)
                MsgBox strRemainingDays, vbInformation, AppName
            Else
                strRemainingDays = GetTxt("Registration_Expired")
                MsgBox strRemainingDays, vbCritical, AppName
                'Registratiecode vragen, indien fout afsluiten
                If Register = False Then End
            End If
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Public Sub WriteFirstUseDate()
        Call SaveSetting("Microsoft", "MSBMBM60.DLL", "Control", EncryptDate(Now))
    End Sub
'
' FUNCTIONS
'
    Public Function Register() As Boolean
        On Error GoTo ErrorOccured
        errProcStack errPush, "modTrial.Register"
        '
        Register = RegisteredVersion
        '
        If Register = False Then
            Dim strRegistrationPrompt As String
            '
            strRegistrationPrompt = GetTxt("RegistrationPrompt")
            strRegistrationPrompt = Replace(strRegistrationPrompt, "$serialnumber$", CStr(GeneratePCCode))
            '
            Dim strRegistrationKey As String: strRegistrationKey = GetDBSetting("RegistrationKey")
            '
            strRegistrationKey = InputBox(strRegistrationPrompt, AppName)
            '
            If UCase(GenerateRegKey) = UCase(strRegistrationKey) Then
                'Juiste key
                MsgBox GetTxt("RegistrationKey_Correct"), vbInformation, AppName
                LetDBSetting "RegistrationKey", strRegistrationKey
                Register = True
            Else
                If strRegistrationKey = "/resetme" Then
                    'Noodoplossing
                    Call WriteFirstUseDate
                Else
                    'Foute key
                    MsgBox GetTxt("RegistrationKey_Wrong"), vbExclamation, AppName
                End If
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
    Public Function DaysRemaining() As Integer
        On Error Resume Next
        '
        DaysRemaining = 30 - DateDiff("d", DecryptDate(GetFirstUseDate(EncryptDate(Now))), Now)
    End Function
    '
    Private Function GeneratePCCode() As Long
        On Error Resume Next
        '
        'Code
        Call GetVolumeInformation(Left(App.Path, 3), vbNullString, 0, GeneratePCCode, ByVal 0&, ByVal 0&, vbNullString, 0)
        GeneratePCCode = Abs(GeneratePCCode)
    End Function
    '
    Public Function GenerateRegKey() As String
        On Error Resume Next
        '
        'Tegencode
        Dim i As Double: i = 1
        Dim strDECMachineCode As Long
        '
        strDECMachineCode = CLng(Left(CStr(Abs(Val("&H" & GeneratePCCode))), 9))
        Do
            i = i * 10
        Loop Until i > strDECMachineCode
        '
        GenerateRegKey = CStr(Hex(i - strDECMachineCode))
    End Function
    '
    Public Function RegisteredVersion() As Boolean
        On Error Resume Next
        '
        RegisteredVersion = Not (blnTrialVersion)
        If RegisteredVersion = False Then RegisteredVersion = (GetDBSetting("RegistrationKey") = GenerateRegKey)
    End Function
    '
    Public Function EncryptDate(dtmDecryptedDate As Date) As String
        On Error Resume Next
        '
        'EncryptDate = Hex(Format(Now, "ddmmyyyy")) & "G" & Hex(Format(Now, "hhmmss"))
        EncryptDate = Hex("1" & GetDateSerial) & "G" & Hex(Format(Now, "hhmmss"))
    End Function
    '
    Public Function DecryptDate(strEncryptedDate As String) As Date
        On Error Resume Next
        '
        strEncryptedDate = Val("&H" & Left(strEncryptedDate, InStr(strEncryptedDate, "G") - 1)) & Val("&H" & Mid(strEncryptedDate, InStr(strEncryptedDate, "G") + 1))
        '
        Dim strDate(0 To 2) As String
        strDate(0) = Mid(strEncryptedDate, 2, 2)
        strDate(1) = Mid(strEncryptedDate, 4, 2)
        strDate(2) = Mid(strEncryptedDate, 6, 4)
        'strDate(3) = Mid(strEncryptedDate, 9, 2)
        'strDate(4) = Mid(strEncryptedDate, 11, 2)
        'strDate(5) = Mid(strEncryptedDate, 13, 2)
        '
        DecryptDate = DateSerial(strDate(2), strDate(1), strDate(0)) '+ TimeSerial(strDate(3), strDate(4), strDate(5))
    End Function
    '
    Public Function GetFirstUseDate(Optional strDefault As String) As String
        On Error Resume Next
        '
        GetFirstUseDate = GetSetting("Microsoft", "MSBMBM60.DLL", "Control", IIf(strDefault = "", EncryptDate(Now), strDefault))
    End Function
    '
    Private Function GetDateSerial() As String
        On Error Resume Next
        '
        GetDateSerial = IIf(Day(Now) < 10, "0" & Day(Now), Day(Now)) & _
                        IIf(Month(Now) < 10, "0" & Month(Now), Month(Now)) & _
                        Year(Now)
    End Function
    '
    Public Function GenerateCustomRegKey(strCustomRegKey As String) As String
        On Error Resume Next
        '
        'Tegencode
        Dim i As Double: i = 1
        Dim strDECMachineCode As Long
        '
        strDECMachineCode = CLng(Left(CStr(Abs(Val("&H" & strCustomRegKey))), 9))
        Do
            i = i * 10
        Loop Until i > strDECMachineCode
        '
        GenerateCustomRegKey = CStr(Hex(i - strDECMachineCode))
    End Function
