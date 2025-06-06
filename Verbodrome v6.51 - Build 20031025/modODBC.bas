Attribute VB_Name = "modODBC"
Option Explicit

Private Const KEY_QUERY_VALUE = &H1
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_DWORD = 4
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long ' Note that If you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function regCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
'
Public Function isSZKeyExist(szKeyPath As String, szKeyName As String, _
                             ByRef szKeyValue As String) As Boolean
    Dim bRes As Boolean
    Dim lRes As Long
    Dim hKey As Long
    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, szKeyPath, 0&, KEY_QUERY_VALUE, hKey)
    If lRes <> ERROR_SUCCESS Then
        isSZKeyExist = False
        Exit Function
    End If
    lRes = RegQueryValueEx(hKey, szKeyName, 0&, REG_SZ, ByVal szKeyValue, Len(szKeyValue))
    RegCloseKey (hKey)
    If lRes <> ERROR_SUCCESS Then
        isSZKeyExist = False
        Exit Function
    End If
    isSZKeyExist = True
End Function
'
Public Function checkAccessDriver(ByRef szDriverName As String) As Boolean
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean
    
    bRes = False
    szKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\Microsoft Access Driver (*.mdb)"
    szKeyName = "Driver"
    szKeyValue = String(255, Chr(32))
    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        szDriverName = szKeyValue
        bRes = True
    Else
        bRes = False
    End If
    checkAccessDriver = bRes
End Function
'
Public Function checkWantedDSN(szWantedDSN As String) As Boolean
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean, checkWantedAccessDSN As Boolean
    '
    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
    szKeyName = szWantedDSN
    szKeyValue = String(255, Chr(32))
    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        bRes = True
    Else
        bRes = False
    End If
    checkWantedDSN = bRes
End Function
'
Public Function createAccessDSN(szDriverName As String, szWantedDSN As String, strDBPath As String) As Boolean
    Dim hKey As Long
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim lKeyValue As Long
    Dim lRes As Long
    Dim lSize As Long
    Dim szEmpty As String
    szEmpty = Chr(0)
    lSize = 4
    lRes = regCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN, hKey)
    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If
    lRes = RegSetValueExString(hKey, "UID", 0&, REG_SZ, szEmpty, Len(szEmpty))
    szKeyValue = strDBPath
    lRes = RegSetValueExString(hKey, "DBQ", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
    szKeyValue = szWantedDSN
    lRes = RegSetValueExString(hKey, "Description", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
    szKeyValue = szDriverName
    lRes = RegSetValueExString(hKey, "Driver", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
    szKeyValue = "MS Access;"
    lRes = RegSetValueExString(hKey, "FIL", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
    lKeyValue = 25
    lRes = RegSetValueExLong(hKey, "DriverId", 0&, REG_DWORD, lKeyValue, 4)
    lKeyValue = 0
    lRes = RegSetValueExLong(hKey, "SafeTransactions", 0&, REG_DWORD, lKeyValue, 4)
    lRes = RegCloseKey(hKey)
    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN & "\Engines\Jet"
    lRes = regCreateKey(HKEY_LOCAL_MACHINE, szKeyPath, hKey)
    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If
    lRes = RegSetValueExString(hKey, "ImplicitCommitSync", 0&, REG_SZ, szEmpty, Len(szEmpty))
    szKeyValue = "Yes"
    lRes = RegSetValueExString(hKey, "UserCommitSync", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
    lKeyValue = 2048
    lRes = RegSetValueExLong(hKey, "MaxBufferSize", 0&, REG_DWORD, lKeyValue, 4)
    lKeyValue = 5
    lRes = RegSetValueExLong(hKey, "PageTimeout", 0&, REG_DWORD, lKeyValue, 4)
    lKeyValue = 3
    lRes = RegSetValueExLong(hKey, "Threads", 0&, REG_DWORD, lKeyValue, 4)
    lRes = RegCloseKey(hKey)
    lRes = regCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKey)
    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If
    szKeyValue = "Microsoft Access Driver (*.mdb)"
    lRes = RegSetValueExString(hKey, szWantedDSN, 0&, REG_SZ, szKeyValue, Len(szKeyValue))
    lRes = RegCloseKey(hKey)
    createAccessDSN = True
End Function
'
Public Sub CheckDSN(strDSN As String, strDBPath As String)
    '
    ' Procedure die de DSN koppeling controleert, indien niet bestaande, aanmaakt
    '
    Dim szDriverName As String
    szDriverName = String(255, Chr(32))
    '
    If Not checkAccessDriver(szDriverName) Then
        MsgBox "No Access Drivers on this PC.", vbOK + vbCritical
        Exit Sub
    End If
    '
    'If Not (checkWantedDSN(strDSN)) Then
        If szDriverName = "" Then
            MsgBox "Can't find access ODBC driver.", vbOK + vbCritical
        Else
            If Not createAccessDSN(szDriverName, strDSN, strDBPath) Then
                MsgBox "Aanmaken ODBC mislukt.", vbOK + vbCritical
            End If
        End If
    'End If
End Sub


'Option Explicit
'
'Private Const KEY_QUERY_VALUE = &H1
'Private Const ERROR_SUCCESS = 0&
'Private Const REG_SZ = 1
'Private Const HKEY_LOCAL_MACHINE = &H80000002
'Private Const REG_DWORD = 4
'Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long ' Note that If you declare the lpData parameter as String, you must pass it By Value.
'Private Declare Function regCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
''
'Public Function isSZKeyExist(szKeyPath As String, szKeyName As String, _
'                             ByRef szKeyValue As String) As Boolean
'    Dim bRes As Boolean
'    Dim lRes As Long
'    Dim hKey As Long
'    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, szKeyPath, 0&, KEY_QUERY_VALUE, hKey)
'    If lRes <> ERROR_SUCCESS Then
'        isSZKeyExist = False
'        Exit Function
'    End If
'    lRes = RegQueryValueEx(hKey, szKeyName, 0&, REG_SZ, ByVal szKeyValue, Len(szKeyValue))
'    RegCloseKey (hKey)
'    If lRes <> ERROR_SUCCESS Then
'        isSZKeyExist = False
'        Exit Function
'    End If
'    isSZKeyExist = True
'End Function
''
'Public Function checkAccessDriver(ByRef szDriverName As String) As Boolean
'    Dim szKeyPath As String
'    Dim szKeyName As String
'    Dim szKeyValue As String
'    Dim bRes As Boolean
'
'    bRes = False
'    szKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\Microsoft Access Driver (*.mdb)"
'    szKeyName = "Driver"
'    szKeyValue = String(255, Chr(32))
'    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
'        szDriverName = szKeyValue
'        bRes = True
'    Else
'        bRes = False
'    End If
'    checkAccessDriver = bRes
'End Function
''
'Public Function checkWantedDSN(szWantedDSN As String) As Boolean
'    Dim szKeyPath As String
'    Dim szKeyName As String
'    Dim szKeyValue As String
'    Dim bRes As Boolean, checkWantedAccessDSN As Boolean
'    '
'    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
'    szKeyName = szWantedDSN
'    szKeyValue = String(255, Chr(32))
'    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
'        bRes = True
'    Else
'        bRes = False
'    End If
'    checkWantedDSN = bRes
'End Function
''
'Public Function createAccessDSN(szDriverName As String, szWantedDSN As String, strDBPath As String) As Boolean
'    Dim hKey As Long
'    Dim szKeyPath As String
'    Dim szKeyName As String
'    Dim szKeyValue As String
'    Dim lKeyValue As Long
'    Dim lRes As Long
'    Dim lSize As Long
'    Dim szEmpty As String
'    szEmpty = Chr(0)
'    lSize = 4
'    lRes = regCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN, hKey)
'    If lRes <> ERROR_SUCCESS Then
'        createAccessDSN = False
'        Exit Function
'    End If
'    lRes = RegSetValueExString(hKey, "UID", 0&, REG_SZ, szEmpty, Len(szEmpty))
'    szKeyValue = strDBPath
'    lRes = RegSetValueExString(hKey, "DBQ", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
'    szKeyValue = szWantedDSN
'    lRes = RegSetValueExString(hKey, "Description", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
'    szKeyValue = szDriverName
'    lRes = RegSetValueExString(hKey, "Driver", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
'    szKeyValue = "MS Access;"
'    lRes = RegSetValueExString(hKey, "FIL", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
'    lKeyValue = 25
'    lRes = RegSetValueExLong(hKey, "DriverId", 0&, REG_DWORD, lKeyValue, 4)
'    lKeyValue = 0
'    lRes = RegSetValueExLong(hKey, "SafeTransactions", 0&, REG_DWORD, lKeyValue, 4)
'    lRes = RegCloseKey(hKey)
'    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN & "\Engines\Jet"
'    lRes = regCreateKey(HKEY_LOCAL_MACHINE, szKeyPath, hKey)
'    If lRes <> ERROR_SUCCESS Then
'        createAccessDSN = False
'        Exit Function
'    End If
'    lRes = RegSetValueExString(hKey, "ImplicitCommitSync", 0&, REG_SZ, szEmpty, Len(szEmpty))
'    szKeyValue = "Yes"
'    lRes = RegSetValueExString(hKey, "UserCommitSync", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
'    lKeyValue = 2048
'    lRes = RegSetValueExLong(hKey, "MaxBufferSize", 0&, REG_DWORD, lKeyValue, 4)
'    lKeyValue = 5
'    lRes = RegSetValueExLong(hKey, "PageTimeout", 0&, REG_DWORD, lKeyValue, 4)
'    lKeyValue = 3
'    lRes = RegSetValueExLong(hKey, "Threads", 0&, REG_DWORD, lKeyValue, 4)
'    lRes = RegCloseKey(hKey)
'    lRes = regCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKey)
'    If lRes <> ERROR_SUCCESS Then
'        createAccessDSN = False
'        Exit Function
'    End If
'    szKeyValue = "Microsoft Access Driver (*.mdb)"
'    lRes = RegSetValueExString(hKey, szWantedDSN, 0&, REG_SZ, szKeyValue, Len(szKeyValue))
'    lRes = RegCloseKey(hKey)
'    createAccessDSN = True
'End Function
''
'Public Sub CheckDSN(strDSN As String, strDBPath As String)
'    '
'    ' Procedure die de DSN koppeling controleert, indien niet bestaande, aanmaakt
'    '
'    Dim szDriverName As String
'    szDriverName = String(255, Chr(32))
'    '
'    If Not checkAccessDriver(szDriverName) Then
'        MsgBox "No Access Drivers on this PC.", vbOK + vbCritical
'        Exit Sub
'    End If
'    '
'    'If Not (checkWantedDSN(strDSN)) Then
'        If szDriverName = "" Then
'            MsgBox "Can't find access ODBC driver.", vbOK + vbCritical
'        Else
'            If Not createAccessDSN(szDriverName, strDSN, strDBPath) Then
'                MsgBox "Aanmaken ODBC mislukt.", vbOK + vbCritical
'            End If
'        End If
'    'End If
'End Sub
