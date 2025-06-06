Attribute VB_Name = "modRegistry"
Option Explicit
'
' Hulpmodule voor het uitvoeren van registrybewerkingen
'
' Private variables
    Private m_lngRetVal As Long
'
' Constants required for values in the keys
    Private Const REG_NONE                          As Long = 0 ' No value type
    Private Const REG_SZ                            As Long = 1 ' nul terminated string
    Private Const REG_EXPAND_SZ                     As Long = 2 ' nul terminated string w/enviornment var
    Private Const REG_BINARY                        As Long = 3 ' Free form binary
    Private Const REG_DWORD                         As Long = 4 ' 32-bit number
    Private Const REG_DWORD_LITTLE_ENDIAN           As Long = 4 ' 32-bit number (same as REG_DWORD)
    Private Const REG_DWORD_BIG_ENDIAN              As Long = 5 ' 32-bit number
    Private Const REG_LINK                          As Long = 6 ' Symbolic Link (unicode)
    Private Const REG_MULTI_SZ                      As Long = 7 ' Multiple Unicode strings
    Private Const REG_RESOURCE_LIST                 As Long = 8 ' Resource list in the resource map
    Private Const REG_FULL_RESOURCE_DESCRIPTOR      As Long = 9 ' Resource list in the hardware description
    Private Const REG_RESOURCE_REQUIREMENTS_LIST    As Long = 10
'
' Registry Specific Access Rights
    Private Const KEY_QUERY_VALUE                   As Long = &H1
    Private Const KEY_SET_VALUE                     As Long = &H2
    Private Const KEY_CREATE_SUB_KEY                As Long = &H4
    Private Const KEY_ENUMERATE_SUB_KEYS            As Long = &H8
    Private Const KEY_NOTIFY                        As Long = &H10
    Private Const KEY_CREATE_LINK                   As Long = &H20
    Private Const KEY_ALL_ACCESS                    As Long = &H3F
'
' Constants required for key locations in the registry
    Public Const HKEY_CLASSES_ROOT                  As Long = &H80000000
    Public Const HKEY_CURRENT_USER                  As Long = &H80000001
    Public Const HKEY_LOCAL_MACHINE                 As Long = &H80000002
    Public Const HKEY_USERS                         As Long = &H80000003
    Public Const HKEY_PERFORMANCE_DATA              As Long = &H80000004
    Public Const HKEY_CURRENT_CONFIG                As Long = &H80000005
    Public Const HKEY_DYN_DATA                      As Long = &H80000006
'
' Constants required for return values (Error code checking)
    Private Const ERROR_SUCCESS                     As Long = 0
    Private Const ERROR_ACCESS_DENIED               As Long = 5
    Private Const ERROR_NO_MORE_ITEMS               As Long = 259
'
' Open/Create constants
    Private Const REG_OPTION_NON_VOLATILE           As Long = 0
    Private Const REG_OPTION_VOLATILE               As Long = &H1
'
' Declarations required to access the Windows registry
    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long
    Private Declare Function regCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
    Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String) As Long
    Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
    Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
'
' BESTANDSTYPE BEWERKEN
'

Public Sub SetFileAssociation(strExtension As String, strAppPath As String, strIdentifier As String, strDescription As String, strIcon As String, strCommand As String)
    Dim lngKeyHandle    As Long     ' Handle Of Open Key
    Dim lngBufferSize   As Long     ' Size Of String Data To Set
    Dim strOpenCommand  As String   '
    '
    ' Create The New Registry Key, the file extension
    m_lngRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, strExtension, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngKeyHandle, m_lngRetVal)
    '
    ' Check For An Error.
    If m_lngRetVal <> ERROR_SUCCESS Then MsgBox "Error: Creating Key!"
    '
    ' Get Size Of identifier String
    lngBufferSize = Len(strIdentifier)
    ' Set "(Default)" String Value to identifier
    m_lngRetVal = RegSetValueExString(lngKeyHandle, "", 0&, REG_SZ, strIdentifier, lngBufferSize)
    ' Check For An Error.
    If m_lngRetVal <> ERROR_SUCCESS Then
        MsgBox "Error Setting String Value!"
        RegCloseKey (lngKeyHandle)
        Exit Sub
    End If
    '
    ' Close The Registry Key.
    RegCloseKey (lngKeyHandle)
    '
    ' Create The New Registry Key, the file extension identifier
    m_lngRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, strIdentifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngKeyHandle, m_lngRetVal)
    '
    ' Check For An Error.
    If m_lngRetVal <> ERROR_SUCCESS Then MsgBox "Error: Creating Key!"
    '
    ' Get Size Of file type description String
    lngBufferSize = Len(strDescription)
    ' Set (Default) String Value to description of the file type
    m_lngRetVal = RegSetValueExString(lngKeyHandle, "", 0&, REG_SZ, strDescription, lngBufferSize)
    ' Check For An Error.
    If m_lngRetVal <> ERROR_SUCCESS Then
        MsgBox "Error Setting String Value!"
        RegCloseKey (lngKeyHandle)
        Exit Sub
    End If
    '
    ' Close The Registry Key.
    RegCloseKey (lngKeyHandle)
    ' Create The New Registry Key, the default icon key within the identifier key
    m_lngRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, (strIdentifier + "\DefaultIcon"), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngKeyHandle, m_lngRetVal)
    '
    ' Check For An Error.
    If m_lngRetVal <> ERROR_SUCCESS Then MsgBox "Error: Creating Key!"
    '
    ' Get Size Of String
    lngBufferSize = Len(strIcon)
    '
    ' Set (Default) String Value to the full path name of the icon that will be associated with this file type
    m_lngRetVal = RegSetValueExString(lngKeyHandle, "", 0&, REG_SZ, strIcon, lngBufferSize)
    ' Check For An Error.
    If m_lngRetVal <> ERROR_SUCCESS Then
        MsgBox "Error Setting String Value!"
        RegCloseKey (lngKeyHandle)
        Exit Sub
    End If
    '
    ' Close The Registry Key.
    RegCloseKey (lngKeyHandle)
    '
    ' Create The New Registry Key, the "shell" key within the identifier key
    strIdentifier = strIdentifier + "\shell"
    m_lngRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, strIdentifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngKeyHandle, m_lngRetVal)
    '
    ' Check For An Error.
    If m_lngRetVal <> ERROR_SUCCESS Then MsgBox "Error: Creating Key!"
    '
    ' Close The Registry Key.
    RegCloseKey (lngKeyHandle)
    '
    ' Create The New Registry Key, the "open" command key within the shell key
    strIdentifier = strIdentifier + "\open"
    m_lngRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, strIdentifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngKeyHandle, m_lngRetVal)
    '
    ' Check For An Error.
    If m_lngRetVal <> ERROR_SUCCESS Then MsgBox "Error: Creating Key!"
    '
    ' Close The Registry Key.
    RegCloseKey (lngKeyHandle)
    '
    ' Create The New Registry Key, the "command"  key within the "open" command key
    strIdentifier = strIdentifier + "\command"
    m_lngRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, strIdentifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngKeyHandle, m_lngRetVal)
    '
    ' Check For An Error.
    If m_lngRetVal <> ERROR_SUCCESS Then MsgBox "Error: Creating Key!"
    '
    ' Set (Default) String Value of the "command" key to the command line to be used to open the file
    strOpenCommand = Chr(34) & strAppPath & Chr(34) & " " & strCommand & " " & Chr(34) & "%1" & Chr(34)
    lngBufferSize = Len(strOpenCommand)
    m_lngRetVal = RegSetValueExString(lngKeyHandle, "", 0&, REG_SZ, strOpenCommand, lngBufferSize)
    ' Check For An Error.
    If m_lngRetVal <> ERROR_SUCCESS Then
        MsgBox "Error Setting String Value!"
        RegCloseKey (lngKeyHandle)
        Exit Sub
    End If
    '
    ' Close The Registry Key.
    RegCloseKey (lngKeyHandle)
End Sub
'
' TOEGANGS EN BEWERKINGSGEDEELTE
'
    ' Written by Kenneth Ives                     kenaso@home.com
    '
    ' Perform the four basic functions on the Windows registry.
    '           Add
    '           Change
    '           Delete
    '           Query
    '
    Public Function regDelete_Sub_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, ByVal strRegSubKey As String)
        '
        ' Description:
        '           Function for removing a sub key.
        '
        ' Parameters:
        '           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
        '           strRegKeyPath - is name of the key path you wish to traverse.
        '           strRegSubKey - is the name of the key which will be removed.
        '
        ' Syntax:
        '           regDelete_Sub_Key HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products", "StringTestData"
        '           Removes the sub key "StringTestData"
        '
        Dim lngKeyHandle As Long
        '
        If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
            ' Get the key handle
            m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
            '
            ' Delete the sub key.  If it does not exist, then ignore it.
            m_lngRetVal = RegDeleteKey(lngKeyHandle, strRegSubKey)  'RegDeleteValue(lngKeyHandle, strRegSubKey)
            '
            m_lngRetVal = RegCloseKey(lngKeyHandle)
        End If
    End Function
    '
    Public Function regDoes_Key_Exist(ByVal lngRootKey As Long, ByVal strRegKeyPath As String) As Boolean
        ' Description:
        '       Function to see if a key does exist
        '
        ' Parameters:
        '       lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
        '       strRegKeyPath - is name of the key path you want to test
        '
        ' Syntax:
        '       strKeyQuery = regQuery_A_Key(HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products")
        '       Returns the value of TRUE or FALSE
        '
        Dim lngKeyHandle As Long: lngKeyHandle = 0
        '
        ' Query the key path
        m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
        '
        ' If no key handle was found then there is no key.  Leave here.
        If lngKeyHandle = 0 Then
            regDoes_Key_Exist = False
        Else
            regDoes_Key_Exist = True
        End If
        '
        m_lngRetVal = RegCloseKey(lngKeyHandle)
    End Function
    '
    Public Function regQuery_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, ByVal strRegSubKey As String) As Variant
        '
        ' Description:
        '       Function for querying a sub key value.
        '
        ' Parameters:
        '       lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
        '       strRegKeyPath - is name of the key path you wish to traverse.
        '       strRegSubKey - is the name of the key which will be queryed.
        '
        ' Syntax:
        '       strKeyQuery = regQuery_A_Key(HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products", "StringTestData")
        '
        Dim intPosition As Integer
        Dim lngKeyHandle As Long: lngKeyHandle = 0
        Dim lngDataType As Long
        Dim lngBufferSize As Long: lngBufferSize = 0
        Dim lngBuffer As Long
        Dim strBuffer As String
        '
        ' Query the key path
        m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
        '
        ' If no key handle was found then there is no key.  Leave here.
        If lngKeyHandle = 0 Then
            regQuery_A_Key = ""
            m_lngRetVal = RegCloseKey(lngKeyHandle)
            Exit Function
        End If
        '
        ' Query the registry and determine the data type.
        m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal 0&, lngBufferSize)
        '
        ' If no key handle was found then there is no key.  Leave here.
        If lngKeyHandle = 0 Then
            regQuery_A_Key = ""
            m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
            Exit Function
        End If
        '
        ' Make the API call to query the registry based on the type of data.
        Select Case lngDataType
            Case REG_SZ ' String
                ' Preload the receiving buffer area
                strBuffer = Space(lngBufferSize)
                '
                m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, ByVal strBuffer, lngBufferSize)
                '
                If m_lngRetVal <> ERROR_SUCCESS Then
                    regQuery_A_Key = ""
                Else
                    ' Strip out the string data
                    intPosition = InStr(1, strBuffer, Chr(0))  ' look for the first null char
                    If intPosition > 0 Then
                        ' if we found one, then save everything up to that point
                        regQuery_A_Key = Left(strBuffer, intPosition - 1)
                    Else
                        ' did not find one.  Save everything.
                        regQuery_A_Key = strBuffer
                    End If
                End If
            Case REG_DWORD 'Integer
                m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, lngBuffer, 4&)  ' 4& = 4-byte word (long integer)
                '
                ' If NOT a successful call then leave
                If m_lngRetVal <> ERROR_SUCCESS Then
                    regQuery_A_Key = ""
                Else
                    ' Save the captured data
                    regQuery_A_Key = lngBuffer
                End If
            Case Else: 'Onbekend
                regQuery_A_Key = ""
            End Select
            '
            m_lngRetVal = RegCloseKey(lngKeyHandle)
    End Function
    '
    Public Sub regCreate_Key_Value(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, ByVal strRegSubKey As String, varRegData As Variant)
        ' Description:
        '       Function for saving string data.
        '
        ' Parameters:
        '       lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
        '       strRegKeyPath - is name of the key path you wish to traverse.
        '       strRegSubKey - is the name of the key which will be updated.
        '       varRegData - Update data.
        '
        ' Syntax:
        '       regCreate_Key_Value HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products", "StringTestData", "22 Jun 1999"
        '       Saves the key value of "22 Jun 1999" to sub key "StringTestData"
        '
        Dim lngKeyHandle As Long
        Dim lngDataType As Long
        Dim lngKeyValue As Long
        Dim strKeyValue As String
        '
        ' Determine the type of data to be updated
        If IsNumeric(varRegData) Then
            lngDataType = REG_DWORD
        Else
            lngDataType = REG_SZ
        End If
        '
        ' Query the key path
        m_lngRetVal = regCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
        '
        ' Update the sub key based on the data type
        Select Case lngDataType
            Case REG_SZ:       ' String data
                strKeyValue = Trim(varRegData) & Chr(0)     ' null terminated
                m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal strKeyValue, Len(strKeyValue))
            Case REG_DWORD:    ' numeric data
                lngKeyValue = CLng(varRegData)
                m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, lngKeyValue, 4&)  ' 4& = 4-byte word (long integer)
        End Select
        '
        ' Always close the handle in the registry.  We do not want to corrupt these files.
        m_lngRetVal = RegCloseKey(lngKeyHandle)
    End Sub
    '
    Public Function regCreate_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String)
        ' Description:
        '       This function will create a new key
        '
        ' Parameters:
        '       lngRootKey  - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
        '       strRegKeyPath - is name of the key you wish to create. to make sub keys, continue to make this call with each new level.  MS says you can do this in one call; however, the best laid plans of mice and men ...
        '
        ' Syntax:
        '       regCreate_A_Key HKEY_CURRENT_USER, "Software\AAA-Registry Test"
        '       regCreate_A_Key HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products"
        '
        Dim lngKeyHandle As Long
        '
        ' Create the key.  If it already exist, ignore it.
        m_lngRetVal = regCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
        '
        m_lngRetVal = RegCloseKey(lngKeyHandle)
    End Function
    '
    Public Function regDelete_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, ByVal strRegKeyName As String) As Boolean
        ' Description:
        '       Function for removing a complete key.
        '
        ' Parameters:
        '       lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
        '       strRegKeyPath - is name of the key path you wish to traverse.
        '       strRegKeyValue - is the name of the key which will be removed.
        '
        ' Syntax:
        '       regDelete_A_Key HKEY_CURRENT_USER, "Software", "AAA-Registry Test"
        '
        ' Removes the key "AAA-Registry Test" and all of its sub keys.
        ' Returns a True or False on completion.
        '
        Dim lngKeyHandle As Long
        '
        regDelete_A_Key = False
        '
        If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
            ' Get the key handle
            m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
            
            ' Delete the key
            m_lngRetVal = RegDeleteKey(lngKeyHandle, strRegKeyName)
            '
            ' If the value returned is equal zero then we have succeeded
            If m_lngRetVal = 0 Then regDelete_A_Key = True
            
            ' Always close the handle in the registry.  We do not want to corrupt the registry.
            m_lngRetVal = RegCloseKey(lngKeyHandle)
        End If
    End Function
'
' ODBC KOPPELING
'
    Public Sub CheckDSN(strDSN As String)
        Dim szDriverName As String
        szDriverName = String(255, Chr(32))
        If Not checkAccessDriver(szDriverName) Then
            MsgBox "No Access Drivers on this PC", vbOK + vbCritical, AppName
            End
        End If
        'If Not (checkWantedDSN(strDSN)) Then
            If szDriverName = vbNullString Then
                MsgBox "Can't find access ODBC driver.", vbOK + vbCritical, AppName
            Else
                If Not createAccessDSN(szDriverName, strDSN) Then
                    MsgBox "Aanmaken ODBC mislukt.", vbOK + vbCritical, AppName
                End If
            End If
        'End If
    End Sub
    
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
    Public Function createAccessDSN(szDriverName As String, _
                                szWantedDSN As String) As Boolean
        Dim HKey As Long
        Dim szKeyPath As String
        Dim szKeyName As String
        Dim szKeyValue As String
        Dim lKeyValue As Long
        Dim lRes As Long
        Dim lSize As Long
        Dim szEmpty As String
        szEmpty = Chr(0)
        lSize = 4
        lRes = regCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN, HKey)
        If lRes <> ERROR_SUCCESS Then
            createAccessDSN = False
            Exit Function
        End If
        lRes = RegSetValueExString(HKey, "UID", 0&, REG_SZ, szEmpty, Len(szEmpty))
        szKeyValue = App.Path & "\" & gcMdbName
        lRes = RegSetValueExString(HKey, "DBQ", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
        szKeyValue = "Verbodrome Database"
        lRes = RegSetValueExString(HKey, "Description", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
        szKeyValue = szDriverName
        lRes = RegSetValueExString(HKey, "Driver", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
        szKeyValue = "MS Access;"
        lRes = RegSetValueExString(HKey, "FIL", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
        lKeyValue = 25
        lRes = RegSetValueExLong(HKey, "DriverId", 0&, REG_DWORD, lKeyValue, 4)
        lKeyValue = 0
        lRes = RegSetValueExLong(HKey, "SafeTransactions", 0&, REG_DWORD, lKeyValue, 4)
        lRes = RegCloseKey(HKey)
        szKeyPath = "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN & "\Engines\Jet"
        lRes = regCreateKey(HKEY_LOCAL_MACHINE, szKeyPath, HKey)
        If lRes <> ERROR_SUCCESS Then
            createAccessDSN = False
            Exit Function
        End If
        lRes = RegSetValueExString(HKey, "ImplicitCommitSync", 0&, REG_SZ, szEmpty, Len(szEmpty))
        szKeyValue = "Yes"
        lRes = RegSetValueExString(HKey, "UserCommitSync", 0&, REG_SZ, szKeyValue, Len(szKeyValue))
        lKeyValue = 2048
        lRes = RegSetValueExLong(HKey, "MaxBufferSize", 0&, REG_DWORD, lKeyValue, 4)
        lKeyValue = 5
        lRes = RegSetValueExLong(HKey, "PageTimeout", 0&, REG_DWORD, lKeyValue, 4)
        lKeyValue = 3
        lRes = RegSetValueExLong(HKey, "Threads", 0&, REG_DWORD, lKeyValue, 4)
        lRes = RegCloseKey(HKey)
        lRes = regCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", HKey)
        If lRes <> ERROR_SUCCESS Then
            createAccessDSN = False
            Exit Function
        End If
        szKeyValue = "Microsoft Access Driver (*.mdb)"
        lRes = RegSetValueExString(HKey, szWantedDSN, 0&, REG_SZ, szKeyValue, Len(szKeyValue))
        lRes = RegCloseKey(HKey)
        createAccessDSN = True
    End Function
    '
    Public Function isSZKeyExist(szKeyPath As String, szKeyName As String, _
                             ByRef szKeyValue As String) As Boolean
        Dim bRes As Boolean
        Dim lRes As Long
        Dim HKey As Long
        lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, szKeyPath, 0&, KEY_QUERY_VALUE, HKey)
        If lRes <> ERROR_SUCCESS Then
            isSZKeyExist = False
            Exit Function
        End If
        lRes = RegQueryValueEx(HKey, szKeyName, 0&, REG_SZ, ByVal szKeyValue, Len(szKeyValue))
        RegCloseKey (HKey)
        If lRes <> ERROR_SUCCESS Then
            isSZKeyExist = False
            Exit Function
        End If
        isSZKeyExist = True
    End Function
