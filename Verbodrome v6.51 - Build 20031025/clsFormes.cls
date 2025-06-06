VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
Private recForm             As New ADODB.Recordset
'
Private strInfinitives()    As String
'
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
'
' SUBS
'
    ' Constructor
    Private Sub Class_Initialize()
        ReDim strInfinitives(1 To 1)
    End Sub
    '
    ' AddInfinitive
    Public Sub AddInfinitive(strInfinitive As String, intIndexNumber As Integer)
        If intIndexNumber > UBound(strInfinitives) Then ReDim Preserve strInfinitives(1 To intIndexNumber)
        strInfinitives(intIndexNumber) = strInfinitive
    End Sub
    '
    ' ResetPosedDB
    Public Sub ResetPosedDB()
        If LCase(Dir(App.Path & "\LocalUser\LocalSettings.ini")) = "localsettings.ini" Then
            Kill App.Path & "\LocalUser\LocalSettings.ini"
        End If
    End Sub
    '
    ' Destructor
    Private Sub Class_Terminate()
        Set recForm = Nothing
    End Sub
'
' FUNCTIONS
'
    ' GetEnumerator
    Public Function GetEnumerator() As String()
        GetEnumerator = strInfinitives
    End Function
    '
    ' GetInfinitive
    Public Function GetInfinitive(intIndexNumber As Integer) As String
        GetInfinitive = strInfinitives(intIndexNumber)
    End Function
    '
    ' GetIndexNumber
    Public Function GetIndexNumber(strInfinitive As String) As Integer
        Dim i As Integer
        '
        For i = LBound(strInfinitives) To UBound(strInfinitives)
            If strInfinitives(i) = strInfinitive Then
                GetIndexNumber = i
                Exit For
            End If
        Next i
    End Function
    '
    ' ReadDatabase
    Private Function ReadDatabase(intInfinitif As Integer, bytTemps As Byte, bytPersonne As Byte) As String
        Dim strInfinitive As String:    strInfinitive = GetInfinitive(intInfinitif)
        Dim strTemps As String:         strTemps = Convert(AbsNumber, Abreviation, bytTemps)
        Dim strPersonnes() As String
        '
        strSQL = "SELECT " & strTemps & " FROM Verbes WHERE InfPr='" & strInfinitive & "'"
        '
        With recForm
            .Open strSQL, gcnConnection, adOpenStatic, adLockReadOnly
            '
            strPersonnes = Split(.Fields(strTemps), ";")
            ReadDatabase = strPersonnes(bytPersonne - 1)
            '
            .Close
        End With
    End Function
    '
    ' ReadINI
    Private Function ReadINI(strSection As String, strKey As String, strDatabaseLocation As String, Optional strDefault As String) As String
        On Error Resume Next
        '
        Dim sData As String: sData = Space(8192)
        Dim lDataLen As Long: lDataLen = GetPrivateProfileString(strSection, strKey, strDefault, sData, Len(sData), strDatabaseLocation)
        '
        ReadINI = Left(sData, lDataLen)
        '
        If blnPM_DebugLog Then Call LogDebug("ReadINI: Section: " & strSection & ", Key: " & strKey & ", Default: " & strDefault & ", Result: " & ReadINI)
    End Function
    '
    ' WriteINI
    Private Function WriteINI(strSection As String, strKey As String, strSetting As String, strDatabaseLocation As String)
        On Error Resume Next
        '
        Call WritePrivateProfileString(strSection, strKey, strSetting, strDatabaseLocation)
        '
        If blnPM_DebugLog Then Call LogDebug("WriteINI: Section: " & strSection & ", Key: " & strKey & ", Setting: " & strSetting)
    End Function
'
' PROPERTIES
'
    '
    ' Form
    Public Property Get Form(intInfinitif As Integer, bytTemps As Byte, bytPersonne As Byte) As String
        Dim strAnswer As String: strAnswer = ReadDatabase(intInfinitif, bytTemps, bytPersonne)
        '
        Form = Replace(strAnswer, "*", "")
    End Property
    '
    ' Allowed
    Public Property Get Allowed(intInfinitif As Integer, bytTemps As Byte, bytPersonne As Byte) As Boolean
        Dim strForm As String
        '
        strForm = Form(intInfinitif, bytTemps, bytPersonne)
        '
        If Not (blnIsVerbAllowed(intInfinitif)) Or strForm = "" Or strForm = "-" Then
            '1. Werkwoord is niet toegelaten
            '2. Vorm bestaat niet ( leeg of - )
            Allowed = False
        ElseIf bytTemps >= gcIndPr And bytTemps <= gcSubjPQP Then
            'Tijd met 6 vormen
            Allowed = (blnIsTempsAllowed(bytTemps) And blnIsPersonneAllowed(bytPersonne))
        ElseIf bytTemps = gcImpPr Or bytTemps = gcImpPa Then
            'Tijd met 3 vormen: Imperatieven
            Allowed = (blnIsTempsAllowed(bytTemps) And blnIsPersonneAllowed(Choose(bytPersonne, 2, 4, 5)))
        Else
            'Rest
            Allowed = blnIsTempsAllowed(bytTemps)
        End If
    End Property
    '
    ' Posed
    Public Property Get Posed(intInfinitif As Integer, bytTemps As Byte, bytPersonne As Byte) As Boolean
        Dim strKey As String: strKey = GetInfinitive(intInfinitif) & "," & CStr(bytTemps) & "," & CStr(bytPersonne)
        '
        Posed = ReadINI("PosedQuestions", strKey, App.Path & "\LocalUser\LocalSettings.ini", False)
    End Property
    '
    Public Property Let Posed(intInfinitif As Integer, bytTemps As Byte, bytPersonne As Byte, blnPosed As Boolean)
        Dim strKey As String: strKey = GetInfinitive(intInfinitif) & "," & CStr(bytTemps) & "," & CStr(bytPersonne)
        '
        Call WriteINI("PosedQuestions", strKey, CStr(blnPosed), App.Path & "\LocalUser\LocalSettings.ini")
    End Property
    '
    ' Count
    Public Property Get Count() As Integer
        Count = UBound(strInfinitives)
    End Property
