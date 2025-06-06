Attribute VB_Name = "modGlobals"
Option Explicit
'
' VARIABELEN
'
    '
    ' Databasevariabelen
    '
    Public gcnConnection                As New ADODB.Connection
    Public recVerbodrome                As New ADODB.Recordset
    Public strSQL                       As String
    '
    '
    ' Toelatingsvariabelen
    '
    Public intAllowedVerbes()           As Integer
    Public bytAllowedTemps()            As Byte
    Public bytAllowedPersonnes()        As Byte
    '
    Public blnIsTempsAllowed(1 To 22)   As Boolean
    Public blnIsPersonneAllowed(1 To 6) As Boolean
    Public blnIsVerbAllowed()           As Boolean
    '
    ' Taalvariabele
    '
    Public bytLanguage As Byte
    '
    ' Statistiek variabelen
    '
    Public dtmProgramStart              As Date
    Public lngSessionID                 As Long
    '
    ' Runtime variabelen
    '
    Public mySession                    As New clsSession
    Public myFormes                     As New clsFormes
    '
    Public blnHistoryEnabled            As Boolean
    Public blnPM_DBEdit                 As Boolean
    Public blnPM_Debug                  As Boolean
    Public blnPM_DUF_Install            As Boolean
    Public blnPM_DebugLog               As Boolean
    Public blnPM_NoAskSessionResume     As Boolean
    Public blnPM_Session_Editor         As Boolean
    Public blnPM_Presentation           As Boolean
    Public blnPM_DUF_Editor             As Boolean
    Public blnPM_Open_Session           As Boolean
    Public blnPM_Open_DUF               As Boolean
'
' CONSTANTS
'
    '
    ' Programmaconstanten
    '
    Public Const AppName                As String = "Verbodrome V6"
    Public Const gcLoopTreshold         As Byte = 64
    Public Const blnTrialVersion        As Boolean = True
    '
    ' Database constanten
    '
    Public Const gcMdbName              As String = "Verbodrome V6.mdb"
    Public Const gcConnectionString     As String = "DSN=Verbodrome V6"
    '
    ' Tijdconstanten
    '
    Public Const gcInfPr                As Byte = 1
    Public Const gcInfPa                As Byte = 2
    Public Const gcIndPr                As Byte = 3
    Public Const gcIndPC                As Byte = 4
    Public Const gcIndIpf               As Byte = 5
    Public Const gcIndPQP               As Byte = 6
    Public Const gcIndPS                As Byte = 7
    Public Const gcIndPA                As Byte = 8
    Public Const gcIndFS                As Byte = 9
    Public Const gcIndFA                As Byte = 10
    Public Const gcCondPr               As Byte = 11
    Public Const gcCondPa               As Byte = 12
    Public Const gcSubjPr               As Byte = 13
    Public Const gcSubjPa               As Byte = 14
    Public Const gcSubjIpf              As Byte = 15
    Public Const gcSubjPQP              As Byte = 16
    Public Const gcImpPr                As Byte = 17
    Public Const gcImpPa                As Byte = 18
    Public Const gcPartPr               As Byte = 19
    Public Const gcPartPa               As Byte = 20
    Public Const gcGerPr                As Byte = 21
    Public Const gcGerPa                As Byte = 22
    '
    ' Bestandsextentieconstanten
    '
    Public Const gcSessionFileExtention         As String = ".vsf"
    Public Const gcDatabaseUpdateFileExtention  As String = ".vdu"
'
' ENUMS
'
    '
    ' Statusbar-panelen Enum
    '
    Public Enum StatusBarPanels
        pnlOefNr = 1
        pnlScore = 2
        pnlTijd = 3
        pnlTotTijd = 4
        pnlGemTijd = 5
    End Enum
    '
    ' Logboek Enum
    '
    Public Enum EventType
        ProgramStart = 0
        ProgramEnd = 1
        AskedQuestion = 2
        CorrectAnswer = 3
        WrongAnswer = 4
        Command = 5
        SelectionWindow = 6
        Error = 7
        Score = 8
        ResetAskedForms = 9
        TimeLimitOn = 10
        TimeLimitOff = 11
        ChangeSetting = 12
        DUFSuccesful = 13
        DebugLogFile = 14
    End Enum
    '
    ' Conversietype (converteren van ene aanduiding naar de andere)
    '
    Public Enum ConvertType
        AbsNumber = 0
        Abreviation = 2
        Full = 3
        ShortFull = 4
    End Enum
