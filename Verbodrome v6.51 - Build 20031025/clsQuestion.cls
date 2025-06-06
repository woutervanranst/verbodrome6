VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
Private intVerbe As Integer
Private bytTemps As Byte
Private bytPersonne As Byte
Private strAnswers() As String
Private blnFromSessionFile As Boolean
'
'   PROPERTIES
'
    ' Verbe
    Public Property Get Verbe() As String
        Verbe = VerbeText
    End Property
    '
    ' Temps
    Public Property Get Temps() As String
        Temps = TempsText
    End Property
    '
    ' Personne
    Public Property Get Personne() As String
        Personne = PersonneText
    End Property
    '
    ' VerbeIndex
    Public Property Get VerbeIndex() As Integer
        VerbeIndex = intVerbe
    End Property
    '
    Public Property Let VerbeIndex(intNewVerbIndex As Integer)
        intVerbe = intNewVerbIndex
    End Property
    '
    ' TempsIndex
    Public Property Get TempsIndex() As Byte
        TempsIndex = bytTemps
    End Property
    '
    Public Property Let TempsIndex(bytNewTempsIndex As Byte)
        bytTemps = bytNewTempsIndex
    End Property
    '
    ' PersonneIndex
    Public Property Get PersonneIndex() As Byte
        PersonneIndex = bytPersonne
    End Property
    '
    Public Property Let PersonneIndex(bytNewPersonneIndex As Byte)
        bytPersonne = bytNewPersonneIndex
    End Property
    '
    ' CorrectAnswer
    Public Property Get CorrectAnswer() As String
        CorrectAnswer = myFormes.Form(intVerbe, bytTemps, bytPersonne)
    End Property
    '
    ' PreviousAnswers
    Public Property Get PreviousAnswers() As String()
        PreviousAnswers = strAnswers
    End Property
    '
    Public Property Let PreviousAnswers(strAllPreviousAnswers() As String)
        strAnswers = strAllPreviousAnswers
    End Property
    '
    ' QuestionFromSessionFile
    Public Property Get FromSessionFile() As Boolean
        FromSessionFile = blnFromSessionFile
    End Property
    '
    Public Property Let FromSessionFile(blnNewFromSessionFile As Boolean)
        blnFromSessionFile = blnNewFromSessionFile
    End Property
    '
    ' QuestionText
    '   Er staat een kopie van deze functie in de frmHistory
    Public Property Get QuestionText() As String
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsQuestion.QuestionText"
        '
        If bytTemps >= gcIndPr And bytTemps <= gcSubjPQP Then
            'Met persoon
            QuestionText = GetTxt("QuestionTime")
            QuestionText = Replace(QuestionText, "$persoon$", PersonneText)
        Else
            'Zonder persoon
            QuestionText = GetTxt("QuestionTimeLess")
        End If
        '
        'Infinitief implementeren
        QuestionText = Replace(QuestionText, "$infinitief$", VerbeText)
        '
        'Tijd implementeren
        QuestionText = Replace(QuestionText, "$tijd$", TempsText)
        '
        'In het Frans: niet 'donnez le infinitif' maar 'donnez l'infinitif'
        If bytLanguage = 1 Then
            QuestionText = Replace(QuestionText, "le i", "l'i")
        End If
        '
        If blnPM_DebugLog Then Call LogDebug("Result: " & QuestionText)
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Property
'
' FUNCTIONS
'
    Private Function VerbeText() As String
        VerbeText = myFormes.GetInfinitive(intVerbe)
    End Function
    '
    Private Function TempsText() As String
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsQuestion.TempsText"
        '
        If bytTemps = gcImpPr Or bytTemps = gcImpPa Then
            'Imperatieven
            TempsText = LCase(Convert(AbsNumber, ShortFull, bytTemps) & " " & Choose(bytPersonne, "Sing.(-)", "1e Pl.(+)", "2e Pl(+)"))
        ElseIf bytTemps = gcPartPa Then
            'Participe Passé
            TempsText = LCase(Convert(AbsNumber, ShortFull, bytTemps) & IIf(bytPersonne = 1, "", " Pass."))
        Else
            'Alle andere tijden
            TempsText = LCase(Convert(AbsNumber, ShortFull, bytTemps))
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Private Function PersonneText() As String
        Dim strPersonneCaptions() As String
        strPersonneCaptions = Split(GetTxt("QuestionPersons"), ";")
        PersonneText = strPersonneCaptions(bytPersonne - 1)
    End Function
'
' SUBS
'
    Private Sub Class_Initialize()
        ReDim strAnswers(0)
    End Sub
    '
    Public Sub AddPreviousAnswer(strPreviousAnswer As String)
        If strAnswers(UBound(strAnswers)) <> "" Then ReDim Preserve strAnswers(LBound(strAnswers) To UBound(strAnswers) + 1)
        strAnswers(UBound(strAnswers)) = LCase(strPreviousAnswer)
    End Sub


