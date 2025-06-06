VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScoreBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
' NOOT: Bij het opmaken van een statistiekrapport: de Tries bij Correct en Not(Correct) zijn NIET dezelfde!!!!!
'
Private intMaxScore(1 To 22)            As Integer
Private intBehScore(1 To 22)            As Integer
'
Private intProbabilityRatings(1 To 22)  As Integer
'
'   SUBS
'
    Public Sub AddEntry(blnCorrect As Boolean, bytAfterTries As Byte, intTemps As Integer)
        'Voegt een entry toe aan het scorerapport
        intMaxScore(intTemps) = intMaxScore(intTemps) + 1
        If blnCorrect Then intBehScore(intTemps) = intBehScore(intTemps) + 1
    End Sub
    '
    Public Sub Reset()
        'Alle punten resetten
        Dim i As Integer
        '
        For i = LBound(intMaxScore) To UBound(intMaxScore)
            intMaxScore(i) = 0
            intBehScore(i) = 0
        Next i
    End Sub
'
'   FUNCTIES
'
    Public Function Score(intTemps As Integer, Optional intMaximum As Integer) As Integer
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsScoreBoard.Score"
        '
        'Geeft de score van item intTemps weer, op schaal intMaximum
        'Als intMaximum is, --> optional, aantal juist beantwoorde vragen teruggeven
        If intMaximum = 0 Then
            Score = intBehScore(intTemps)
            If blnPM_DebugLog Then Call LogDebug("Result: " & Score & ", no maximum specified")
        Else
            Score = intBehScore(intTemps) / intMaxScore(intTemps) * intMaximum
            If blnPM_DebugLog Then Call LogDebug("Result: " & Score & ", Maximum: " & intMaximum)
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Public Function NumberOfEntries(intTemps As Integer) As Integer
        'Geeft het aantal scoreentries van intTemps weer
        NumberOfEntries = intMaxScore(intTemps)
    End Function
    '
    Public Function TotalNumberOfEntries() As Integer
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsScoreBoard.TotalNumberOfEntries"
        '
        'Geeft het totale aantal entries weer
        Dim i As Integer
        '
        For i = LBound(intMaxScore) To UBound(intMaxScore)
            TotalNumberOfEntries = TotalNumberOfEntries + intMaxScore(i)
        Next i
        '
        If blnPM_DebugLog Then Call LogDebug("Result: " & TotalNumberOfEntries)
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Public Function TotalScore(intMaximum As Integer) As Integer
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsScoreBoard.TotalScore"
        '
        'Heeft de totaalscore weer, op schaal intMaximum
        Dim i As Integer
        Dim intTotalMaxScore As Integer
        Dim intTotalBehScore As Integer
        '
        For i = LBound(intMaxScore) To UBound(intMaxScore)
            intTotalMaxScore = intTotalMaxScore + intMaxScore(i)
            intTotalBehScore = intTotalBehScore + intBehScore(i)
        Next i
        '
        TotalScore = intTotalBehScore / intTotalMaxScore * intMaximum
        '
        If blnPM_DebugLog Then Call LogDebug("Result: " & TotalScore & ", Maximum: " & intMaximum)
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Public Function ProbabilityChain() As String()
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsScoreBoard.ProbabilityChain"
        '
        If GetDBSetting("AutomaticProbability", True) Then
            'Dynamische probabiliteit
            Call LoadDynamicProbabilities
        Else
            'Statische probabiliteit
            Call LoadStaticProbabilities
        End If
        '
        Dim strProbabilityChain As String
        Dim i As Byte
        '
        For i = 1 To 22
            strProbabilityChain = strProbabilityChain & Replace(Space(intProbabilityRatings(i)), " ", i & ",")
        Next i
        '
        strProbabilityChain = Left(strProbabilityChain, Len(strProbabilityChain) - 1)
        ProbabilityChain = Split(strProbabilityChain, ",")
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Private Function LoadStaticProbabilities()
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsScoreBoard.LoadStaticProbabilities"
        '
        'Kopie van deze functie in frmSelection
        '
        'Probabiliteit inladen van de database
        strSQL = "SELECT Name, Probability FROM Allowance WHERE Len(Name)>2"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockReadOnly
        '
        With recVerbodrome
            While Not .EOF
                intProbabilityRatings(Convert(Abreviation, AbsNumber, !Name)) = !Probability
                .MoveNext
            Wend
            .Close
        End With
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Function
    '
    Private Sub LoadDynamicProbabilities()
        On Error GoTo ErrorOccured
        errProcStack errPush, "clsScoreBoard.LoadDynamicProbabilities"
        '
        'Probabiliteit dynamisch berekenen aan de hand van de score
        '   NOOT :  Dit model is goed genoeg.
        '           Voor nog meer accuraatheid: bereken één probabiliteit pér unieke score
        '           (bij gelijke score is de probabiliteit nu verschillend)
        '   NOOT :  Bij /debug mode gaat de bias uit naar de eerst ingelezen tijd, waardoor het LIJKT alsof dit algoritme niet goed werkt, maar in de dagdagelijkse situaties komt dit niet voor
        '
        Dim i As Integer, j As Byte
        '
        '1. Som berekenen van het aantal toegelaten tijden
        Dim bytSumActiveTenses As Byte
        'For i = 1 To UBound(bytAllowedTemps)
        '    bytSumActiveTenses = bytSumActiveTenses + i
        For i = 0 To UBound(bytAllowedTemps)
            bytSumActiveTenses = bytSumActiveTenses + i + 1
        Next i
        '
        '2. Werkvariabele initaliseren
        Dim bytDynProbability() As Byte
        ReDim bytDynProbability(0 To UBound(bytAllowedTemps), 0 To 1)
        'bytDynProbability(x, 0 to 1)
        '   0: Absolute tijdsnotatie
        '   1: Procentuele score
        '
        For i = LBound(bytAllowedTemps) To UBound(bytAllowedTemps)
            '2.1 Absolute tijdsnotatie van deze tijd opslaan
            bytDynProbability(i, 0) = bytAllowedTemps(i)
            '
            '2.2 Scores berekenen
            If NumberOfEntries(CByte(bytAllowedTemps(i))) = 0 Then
                'Deze tijd is nog niet gevraagd
                ' --> Maximale probabiliteit
                bytDynProbability(i, 1) = 0
            Else
                'Deze tijd is al gevraagd
                bytDynProbability(i, 1) = CByte(Round(Score(CByte(bytAllowedTemps(i))) / NumberOfEntries(CByte(bytAllowedTemps(i))) * 100, 0))
            End If
        Next i
        '
        '3. Werkvariabele sorteren op procentuele score
        Dim bytTempValue(0 To 1) As Byte
        For i = UBound(bytDynProbability) To LBound(bytDynProbability) Step -1
            For j = LBound(bytDynProbability) + 1 To i
                If bytDynProbability(j - 1, 1) < bytDynProbability(j, 1) Then
                    'Swap
                    bytTempValue(0) = bytDynProbability(j - 1, 0)
                    bytTempValue(1) = bytDynProbability(j - 1, 1)
                    '
                    bytDynProbability(j - 1, 0) = bytDynProbability(j, 0)
                    bytDynProbability(j - 1, 1) = bytDynProbability(j, 1)
                    '
                    bytDynProbability(j, 0) = bytTempValue(0)
                    bytDynProbability(j, 1) = bytTempValue(1)
                End If
            Next j
        Next i
        '
        '4. Dynamische probabiliteit berekenen
        For i = LBound(bytDynProbability) To UBound(bytDynProbability)
            intProbabilityRatings(bytDynProbability(i, 0)) = CInt((2 * (i + 1)) / bytSumActiveTenses * 100)
        Next i
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
'
'   PROPERTIES
'
    ' BehSc
    Public Property Let BehSc(intTemps As Integer, intScore As Integer)
        intBehScore(intTemps) = intScore
    End Property
    '
    ' MaxSc
    Public Property Let MaxSc(intTemps As Integer, intScore As Integer)
        intMaxScore(intTemps) = intScore
    End Property

