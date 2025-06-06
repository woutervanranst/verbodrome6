Attribute VB_Name = "modMain"
Option Explicit
'
'**********************************************************************************************
'* Verbodrome (Bescherelle) Oefenprogramma V6 (c) 2002 - 2003 by Wouter Van Ranst             *
'* ------------------------------------------------------------------------------             *
'*                                                                                            *
'* V6.2:                                                                                      *
'* -----                                                                                      *
'* Geprogrammeerd in de maand Augustus gedurende 126 uur:                                     *
'* Datum          Tot. Day    Tot. Month                                                      *
'* 20/07/2002     4:22                                                                        *
'* 31/07/2002     6:05        10:27                                                           *
'*  1/08/2002     7:57                                                                        *
'*  2/08/2002    10:26                                                                        *
'*  3/08/2002     1:28                                                                        *
'*  4/08/2002     1:29                                                                        *
'*  5/08/2002     4:47                                                                        *
'*  6/08/2002     0:46                                                                        *
'*  7/08/2002     5:40                                                                        *
'*  8/08/2002    12:35                                                                        *
'*  9/08/2002    10:08                                                                        *
'* 10/08/2002     3:12                                                                        *
'* 11/08/2002     3:56                                                                        *
'* 12/08/2002    10:39                                                                        *
'* 13/08/2002    10:56                                                                        *
'* 14/08/2002    19:01                                                                        *
'* 15/08/2002     0:16    'Knokke -> 18/08                                                    *
'* 18/08/2002     0:22                                                                        *
'* 19/08/2002     2:47                                                                        *
'* 20/08/2002     6:26                                                                        *
'* 21/08/2002     3:01                                                                        *
'* 22/08/2002     1:12                                                                        *
'* 23/08/2002     0:34                                                                        *
'* 24/08/2002     0:12        117:50:00                                                       *
'*  2/09/2002     0:17                                                                        *
'*  4/09/2002     0:28                                                                        *
'*  7/09/2002     0:09        0:54                                                            *
'*  8/10/2002     0:08                                                                        *
'* 11/10/2002     0:13        0:21                                                            *
'*  9/11/2002     0:06                                                                        *
'* 23/11/2002     0:34                                                                        *
'* 24/11/2002     0:08                                                                        *
'* 26/11/2002     0:01        0:49                                                            *
'*                                                                                            *
'* V6.3:                                                                                      *
'* -----                                                                                      *
'* 28/11/2002     2:45                                                                        *
'* 30/11/2002     0:04                                                                        *
'*  6/12/2002     0:16                                                                        *
'*  8/12/2002     0:49                                                                        *
'*  9/12/2002     2:59                                                                        *
'* 10/12/2002     0:11                                                                        *
'* 13/12/2002     6:31    'Vrijdag: na exaam biologie, maandag Wiskunde                       *
'* 14/12/2002     2:28                                                                        *
'* 16/12/2002     4:07                                                                        *
'* 17/12/2002     6:51                                                                        *
'* 18/12/2002     0:21                                                                        *
'* 19/12/2002    10:44    'Dag na de examens                                                  *
'* 20/12/2002     0:34                                                                        *
'* 21/12/2002     4:34                                                                        *
'* 22/12/2002     1:19                                                                        *
'* 23/12/2002     0:25                                                                        *
'* 24/12/2002     3:04                                                                        *
'* 25/12/2002     1:17                                                                        *
'* 26/12/2002     2:46                                                                        *
'* 27/12/2002     5:55                                                                        *
'* 30/12/2002     0:02                                                                        *
'* 31/12/2002     0:11                                                                        *
'*  1/01/2003     0:52                                                                        *
'*  2/01/2003     0:21       56:41                                                            *
'*                                                                                            *
'* Autorun:                                                                                   *
'* --------                                                                                   *
'* 30/12/2002     0:08                                                                        *
'* 31/12/2002     3:14                                                                        *
'*  1/01/2003     0:05                                                                        *
'*                                                                                            *
'* V6.31                                                                                      *
'* -----                                                                                      *
'* 15/04/2003     8:01                                                                        *
'* 16/04/2003     6:47                                                                        *
'* 17/04/2003     5:59                                                                        *
'* 18/04/2003     0:08                                                                        *
'* 19/04/2003     0:10       20:47                                                            *
'*                                                                                            *
'* V6.4                                                                                       *
'* ----                                                                                       *
'* 09/07/2003  0:50                                                                           *
'* 22/07/2003  0:46                                                                           *
'* 24/07/2003  1:10                                                                           *
'* 31/07/2003  0:06           3:13                                                            *
'* 16/08/2003  0:41         'Knokke                                                           *
'* 17/08/2003  2:25         'Knokke                                                           *
'* 18/08/2003  1:07         'Knokke                                                           *
'* 19/08/2003  3:37         'Knokke                                                           *
'* 20/08/2003  4:16         'Knokke                                                           *
'* 21/08/2003  0:59         'Knokke                                                           *
'* 22/08/2003  0:14         'Knokke                                                           *
'* 23/08/2003  3:39         'Knokke                                                           *
'* 24/08/2003  0:48         'Knokke                                                           *
'* 25/08/2003  4:05                                                                           *
'* 26/08/2003  4:00                                                                           *
'*                                                                                            *
'* V6.5                                                                                       *
'* ----                                                                                       *
'* 27/08/2003  1:15                                                                           *
'* 28/08/2003  8:41                                                                           *
'* 29/08/2003  5:09                                                                           *
'* 30/08/2003  11:53                                                                          *
'* 31/08/2003  10:54         63:39                                                            *
'* 01/09/2003  4:04                                                                           *
'* 02/09/2003  2:48                                                                           *
'* 03/09/2003  2:15                                                                           *
'* 07/09/2003  1:12                                                                           *
'* 08/09/2003  0:47                                                                           *
'* 09/09/2003  1:29                                                                           *
'* 10/09/2003  2:09                                                                           *
'* 12/09/2003  2:03                                                                           *
'* 13/09/2003  3:07                                                                           *
'* 18/09/2003  0:04                                                                           *
'* 19/09/2003  0:58                                                                           *
'* 20/09/2003  0:38                                                                           *
'* 23/09/2003  0:19                                                                           *
'* 24/09/2003  0:51                                                                           *
'* 25/09/2003  0:25                                                                           *
'* 26/09/2003  0:29                                                                           *
'* 27/09/2003  0:21                                                                           *
'* 29/09/2003  0:44          24:43                                                            *
'* 01/10/2003  0:19                                                                           *
'* 02/10/2003  0:45          1:04                                                             *
'*                                                                                            *
'* Totale werktijd me alle geprul: 284 uur en 44 minuten                                      *
'*                                                                                            *
'* Dit programma is de kroon op mijn VB-werk, dat ik nu afsluit, om verder te gaan met        *
'* C#.NET Dit programma is zo perfect mogelijk geschreven, en talloze keren gedebugd.         *
'*                                                                                            *
'* Ik weet niet wie deze comment leest, misschien ikzelf binnen 10 jaar, als ik dan nog besta *
'* maar denk nog eens aan mij.                                                                *
'*                                                                                            *
'* De conversie naar Verbodrome ( naam uitgevonden door Christian Wildiers op 21/01/2003 )    *
'* in de dagen daarna.                                                                        *
'*                                                                                            *
'* Groetjes aan mijzelf ;)            Wouter, 8/10/2002, 2/01/2003, 19/04/2003, 26/08/2003    *
'*                                            7/10/2003                                       *
'*                                                                                            *
'**********************************************************************************************
'
' SUBS
'
    Sub Main()
        On Error GoTo ErrorOccured
        errProcStack errPush, "modMain.Main"
        '
        'GLOBALE OPSTARTPROCEDURE
        '
        '1. Controleren of er al een andere instantie draait
        If App.PrevInstance Then
            MsgBox GetTxt("OtherInstanceRunning"), vbCritical, AppName
            End
        End If
        '
        '2. Leuk extraatje
        If Year(Now) > 2009 Then
            MsgBox "That they are still runing Verbodrome (Bescherelle) in " & Year(Now) & vbCrLf & vbCrLf & _
            "Tell me, how did .NET go?", vbInformation, AppName
        End If
        '
        '3. Opslaan waar de EXE van dit programma staat (voor eventuele client-programmas of patches)
        SaveSetting AppName, "AppPath", "AppPath", App.Path
        SaveSetting AppName, "AppPath", "BuildNr", App.Comments
        '
        '4. Databasekoppeling controleren, en indien niet bestaande, aanmaken
        Call modRegistry.CheckDSN(App.EXEName)
        Call GetSessionID
        dtmProgramStart = Now
        If Dir(App.Path & "\LocalUser", vbDirectory) <> "LocalUser" Then MkDir App.Path & "\LocalUser"
        '
        '5. Taal ophalen en instellen
        bytLanguage = GetDBSetting("Language", 99)
        If bytLanguage = 99 Then
            frmChooseLanguage.Show 1
            LetDBSetting "Language", CStr(bytLanguage)
        End If
        '
        '6. Parameters controleren
        If InStr(VBA.Command, "/debug") Then blnPM_Debug = True
        If InStr(VBA.Command, "/dbedit") Then blnPM_DBEdit = True
        If InStr(VBA.Command, "/logdebuginfo") Then blnPM_DebugLog = True: If Dir(App.Path & "\Debug", vbDirectory) <> "Debug" Then MkDir App.Path & "\Debug"
        If InStr(VBA.Command, "/noasksessionresume") Then blnPM_NoAskSessionResume = True
        If InStr(VBA.Command, "/presentation") Then blnPM_Presentation = True
        If InStr(VBA.Command, "/dufcreator") Then blnPM_DUF_Editor = True
        If InStr(VBA.Command, "/dufinstall") Then blnPM_DUF_Install = True
        If InStr(VBA.Command, "/sessioncreator") Then blnPM_Session_Editor = True
        If InStr(VBA.Command, "/openvsf") Then blnPM_Open_Session = True
        If InStr(VBA.Command, "/openvdu") Then blnPM_Open_DUF = True
        '
        '8. Trial versie
        Call ShowRemainingDays
        '
        '9. Laadform tonen
        frmLoad.Show
        Screen.MousePointer = vbHourglass
        DoEvents
        '
        '7. Bestandstypen registreren
        Call SetFileAssociation(gcSessionFileExtention, App.Path + "\" + App.EXEName + ".exe", "Verbodrome.SessionFile", GetTxt("SessionFile"), App.Path + "\" + App.EXEName + ".exe,3", "/openvsf")
        Call SetFileAssociation(gcDatabaseUpdateFileExtention, App.Path + "\" + App.EXEName + ".exe", "Verbodrome.DatabaseUpdateFile", GetTxt("UpdateFile"), App.Path + "\" + App.EXEName + ".exe,2", "/openvdu")
        '
        '10. Databasekoppeling openen
        gcnConnection.ConnectionString = gcConnectionString
        gcnConnection.Open
        blnHistoryEnabled = (GetDBSetting("HistoryStatus", "Enabled") = "Enabled")
        '
        '11. Werkwoorden inladen
        Call LoadVerbs
        '
        ' Zie ook naar deze reeks in clsSessions (nodig om de statusbar enkel in de juiste modus te koppelen)
        If blnPM_DUF_Install Or blnPM_Open_DUF Or blnPM_DBEdit Or blnPM_Session_Editor Or blnPM_DUF_Editor Then
            Unload frmLoad
            Screen.MousePointer = vbDefault
        End If
        '
        If blnPM_DUF_Install Or blnPM_Open_DUF Then
            '12. DUF-interface tonen
            If blnPM_DebugLog Then Call LogDebug("Showing DUF Install-Interface")
            Call LoadDUF
            GoTo ExitProc
        ElseIf blnPM_DBEdit Then
            '12. Toegang tot de database verschaffen
            If blnPM_DebugLog Then Call LogDebug("Showing Database access")
            frmAccesDB.Show
            GoTo ExitProc
        ElseIf blnPM_Session_Editor Then
            '12. Session Creator interface tonen
            If blnPM_DebugLog Then Call LogDebug("Showing Session Creator-Interface")
            frmCreatorSession.Show
            GoTo ExitProc
        ElseIf blnPM_DUF_Editor Then
            '12. DUF Creator interface tonen
            If blnPM_DebugLog Then Call LogDebug("Showing DUF Creator-Interface")
            frmCreatorDUF.Show
            GoTo ExitProc
        Else
            '12. Hoofdscherm tonen
            '
            'Controleren of de history niet (te) vol zit
            If NumberOfHistoryItems > 50000 Then
                If MsgBox(GetTxt("HistoryFull"), vbQuestion + vbYesNo, AppName) = vbYes Then
                    'Historiekgegevens verwijderen
                    Call EmptyHistory
                End If
            End If
            '
            If blnPM_DebugLog Then Call LogDebug("Showing Main Program")
            Call LogEvent(ProgramStart)
            '
            If blnPM_Open_Session Then
                mySession.SessionFile = VBA.Command
                mySession.SessionFile = Mid(mySession.SessionFile, InStr(mySession.SessionFile, Chr(34)) + 1)
                mySession.SessionFile = Left(mySession.SessionFile, InStr(mySession.SessionFile, Chr(34)) - 1)
                '
                Call mySession.OpenSessionFile
            Else
                Call mySession.PromptOpenLastSession
            End If
            '
            DoEvents
            Load frmMain
            Unload frmLoad
            frmMain.Show
            '
            GoTo ExitProc
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub LoadDUF()
        On Error GoTo ErrorOccured
        errProcStack errPush, "modMain.LoadDUF"
        '
        Dim clsMyDUFUpdate As clsDUFInstaller
        Set clsMyDUFUpdate = New clsDUFInstaller
        Dim strDUFFile As String
        '
        If blnPM_Open_DUF = True Then
            'Bestandsnaam in de argumentenlijst
            strDUFFile = VBA.Command
            strDUFFile = Mid(strDUFFile, InStr(strDUFFile, Chr(34)) + 1)
            strDUFFile = Left(strDUFFile, InStr(strDUFFile, Chr(34)) - 1)
            '
            If strDUFFile = "" Then GoTo ExitProc
        Else
            'Geen bestand gespecifieerd
            With frmCommonDialog.dlgCommonDialog
                .DialogTitle = GetTxt("SelectDUF")
                .Filter = GetTxt("UpdateFile") & " (*" & gcDatabaseUpdateFileExtention & ")|*" & gcDatabaseUpdateFileExtention
                .CancelError = True
                .ShowOpen
                '
                If .FileName <> "" Then
                    strDUFFile = .FileName
                Else
                    GoTo ExitProc
                End If
            End With
        End If
        '
        If MsgBox(GetTxt("DUFInstaller_ProceedPrompt"), vbQuestion + vbYesNo) = vbYes Then
            Call clsMyDUFUpdate.InstallDUF(strDUFFile)
        End If
        '
        End
        '
        GoTo ExitProc
ErrorOccured::
        If Err.Number = 32755 Then
            'Op Cancel gedrukt
            GoTo ExitProc
        Else
            If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
        End If
ExitProc::
        errProcStack errPop, ""
        End
    End Sub
    '
    Public Sub ExitProgram()
        On Error GoTo ErrorOccured
        errProcStack errPush, "modMain.ExitProgram"
        '
        If frmMain.mnuViewHideBackGround.Checked Then Unload frmBlank
        frmMain.Visible = False
        DoEvents
        '
        Screen.MousePointer = vbHourglass
        '
        '1. Laadform tonen
        frmLoad.Show
        frmLoad.lblCurrentJob = GetTxt("ExitingProgram")
        frmLoad.MousePointer = vbHourglass
        DoEvents
        '
        '2. Score wegschrijven
        strSQL = "SELECT Id, BehSc, MaxSc FROM Score"
        recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
        '
        With recVerbodrome
            While Not .EOF
                !BehSc = !BehSc + CLng(mySession.myScoreBoard.Score(!Id))
                !MaxSc = !MaxSc + CLng(mySession.myScoreBoard.NumberOfEntries(!Id))
                '
                .Update
                .MoveNext
            Wend
            '
            .Close
        End With
        '
        '4. Event loggen
        Call LogEvent(ProgramEnd)
        '
        '5. Huidige sessie opslaan
        mySession.SessionFile = App.Path & "\LastSession.vsf"
        'TODO: Alle OOPs = Nothing
        Call mySession.SaveSessionFile
        Set mySession = Nothing
        Set myFormes = Nothing
        Set frmMain = Nothing
        '
        '6. Programma afsluiten
        Unload frmLoad
        Screen.MousePointer = vbDefault
        End
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
