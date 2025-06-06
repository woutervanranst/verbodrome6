VERSION 5.00
Begin VB.Form frmBlank 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgWallpaper 
      Height          =   495
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmBlank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' FORM
'
    Private Sub Form_GotFocus()
        frmMain.Show
    End Sub
    '
    Private Sub Form_Load()
        Me.Left = 0
        Me.Top = 0
        Me.Width = Screen.Width
        Me.Height = Screen.Height
        '
        Call InitWallpaper
        '
        If blnPM_DebugLog Then Call LogDebug("frmBlank loaded")
        frmMain.Show
    End Sub
    '
    Private Sub Form_Resize()
        With imgWallpaper
            .Left = 0
            .Top = 0
            .Width = Me.Width
            .Height = Me.Height
        End With
    End Sub
    '
    Public Sub InitWallpaper()
        On Error GoTo ErrorOccured
        errProcStack errPush, "modDatabase.EmptyHistory"
        '
        If GetDBSetting("BackgroundType") = "Black" Then
            imgWallpaper.Picture = LoadPicture("")
            If blnPM_DebugLog Then Call LogDebug("Black background succesfully loaded")
        ElseIf GetDBSetting("BackgroundType") = "Wallpaper" Then
            imgWallpaper.Picture = LoadPicture(regQuery_A_Key(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper"))
            If blnPM_DebugLog Then Call LogDebug("Windows wallpaper background succesfully loaded")
        End If
        '
        GoTo ExitProc
ErrorOccured::
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub Form_Unload(Cancel As Integer)
        If blnPM_DebugLog Then Call LogDebug("frmBlank unloaded")
    End Sub
