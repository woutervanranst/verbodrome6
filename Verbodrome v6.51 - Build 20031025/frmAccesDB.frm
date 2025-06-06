VERSION 5.00
Begin VB.Form frmAccesDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verbodrome Database Editor"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdLookup 
      Caption         =   "&Lookup"
      Default         =   -1  'True
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtTable 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "Settings"
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1260
      Width           =   450
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Field:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   540
      Width           =   375
   End
   Begin VB.Label lblCriteria 
      AutoSize        =   -1  'True
      Caption         =   "Criteria:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   900
      Width           =   525
   End
   Begin VB.Label lblTable 
      AutoSize        =   -1  'True
      Caption         =   "Table:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   450
   End
End
Attribute VB_Name = "frmAccesDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' FORM
'
    Private Sub cmdExit_Click()
        Unload Me
        End
    End Sub
    '
    Private Sub cmdLookup_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmAccesDB.cmdLookup_Click"
        '
        If txtTable.Text <> "" And txtField.Text <> "" And txtCriteria.Text <> "" Then
            'Merci Danny voor deze procedure
            strSQL = "SELECT " & Me.txtField.Text & " FROM " & Me.txtTable.Text & " WHERE " & Me.txtCriteria.Text
            recVerbodrome.Open strSQL, gcConnectionString, adOpenForwardOnly, adLockBatchOptimistic ', adCmdText
            '
            With recVerbodrome
                If Not (.EOF And .BOF) Then
                    Do While Not .EOF
                        Me.txtValue.Text = recVerbodrome.Fields(0)
                        .MoveNext
                    Loop
                End If
                .Close
            End With
        End If
        '
        GoTo ExitProc
ErrorOccured::
        txtValue.Text = "Error"
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub cmdUpdate_Click()
        On Error GoTo ErrorOccured
        errProcStack errPush, "frmAccesDB.cmdUpdate_Click"
        '
        If MsgBox("Are you sure you want to update this database field?" & vbCrLf & "The previous data will be erased!", vbQuestion + vbYesNo, AppName) = vbYes Then
            strSQL = "SELECT " & Me.txtField.Text & " FROM " & Me.txtTable.Text & " WHERE " & Me.txtCriteria.Text
            recVerbodrome.Open strSQL, gcConnectionString, adOpenStatic, adLockOptimistic
            '
            With recVerbodrome
                If Not (.EOF And .BOF) Then
                    recVerbodrome.Fields(0) = Me.txtValue.Text
                    .Update
                End If
                .Close
            End With
        End If
        '
        GoTo ExitProc
ErrorOccured::
        txtValue.Text = "Error"
        If errHandler(Err.Number, Err.Description) = errNext Then Resume Next
ExitProc::
        errProcStack errPop, ""
    End Sub
    '
    Private Sub Form_Load()
        Me.Icon = LoadResPicture(107, vbResIcon) 'frmLoad.Icon
    End Sub
    '
    Private Sub Form_Unload(Cancel As Integer)
        Cancel = vbCancel
        Call cmdExit_Click
    End Sub
    '
    Private Sub txtValue_Change()
        cmdUpdate.Enabled = (txtValue.Text <> "")
    End Sub
