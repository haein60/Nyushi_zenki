VERSION 5.00
Begin VB.Form frmExamCheckPara 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmExamCheckPara_20211209BK.frx":0000
   ScaleHeight     =   3060
   ScaleWidth      =   7095
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.CommandButton cmdImput 
      Caption         =   "ÉCÉìÉ|Å[Ég"
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdDisp 
      Caption         =   "ï\é¶"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtJukenNumberTo 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      IMEMode         =   3  'µÃå≈íË
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "[iNendo]"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtJukenNumberFrom 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      IMEMode         =   3  'µÃå≈íË
      Left            =   1785
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "[iNendo]"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtNendo 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      IMEMode         =   3  'µÃå≈íË
      Left            =   1785
      MaxLength       =   4
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "[iNendo]"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "åèêî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Tag             =   "1804"
      Top             =   2340
      Width           =   945
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "éÛå±î‘çÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Tag             =   "1804"
      Top             =   1740
      Width           =   825
   End
   Begin VB.Label lblNendo 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "îNìx"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Tag             =   "1804"
      Top             =   1140
      Width           =   825
   End
End
Attribute VB_Name = "frmExamCheckPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisp_Click()

    If Trim(txtNendo) = "" Then
        MsgBox "îNìxÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢"
        Exit Sub
    End If
    gsExamCheckNendo = Trim(txtNendo)

    If Trim(txtJukenNumberFrom) = "" Then
        MsgBox "äJénéÛå±î‘çÜÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢"
        Exit Sub
    End If
    gsExamIDFrom = txtJukenNumberFrom

    If Trim(txtJukenNumberTo) = "" Then
        MsgBox "åèêîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢"
        Exit Sub
    End If
    gsExamIDTo = txtJukenNumberTo

    frmExamineeCheck.Show
    frmExamineeCheck.ZOrder 0

    Unload Me

End Sub

Private Sub cmdImput_Click()
'    frmExamineeImport.Visible = True
frmExamineeImport.txtNendo.Text = Me.txtNendo.Text
frmExamineeImport.Show 1
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    fMainForm.mnuTools.Enabled = False  ' disable tools menu
    Dim Index
    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Load()

Dim sSQL As String
Dim oRs As ADODB.Recordset
Dim l_int_CurrentPhase As Integer

On Error GoTo ErrProc

    Me.Caption = "éÛå±é“èÓïÒÉÅÉìÉeÉiÉìÉX"
    txtNendo.Text = g_int_CurrentNendo

    ' get the current phase
    sSQL = "SELECT iCurrentPhase FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag=1" & _
        " AND iCurrentPhase IS NOT NULL"
    Set oRs = g_obj_Conn.Execute(sSQL)

    If Not oRs.EOF Then
        l_int_CurrentPhase = oRs.Fields("iCurrentPhase").Value
        oRs.Close
        Set oRs = Nothing
    Else
        l_int_CurrentPhase = 1
        Set oRs = Nothing
    End If

    If l_int_CurrentPhase = 0 Then
        txtJukenNumberTo.Text = "50"
    Else
        txtJukenNumberTo.Text = "1"
    End If

'    txtJukenNumberFrom.SetFocus

Exit Sub
ErrProc:

End Sub

Private Sub txtJukenNumberFrom_KeyPress(KeyAscii As Integer)
    Call NumericOnly(Me, KeyAscii)
End Sub

Private Sub txtJukenNumberTo_KeyPress(KeyAscii As Integer)
    Call NumericOnly(Me, KeyAscii)
End Sub

Private Sub txtNendo_KeyPress(KeyAscii As Integer)
    Call NumericOnly(Me, KeyAscii)
End Sub
