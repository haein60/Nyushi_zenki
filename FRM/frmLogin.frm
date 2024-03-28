VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1052"
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "1061"
      Height          =   360
      Left            =   2100
      TabIndex        =   5
      Tag             =   "1056"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "1060"
      Default         =   -1  'True
      Height          =   360
      Left            =   495
      TabIndex        =   4
      Tag             =   "1055"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1305
      TabIndex        =   3
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "1903"
      Height          =   248
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Tag             =   "1054"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "1902"
      Height          =   248
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Tag             =   "1053"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long


Public OK As Boolean
Private Sub Form_Load()
    Dim sBuffer As String
    Dim lSize As Long


    LoadResStrings Me
    Me.Caption = LoadResString(1901)

    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        txtUserName.Text = Left$(sBuffer, lSize)
    Else
        txtUserName.Text = vbNullString
    End If
End Sub



Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    'ToDo: create test for correct password
    'check for correct password
    If txtPassword.Text = "" Then
        OK = True
        Me.Hide
    Else
        MsgBox LoadResString(1904), , LoadResString(1901)
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If
End Sub

