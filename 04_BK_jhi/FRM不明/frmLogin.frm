VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "ログイン"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3525
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmd 
      Caption         =   "Command3"
      Height          =   345
      Left            =   1275
      TabIndex        =   6
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ログイン"
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "閉じる"
      Height          =   345
      Left            =   2310
      TabIndex        =   4
      Top             =   1140
      Width           =   975
   End
   Begin VB.TextBox txtPwd 
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Top             =   660
      Width           =   1965
   End
   Begin VB.TextBox txtLogin 
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   240
      Width           =   1965
   End
   Begin VB.Label Label2 
      Caption         =   "パスワード"
      Height          =   225
      Left            =   300
      TabIndex        =   2
      Top             =   690
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "ユーザ"
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   270
      Width           =   705
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
    End
End Sub
