VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "出力中"
   ClientHeight    =   1710
   ClientLeft      =   3570
   ClientTop       =   5235
   ClientWidth     =   8160
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStop 
      Caption         =   "中止"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblMsg 
      Caption         =   "Now　Printing・・・・・・・・・・"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   26.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStop_Click()

    gbStop = True

End Sub

Private Sub Form_Load()

    Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2

End Sub
