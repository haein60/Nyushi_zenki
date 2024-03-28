VERSION 5.00
Begin VB.Form frmLogo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10410
   WindowState     =   2  'Maximized
   Begin VB.Label lblNyushi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1905"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   4680
      TabIndex        =   0
      Top             =   3720
      Width           =   5895
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    fMainForm.mnuTools.Enabled = False
    Dim index
    For index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(index).Enabled = False
    Next
End Sub

Private Sub Form_Load()
    LoadResStrings Me
    Me.Caption = LoadResString(1905)
    fMainForm.mnuTools.Enabled = False
End Sub
