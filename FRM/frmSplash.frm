VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4695
      Left            =   45
      TabIndex        =   0
      Tag             =   "1905"
      Top             =   0
      Width           =   7380
      Begin VB.PictureBox picLogo 
         Height          =   2385
         Left            =   510
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   2325
         ScaleWidth      =   1755
         TabIndex        =   1
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "1905"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   2670
         TabIndex        =   5
         Tag             =   "1050"
         Top             =   1200
         Width           =   2040
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2428"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6405
         TabIndex        =   4
         Tag             =   "1047"
         Top             =   2760
         Width           =   600
      End
      Begin VB.Label lblWarning 
         Caption         =   "2430"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Tag             =   "1046"
         Top             =   3720
         Width           =   6855
      End
      Begin VB.Label lblCompany 
         Caption         =   "2429"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4710
         TabIndex        =   3
         Tag             =   "1045"
         Top             =   3330
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmBrowse
'Author         :   Dileep Cherian
'Created On     :
'Description    :   This Screen shows the logo on startup
'Reference      :
'**************************************************************************************************
Private Sub Form_Load()
    'LoadResStrings Me
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    LoadResStrings Me
    Call g_void_SetFontProperties(Me)     ' set the font properties
End Sub

