VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmHighSchoolType 
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14010
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmHighSchoolType.frx":0000
   ScaleHeight     =   9795
   ScaleWidth      =   14010
   Tag             =   "1101"
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.CommandButton Command1 
      Caption         =   "2463"
      Height          =   350
      Left            =   5880
      TabIndex        =   44
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox txtZipCodeId 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   43
      Tag             =   "[iZipCodeId]"
      Top             =   2040
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgsearchgrid 
      Height          =   3855
      Left            =   240
      TabIndex        =   42
      Top             =   5040
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6800
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cboLetterFlag 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   10335
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   10
      Tag             =   "[iLetterFlag]"
      Top             =   2025
      Width           =   1950
   End
   Begin VB.ComboBox cboHighSchoolRecommendation 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   10335
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   7
      Tag             =   "[iHighSchoolRecommendation]"
      Top             =   1545
      Width           =   1950
   End
   Begin VB.TextBox txtHighSchoolDropRecommendation 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   10320
      MaxLength       =   4
      TabIndex        =   22
      Tag             =   "[iHighSchoolDropRecommendationYear]"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtHighSchoolRecommendationYear2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   10320
      MaxLength       =   4
      TabIndex        =   26
      Tag             =   "[iHighSchoolRecommendationYear2]"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtHighSchoolRecommendationYear1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   24
      Tag             =   "[iHighSchoolRecommendationYear1]"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtTelephoneNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3960
      MaxLength       =   24
      TabIndex        =   16
      Tag             =   "[vTelephoneNo]"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtHighSchoolName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   5
      Tag             =   "[vHighSchoolName]"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtFaxNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   10320
      MaxLength       =   10
      TabIndex        =   18
      Tag             =   "[vFaxNo]"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtAddress1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   12
      Tag             =   "[vAddress1]"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtRepresentativename 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   20
      Tag             =   "[vRepresentativename]"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtAddress2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   10320
      MaxLength       =   15
      TabIndex        =   14
      Tag             =   "[vAddress2]"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtHighSchoolCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   10320
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "[vHighSchoolCode]"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtHighSchoolId 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   1
      Tag             =   "[iHighSchoolId]"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   14
      Left            =   10080
      TabIndex        =   41
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHighSchoolDropRecommendation 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1115"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   6480
      TabIndex        =   21
      Tag             =   "1115"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   13
      Left            =   10080
      TabIndex        =   40
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHighSchoolRecommendationYear2 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1114"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   6480
      TabIndex        =   25
      Tag             =   "1114"
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label lblHighSchoolRecommendationYear1 
      AutoSize        =   -1  'True
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1113"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   240
      TabIndex        =   23
      Tag             =   "1113"
      Top             =   3960
      Width           =   3315
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   3720
      TabIndex        =   39
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrorMsg 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   4560
      Visible         =   0   'False
      Width           =   12015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   37
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   36
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblZipCodeId 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1106"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Tag             =   "1106"
      Top             =   2040
      Width           =   3315
   End
   Begin VB.Label lblHighSchoolRecommendation 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1105"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   6480
      TabIndex        =   6
      Tag             =   "1105"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblLetterFlag 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1112"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Tag             =   "1112"
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label lblTelephoneNo 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1109"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Tag             =   "1109"
      Top             =   3000
      Width           =   3315
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   10080
      TabIndex        =   35
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   3720
      TabIndex        =   34
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   10080
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHighSchoolName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1104"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Tag             =   "1104"
      Top             =   1560
      Width           =   3315
   End
   Begin VB.Label lblFaxNo 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1110"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Tag             =   "1110"
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label lblAddress1 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1107"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Tag             =   "1107"
      Top             =   2520
      Width           =   3315
   End
   Begin VB.Label lblRepresentativename 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1111"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   240
      TabIndex        =   19
      Tag             =   "1111"
      Top             =   3480
      Width           =   3315
   End
   Begin VB.Label lblAddress2 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1108"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Tag             =   "1108"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   10080
      TabIndex        =   31
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   10080
      TabIndex        =   30
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   3720
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   10080
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHighSchoolCode 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1103"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Tag             =   "1103"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lblHighSchoolId 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1102"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Tag             =   "1102"
      Top             =   1080
      Width           =   3315
   End
End
Attribute VB_Name = "frmHighSchoolType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmHighSchoolType
'Author         :   Vishal Kamath
'Created On     :   10/8/01
'Description    :   This form makes a provision for master maintenance of tbSTRHighschoolType Table.
'Reference      :   Functional Specs Of MasterMaintenance Ver 1.0
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History   -   04/04/2002  -   Dileep Cherian
'On activation of the form, only the "new" and "query" toolbar icons should be enabled
'**************************************************************************************************
Dim m_bload As Boolean
Public m_TableName As String
Public m_colFieldDetails As New FieldDetails
Public m_bDirty As Boolean
Public m_bChangeOn As Boolean
Public m_lngSearchGridTop As Long
Public m_bMode As String
Public m_lngCurrentRow As Long
Public m_ComboDetails As New ComboCollection
Public f_int_PrevRow As Long

Private Sub chkHighSchoolRecommendation_Click()
    SetChange
End Sub

Private Sub chkLetterFlag_Click()
    SetChange
End Sub

Private Sub cboHighSchoolRecommendation_Click()
    SetChange
End Sub

Private Sub cboLetterFlag_Click()
    SetChange
End Sub

Private Sub cboZipCode_Click()
    SetChange
End Sub

Private Sub Command1_Click()
dlgZipHighschool.Show 1
End Sub

Private Sub Form_Activate()
    Dim lngRow As Integer
    Dim Index As Integer
    On Error GoTo ErrHandler
    
    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       If Index = 1 Or Index = 6 Then
            fMainForm.Toolbar1.Buttons(Index).Enabled = True
        Else
            fMainForm.Toolbar1.Buttons(Index).Enabled = False
        End If
    Next
    
    If m_bload = False Then
        m_lngSearchGridTop = 3950 ' changed by team
        
        'initialize the search grid
        InitializeSearchGrid
        'populate the grid with empty row and column headings
        SearchRecords True   ' by default, the form will be "search" mode
        fMainForm.mnuToolsSave.Enabled = False
        fMainForm.mnuToolsCancel.Enabled = False
        
        fMainForm.Toolbar1.Buttons("Save").Enabled = False
        fMainForm.Toolbar1.Buttons("Cancel").Enabled = False
        
    End If
    m_bload = True
    fMainForm.mnuTools.Enabled = True
        
Exit Sub
ErrHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Load()
    Dim l_str_sqlZip As String
    Dim l_obj_rsZipCode As New ADODB.Recordset
    On Error GoTo ErrHandler
    
    LoadResStrings Me
    Me.Caption = LoadResString(1101)
    Call g_void_SetFontProperties(Me)     ' set the font properties
    ' set the table name
    m_TableName = "tbSTEHighSchoolType"
    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained
    With m_colFieldDetails
        .Add "[iHighSchoolId]", "[" & LoadResString(1102) & "]", 1, True, True, "", "INTEGER", 1600, "", "[iHighSchoolId]"
        .Add "[vHighSchoolCode]", "[" & LoadResString(1103) & "]", 2, True, False, "", "STRING", 1900, "", "[vHighSchoolCode]"
        .Add "[vHighSchoolName]", "[" & LoadResString(1104) & "]", 3, True, False, "", "STRING", 2000, "", "[vHighSchoolName]"
        .Add "[iHighSchoolRecommendation]", "[" & LoadResString(1105) & "]", 4, False, False, "", "COMBO", 3000, "", "[iHighSchoolRecommendation]"
        .Add "[iZipCodeId]", "[" & LoadResString(1106) & "]", 5, False, False, "", "INTEGER", 1300, "", "[iZipCodeId]"
        .Add "[vAddress1]", "[" & LoadResString(1107) & "]", 6, False, False, "", "STRING", 1300, "", "[vAddress1]"
        .Add "[vAddress2]", "[" & LoadResString(1108) & "]", 7, False, False, "", "STRING", 1300, "", "[vAddress2]"
        .Add "[vTelephoneNo]", "[" & LoadResString(1109) & "]", 8, False, False, "", "STRING", 1800, "", "[vTelephoneNo]"
        .Add "[vFaxNo]", "[" & LoadResString(1110) & "]", 9, False, False, "", "STRING", 800, "", "[vFaxNo]"
        .Add "[vRepresentativeName]", "[" & LoadResString(1111) & "]", 10, False, False, "", "STRING", 2300, "", "[vRepresentativeName]"
        .Add "[iLetterFlag]", "[" & LoadResString(1112) & "]", 11, False, False, "", "COMBO", 1300, "", "[iLetterFlag]"
        .Add "[iHighSchoolRecommendationYear1]", "[" & LoadResString(1113) & "]", 12, False, False, "", "INTEGER", 3600, "", "[iHighSchoolRecommendationYear1]"
        .Add "[iHighSchoolRecommendationYear2]", "[" & LoadResString(1114) & "]", 13, False, False, "", "INTEGER", 3600, "", "[iHighSchoolRecommendationYear2]"
        .Add "[iHighSchoolDropRecommendationYear]", "[" & LoadResString(1115) & "]", 14, False, False, "", "INTEGER", 3900, "", "[iHighSchoolDropRecommendationYear]"
    End With
        
    cboHighSchoolRecommendation.AddItem ""
    cboHighSchoolRecommendation.AddItem LoadResString(2503)
    cboHighSchoolRecommendation.AddItem LoadResString(2502)
    m_ComboDetails.Add 0, LoadResString(2503), "[iHighSchoolRecommendation]"
    m_ComboDetails.Add 1, LoadResString(2502), "[iHighSchoolRecommendation]"
    
    cboLetterFlag.AddItem ""
    cboLetterFlag.AddItem LoadResString(2504)
    cboLetterFlag.AddItem LoadResString(2505)
    m_ComboDetails.Add 0, LoadResString(2504), "[iLetterFlag]"
    m_ComboDetails.Add 1, LoadResString(2505), "[iLetterFlag]"
    
    'comment by ganesh 21/08/2002 to resolve machine hang problem and changed zipcode combo to text box
   '***************************************
    ' add zip code to combo
'    l_str_sqlZip = "Select iZipCodeId, vZipCodeName from tbSTEZipCodeMaster"
'    Set l_obj_rsZipCode = g_obj_Conn.Execute(l_str_sqlZip)
'
'         cboZipCode.AddItem "", 0
'
'         If Not l_obj_rsZipCode.EOF Then
'            l_obj_rsZipCode.MoveFirst
'         Else
'            Exit Sub
'         End If
'
'         With m_ComboDetails
'            While Not l_obj_rsZipCode.EOF
'               cboZipCode.AddItem l_obj_rsZipCode.Fields("vZipCodeName")
'               m_ComboDetails.Add l_obj_rsZipCode.Fields("iZipCodeId"), l_obj_rsZipCode.Fields("vZipCodeName"), cboZipCode.Tag
'               l_obj_rsZipCode.MoveNext
'            Wend
'         End With
'
'     l_obj_rsZipCode.Close
'     Set l_obj_rsZipCode = Nothing
    '********************************************
    
    m_bload = False
Exit Sub
ErrHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_colFieldDetails = Nothing
    m_bDirty = False
    Call g_void_CloseChildForm
End Sub

Private Sub hfgSearchGrid_DblClick()
    'check  if existing data is dirty if yes then prompt for saving
    ' else
    'populate the new row in the text boxes
    AssignValues True
    Call g_void_HighlightRow(hfgSearchGrid.Row, f_int_PrevRow)
    f_int_PrevRow = hfgSearchGrid.Row
End Sub

Private Sub hfgSearchGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        hfgSearchGrid_DblClick
    End If
End Sub

Private Sub SetChange()
    lblErrorMsg.Caption = ""
    lblErrorMsg.Visible = False
    If m_bMode = "SEARCH" Then Exit Sub
    If m_bChangeOn = False Then m_bDirty = True
    If m_bDirty = True And fMainForm.mnuToolsSave.Enabled = False Then
        fMainForm.mnuToolsSave.Enabled = True
        fMainForm.mnuToolsCancel.Enabled = True
        
        fMainForm.Toolbar1.Buttons("Save").Enabled = True
        fMainForm.Toolbar1.Buttons("Cancel").Enabled = True
    End If
End Sub

Public Function ExtraValidation() As Boolean
    On Error GoTo ErrorHandler
    ' if goes through then ExtraValidation is true if fails then false
    
    ExtraValidation = True
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Private Sub txtAddress1_Change()
    SetChange
End Sub

Private Sub txtAddress2_Change()
    SetChange
End Sub

Private Sub txtFaxNo_Change()
    SetChange
End Sub

Private Sub txtFaxNo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtHighSchoolCode_Change()
    SetChange
End Sub

Private Sub txtHighSchoolDropRecommendation_Change()
    SetChange
End Sub

Private Sub txtHighSchoolDropRecommendation_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtHighSchoolId_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtHighSchoolName_Change()
    SetChange
End Sub

Private Sub txtHighSchoolRecommendationYear1_Change()
    SetChange
End Sub

Private Sub txtHighSchoolRecommendationYear1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtHighSchoolRecommendationYear2_Change()
    SetChange
End Sub

Private Sub txtHighSchoolRecommendationYear2_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRepresentativename_Change()
    SetChange
End Sub

Private Sub txtTelephoneNo_Change()
    SetChange
End Sub

Private Sub txtTelephoneNo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtZipCodeId_Change()
    SetChange
End Sub

Private Sub txtZipCodeId_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
