VERSION 5.00
Begin VB.Form frmChooseiReport 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   10110
   ClientLeft      =   1275
   ClientTop       =   900
   ClientWidth     =   13230
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmChooseiReport.frx":0000
   ScaleHeight     =   10110
   ScaleWidth      =   13230
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdExpl 
      Caption         =   "仮計算"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   39
      Top             =   4920
      Width           =   1695
   End
   Begin VB.ComboBox cboSubject 
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
      Left            =   2820
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   23
      Top             =   1080
      Width           =   2100
   End
   Begin VB.ComboBox cboSubjectId 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "1071"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   21
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtChoseiScore 
      Alignment       =   1  '右揃え
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2880
      MaxLength       =   6
      TabIndex        =   20
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtChoseiScore 
      Alignment       =   1  '右揃え
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   4800
      MaxLength       =   6
      TabIndex        =   19
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtChoseiScore 
      Alignment       =   1  '右揃え
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   18
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtChoseiScore 
      Alignment       =   1  '右揃え
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   2880
      MaxLength       =   6
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtChoseiScore 
      Alignment       =   1  '右揃え
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   4800
      MaxLength       =   6
      TabIndex        =   16
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtChoseiScore 
      Alignment       =   1  '右揃え
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   15
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtChoseiScore 
      Alignment       =   1  '右揃え
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   2880
      MaxLength       =   6
      TabIndex        =   14
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtChoseiScore 
      Alignment       =   1  '右揃え
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   4800
      MaxLength       =   6
      TabIndex        =   13
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtChoseiScore 
      Alignment       =   1  '右揃え
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   12
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox txtAve 
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
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblErrorDetails 
      BackStyle       =   0  '透明
      Caption         =   "Error Details"
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
      Left            =   1200
      TabIndex        =   38
      Top             =   4080
      Visible         =   0   'False
      Width           =   9735
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  '透明
      Caption         =   "1753"
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
      Height          =   375
      Left            =   360
      TabIndex        =   37
      Top             =   1095
      Width           =   2175
   End
   Begin VB.Label lblDay1 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "lblDay1"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   36
      Top             =   2460
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "男性調整点"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2520
      TabIndex        =   35
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "女性調整点"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4440
      TabIndex        =   34
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "日別調整点"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   33
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblDay2 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "lblDay2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   32
      Top             =   3060
      Width           =   1455
   End
   Begin VB.Label lblDay3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "lblDay3"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   3660
      Width           =   1455
   End
   Begin VB.Label lblDay12 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "lblDay1"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   30
      Top             =   4980
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "男性平均点"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "女性平均点"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4440
      TabIndex        =   28
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "日別平均点"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   27
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblDay22 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "lblDay2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   5580
      Width           =   1455
   End
   Begin VB.Label lblDay32 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "lblDay3"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   6180
      Width           =   1455
   End
   Begin VB.Label lblDayT 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "全体"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   24
      Top             =   6780
      Width           =   1455
   End
End
Attribute VB_Name = "frmChooseiReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmChooseiScore
'Author         :   Dileep Cherian
'Created On     :   18/9/01
'Description    :   This form will be used as a mechanism to insert choosei score for first Examination.
'Reference      :   FunctionalSpecs OF CHOSEISCORE.doc(Ver 1.1)
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History   -   04/04/2002  -   Dileep Cherian
'User should be able to resize the coulmns, incase part of data is not visible in the normal display
'On pressing ente after editing a column value, the focus should move to the next row (same column)
'**************************************************************************************************
' Modification History
' Changes by : Mahesh (10-11-2001)
' Changed to allow selection of raw score in a grid fashion instead of combo selection
' Ammendments - NyushiChangesSummary.doc ver 1.0
' Changes by : Mahesh (17-5-2002)
' Chenged to add room checkbox, on clicking of which records will be added to grid for every room in rommMaster
' Changed by : Mahesh (25/5/2002)
' Changed to add list box for rooms. instead of all rooms from room master, only selected rooms data will be added to grod.

'2004/04　全面的に作り変え
'面接用に別出し
'旧システムにあわせた画面仕様に変更（設定可能な情報をすべて表示し、入力なしで調整なしとする）

'tbSTEScoreProfileのiChoseiScoreに設定条件の一致する学生（学生情報tbSTEExamineeなど）を
'条件に登録する。ここでは一律に更新し、結果０以下の場合は帳票などのほうで表示を０とする
'調整点は積算する。更新処理ではiChoseiScore=iChoseiScore+入力値となる
'調整条件とその点数はtbSTEChoseiJokenに履歴登録する。
'入力取り消し処理はなく、取り消し時は+-を逆転した調整点を再度行うことにより実施する

'つまり、素点入力済みでなければこまったことになる。
'入力途中だった場合は手動で調整点をクリアし、tbSTEChoseiJokenも削除する必要がある

'tbSTEScoreDetailはさわらない

'条件は
'日付毎に男女別、区別なしの３つ
'情報として、上記条件別の素点平均点を表示

Option Explicit

' database related variables
Dim m_obj_Rst As New ADODB.Recordset    ' recordset object
Dim m_str_SQL As String                 ' to store the SQL string
Dim m_int_SelectedSubject As Long    ' to store the selected subject from the subject combo
'Dim m_int_NoOfErr As Long            ' to keep track of no of errors
Dim m_int_NoOfConditions As Long     ' to track the no of conditions
Public m_int_ChoseiJoken As Long          ' to diff b/w Grace Score and Suisen Score
Dim m_bln_OnceEntered As Boolean        ' boolean stores whether the conditions have been entered once. if so,user hav to clear off first

Dim prvsSecondExamDay1 As String '２次試験１日目の日付(YYYY/MM/DD)
Dim prvsSecondExamDay2 As String '２次試験２日目の日付(YYYY/MM/DD)
Dim prvsSecondExamDay3 As String '２次試験３日目の日付(YYYY/MM/DD)、２日のときは空文字

Private Sub cboSubject_Click()
    cboSubjectId.ListIndex = cboSubject.ListIndex
End Sub
'
'Private Sub chkRawScore_Click()
'    ' if its already checked and some values are there in the rawscore grid, then clear it
'        ' and then make id disabled
'    ' if its not checked yet, check it and make the grid editable - default value beig 0-100
'    Dim l_int_Counter As Integer        ' counter variable
'    On Error GoTo ErrorHandler
'
'    If chkRawScore.Value = 1 Then
'        ' not checked yet - enable the grid
'        vsfselectRawScore.Editable = flexEDKbdMouse
'    Else
'        ' already checked - clear and make disabled
'        vsfselectRawScore.Editable = flexEDNone
'        With vsfselectRawScore
'            If .Rows > 1 Then
'               For l_int_Counter = .Rows - 1 To 2 Step -1  ' for all rows.. remove them
'                   .RemoveItem l_int_Counter
'               Next
'               .Row = 1
'               .Col = 0
'               .Text = 0
'               .Col = 1
'               .Text = 100
'            End If
'        End With
'    End If
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Sub
'
'Private Sub chkRoom_Click()
'    ' enable/disable the check box for room
'    If chkRoom.Value = 1 Then
'        lstRooms.Enabled = True
'    Else
'        lstRooms.Enabled = False
'    End If
'End Sub
'
'Private Sub cmdClear_Click()
'    ' clear the main grid as well as the raw score grid
'    Dim l_int_Counter As Integer                ' counter variable
'    On Error GoTo ErrorHandler
'
'    ' clear the main grid
'    With vsfSearchGrid
'         For l_int_Counter = .Rows - 1 To 1 Step -1    ' for all rows.. remove them
'            .RemoveItem l_int_Counter
'        Next
'    End With
'
'    ' clear the raw score grid
'    With vsfselectRawScore
'        If vsfselectRawScore.Rows > 1 Then
'           For l_int_Counter = .Rows - 1 To 2 Step -1  ' for all rows.. remove them
'               .RemoveItem l_int_Counter
'           Next
'        End If
'        .Row = 1
'        .Col = 0
'        .Text = 0
'        .Col = 1
'        .Text = 100
'    End With
'    lblErrorDetails.Caption = ""
'    m_bln_OnceEntered = False
'    cmdSubmit.Enabled = False
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Sub
'
'Private Sub cmdOK_Click()
'    ' add ros to the grid and populate it, based on the selected input criteria
'    Dim l_int_Counter As Integer        ' counter
'    Dim l_dbl_RawScoreFrom As Double    ' to store lower limit of raw score
'    Dim l_dbl_RawScoreTo As Double      ' to store upper limit of raw score
'    Dim l_int_ChkDay As Integer         ' day is checked or not
'    Dim l_int_Count As Integer          ' counter
'    Dim l_int_room  As Integer          ' Room is checked or not
'    Dim l_int_RoomId As Integer         'Room Id to be populated in Grid
'    Dim l_Str_RoomDesc As String        'Room Desc to be populated in Grid
'    Dim l_int_RoomCount As Integer
'    Dim l_bln_RoomSelected As Boolean       ' boolean stores whether a room is selected
'
'    On Error GoTo ErrorHandler
'
'    ' ask user to clear off the grid, if some data is already displayed on the grid
'    If m_bln_OnceEntered Then
'         lblErrorDetails.Caption = LoadResString(1772)
'         lblErrorDetails.Visible = True
'        Exit Sub
'    End If
'
'    vsfSearchGrid.Redraw = flexRDNone
'
'    If chkRawScore.Value = 1 Then
'        If f_bln_ValidateRange > 0 Then
'           If f_bln_ValidateRange = 1 Then
'              lblErrorDetails.Caption = LoadResString(1762)
'              lblErrorDetails.Visible = True
'           Else
'              lblErrorDetails.Caption = LoadResString(1771)
'              lblErrorDetails.Visible = True
'           End If
'           Exit Sub
'        End If
'    End If
'
'    m_bln_OnceEntered = True
''    m_int_NoOfErr = 0
'
'    'Instead of combo, loop through the vsfselectRawScore Grid
'    For l_int_Counter = 1 To vsfselectRawScore.Rows - 1  ' for all rows
'         If chkRawScore.Value = 1 Then
'             vsfselectRawScore.Row = l_int_Counter  ' row counter
'             vsfselectRawScore.Col = 0   '0th column
'
'             If vsfselectRawScore.Text = "" Then Exit For
'
'             If IsNull(vsfselectRawScore.Text) Then Exit For   'exit if no value in the row
'                l_dbl_RawScoreFrom = vsfselectRawScore.Text
'
'             vsfselectRawScore.Col = 1   'fist column
'             l_dbl_RawScoreTo = vsfselectRawScore.Text
'        Else
'            l_dbl_RawScoreFrom = 0
'            l_dbl_RawScoreTo = 100
'            If l_int_Counter > 1 Then Exit For
'        End If
'
'        With vsfSearchGrid
'
'        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'            l_int_ChkDay = IIf(chkDay.Value = 1, 1, 0)   'Day is checked?
'            l_int_room = IIf(chkRoom.Value = 1, 1, 0)    'Room is Checked?
'        Else
'            l_int_ChkDay = 0
'            l_int_room = 0
'        End If
'        'loop for all rows of romm master if room checkbox is checked
'        If l_int_room = 1 Then
'            'check whether any room is selected or not in the listbox
'            For l_int_RoomCount = 0 To lstRooms.ListCount - 1
'                If lstRooms.Selected(l_int_RoomCount) = True Then
'                    l_bln_RoomSelected = True
'                End If
'            Next
'            If Not l_bln_RoomSelected Then
'                lblErrorDetails.Caption = LoadResString(2495)   '"Select a room"
'                lblErrorDetails.Visible = True
'                Exit Sub
'            End If
'            For l_int_RoomCount = 0 To lstRooms.ListCount - 1
'                If lstRooms.Selected(l_int_RoomCount) = True Then 'if the current item is selected
'                    l_int_RoomId = lstRooms.ItemData(l_int_RoomCount)
'                    l_Str_RoomDesc = lstRooms.List(l_int_RoomCount)
'                    If chkSex.Value = 1 Then
'                        If l_int_ChkDay = 1 Then
'                            For l_int_Count = 1 To 3  'adds 3 rows
'                                ' sex is checked, so add 2 rows to the grid
'                                .AddItem "", .Rows
'                                .Row = .Rows - 1
'                                Call f_void_PopulateGrid(1, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
'
'                                .AddItem "", .Rows
''                                .Row=.Row + 1
'                                .Row = .Rows - 1
'                                Call f_void_PopulateGrid(2, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
'                            Next
'                        Else
'                            ' sex is checked, so add 2 rows to the grid
'                            .AddItem "", .Rows
'                            .Row = .Rows - 1
'                            Call f_void_PopulateGrid(1, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
'
'                            .AddItem "", .Rows
''                            .Row = .Row + 1
'                            .Row = .Rows - 1
'                            Call f_void_PopulateGrid(2, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
'                        End If
'                    Else
'                        If l_int_ChkDay = 1 Then
'                            For l_int_Count = 1 To 3
'                                ' sex not checked, so add only 1 row to the grid
'                                .AddItem "", .Rows
''                                .Row = .Row + 1
'                                .Row = .Rows - 1
'                                Call f_void_PopulateGrid(0, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
'                            Next
'                        Else
'                            ' sex not checked, so add only 1 row to the grid
'                                .AddItem "", .Rows
''                                .Row = .Row + 1
'                                .Row = .Rows - 1
'                                Call f_void_PopulateGrid(0, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
'                        End If
'                    End If
'                End If 'if the item in list is selected
'
'            Next 'for all items in list box
'        Else     'original case. No room checkbox checked
'            If chkSex.Value = 1 Then
'                If l_int_ChkDay = 1 Then
'                    For l_int_Count = 1 To 3  'adds 3 rows
'                        ' sex is checked, so add 2 rows to the grid
'                        .AddItem "", .Rows
'                        .Row = .Rows - 1
'                        Call f_void_PopulateGrid(1, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
'
'                        .AddItem "", .Rows
''                        .Row = .Row + 1
'                        .Row = .Rows - 1
'                        Call f_void_PopulateGrid(2, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
'                    Next
'                Else
'                    ' sex is checked, so add 2 rows to the grid
'                    .AddItem "", .Rows
'                    .Row = .Rows - 1
'                    Call f_void_PopulateGrid(1, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
'
'                    .AddItem "", .Rows
''                    .Row = .Row + 1
'                    .Row = .Rows - 1
'                    Call f_void_PopulateGrid(2, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
'                End If
'            Else
'                If l_int_ChkDay = 1 Then
'                    For l_int_Count = 1 To 3
'                        ' sex not checked, so add only 1 row to the grid
'                        .AddItem "", .Rows
''                        .Row = .Row + 1
'                        .Row = .Rows - 1
'                        Call f_void_PopulateGrid(0, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
'                    Next
'                Else
'                    ' sex not checked, so add only 1 row to the grid
'                        .AddItem "", .Rows
''                        .Row = .Row + 1
'                        .Row = .Rows - 1
'                        Call f_void_PopulateGrid(0, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
'                End If
'            End If
'        End If   'chkRoom checked
'        End With
'    Next
'    cmdSubmit.Enabled = True
'    vsfSearchGrid.Redraw = flexRDBuffered
'
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Sub

Private Sub cmdExpl_Click()

Dim l_Bln_RecordsUpdated As Boolean
Dim sChoseiScore As String
Dim sSQL As String
Dim sTargetDay As String
Dim sTargetSex As String
Dim lRtn As Long

Dim iLoopCnt As Long

On Error GoTo ErrorHandler

    g_obj_Conn.BeginTrans                   ' all the records in the grid has to be updated or else rollback

    'いったんクリアする
    sSQL = "UPDATE tbSTEScoreProfile"
    sSQL = sSQL & " SET fChoseiScore = 0 "
    sSQL = sSQL & ", dtUpdate = '" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' "
    sSQL = sSQL & " WHERE iSubjectProfileId= " & cboSubjectId.Text & " "
    sSQL = sSQL & " AND exists ( select 1 from tbSTEExamineeProfile "
    sSQL = sSQL & "               where iNendo = " & g_int_CurrentNendo
    sSQL = sSQL & "                 and tbSTEScoreProfile.iExamineeProfileID = tbSTEExamineeProfile.iExamineeProfileID ) "

    Call g_obj_Conn.Execute(sSQL)

    For iLoopCnt = 0 To 8
    '0:初日男性     1:初日女性      2:初日日別
    '3:２日男性     4:２日女性      5:２日日別
    '6:３日男性     7:３日女性      8:３日日別
        sChoseiScore = Trim(txtChoseiScore(iLoopCnt).Text)
        '入力なしは次に
        If sChoseiScore = "" Or sChoseiScore = "0" Then GoTo LoopEnd
        '数値認識できない場合はエラー
        If Not gf_DblCheck(sChoseiScore) Then
            GoTo ErrorHandler 'ロールバックは飛び先で実施
        End If
        '範囲(-100～100)外はエラー
        If CDbl(sChoseiScore) < -100 Or CDbl(sChoseiScore) > 100 Then
            GoTo ErrorHandler 'ロールバックは飛び先で実施
        End If
        sSQL = "UPDATE tbSTEScoreProfile"
        sSQL = sSQL & " SET fChoseiScore=isnull(fChoseiScore , 0 ) + " & sChoseiScore & " "
        sSQL = sSQL & ", dtUpdate='" & Format(Date, "YYYY/MM/DD HH:MM:SS") & "' "
        sSQL = sSQL & " WHERE iSubjectProfileId= " & cboSubjectId.Text & " "
        sSQL = sSQL & " AND exists ( select 1 from tbSTEExamineeProfile "
        '日別
        sSQL = sSQL & "               where dtSecondExamDay = ( select "
        Select Case iLoopCnt
        Case 0, 1, 2
            sTargetDay = Left(prvsSecondExamDay1, 4) & Mid(prvsSecondExamDay1, 6, 2) & Mid(prvsSecondExamDay1, 9, 2)
            sSQL = sSQL & " dtSecondExamDay1 "
        Case 3, 4, 5
            sTargetDay = Left(prvsSecondExamDay2, 4) & Mid(prvsSecondExamDay2, 6, 2) & Mid(prvsSecondExamDay2, 9, 2)
            sSQL = sSQL & " dtSecondExamDay2 "
        Case 6, 7, 8
            sTargetDay = Left(prvsSecondExamDay3, 4) & Mid(prvsSecondExamDay3, 6, 2) & Mid(prvsSecondExamDay3, 9, 2)
            sSQL = sSQL & " dtSecondExamDay3 "
        End Select
        sSQL = sSQL & "                                         from tbSTESecondExamProfile where iSystemProfileId = ( select top 1 iSystemProfileId from tbSTEsystemProfile where iActiveFlag = 1 ) ) "
        '男女別
        sTargetSex = "-1"
        Select Case iLoopCnt
        Case 0, 3, 6
            sTargetSex = "0"
            sSQL = sSQL & "                 and iSex = 0 "
        Case 1, 4, 7
            sTargetSex = "1"
            sSQL = sSQL & "                 and iSex = 1 "
        End Select
        sSQL = sSQL & "                 and tbSTEScoreProfile.iExamineeProfileID = tbSTEExamineeProfile.iExamineeProfileID ) "

        Call g_obj_Conn.Execute(sSQL)

        lRtn = gflInsChoseiJoken(g_int_CurrentNendo, CInt(cboSubjectId.Text), 2, CDbl(sTargetSex), -1, sTargetDay, -1, sChoseiScore)

        l_Bln_RecordsUpdated = True

LoopEnd:

    Next iLoopCnt

    '調整点付加後の平均点表示
    Call lsGetAve

    g_obj_Conn.RollbackTrans

Exit Sub

ErrorHandler:

'トランザクションを起こす前、閉じた後のエラー対処のため
On Error GoTo ErrorHandler2

    g_obj_Conn.RollbackTrans
    
On Error GoTo 0
ErrorHandler2:

    lblErrorDetails.Visible = True
    lblErrorDetails.Caption = LoadResString(1761) & _
        vbCrLf & LoadResString(1125)

End Sub

Private Sub cmdSubmit_Click()

Dim l_Bln_RecordsUpdated As Boolean
Dim sChoseiScore As String
Dim sSQL As String
Dim sTargetDay As String
Dim sTargetSex As String
Dim lRtn As Long

Dim iLoopCnt As Long

On Error GoTo ErrorHandler

    g_obj_Conn.BeginTrans                   ' all the records in the grid has to be updated or else rollback

    lRtn = gflDelChoseiJoken(g_int_CurrentNendo, CInt(cboSubjectId.Text), 2)

    'いったんクリアする
    sSQL = "UPDATE tbSTEScoreProfile"
    sSQL = sSQL & " SET fChoseiScore = 0 "
    sSQL = sSQL & ", dtUpdate = '" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' "
    sSQL = sSQL & " WHERE iSubjectProfileId= " & cboSubjectId.Text & " "
    sSQL = sSQL & " AND exists ( select 1 from tbSTEExamineeProfile "
    sSQL = sSQL & "               where iNendo = " & g_int_CurrentNendo
    sSQL = sSQL & "                 and tbSTEScoreProfile.iExamineeProfileID = tbSTEExamineeProfile.iExamineeProfileID ) "

    Call g_obj_Conn.Execute(sSQL)

    For iLoopCnt = 0 To 8
    '0:初日男性     1:初日女性      2:初日日別
    '3:２日男性     4:２日女性      5:２日日別
    '6:３日男性     7:３日女性      8:３日日別
        sChoseiScore = Trim(txtChoseiScore(iLoopCnt).Text)
        '入力なしは次に
        If sChoseiScore = "" Or sChoseiScore = "0" Then GoTo LoopEnd
        '数値認識できない場合はエラー
        If Not gf_DblCheck(sChoseiScore) Then GoTo ErrorHandler 'ロールバックは飛び先で実施
        '入力なしは次に
        If CDbl(sChoseiScore) = 0 Then GoTo LoopEnd
        '範囲(-100～100)外はエラー
        'update,2014/12
        'If CDbl(sChoseiScore) < -100 Or CDbl(sChoseiScore) > 100 Then GoTo ErrorHandler 'ロールバックは飛び先で実施
        If CDbl(sChoseiScore) < -100 Or CDbl(sChoseiScore) > 1000 Then GoTo ErrorHandler 'ロールバックは飛び先で実施
        sSQL = "UPDATE tbSTEScoreProfile"
        sSQL = sSQL & " SET fChoseiScore=fChoseiScore +  " & sChoseiScore & " "
        sSQL = sSQL & ", dtUpdate='" & Format(Date, "YYYY/MM/DD HH:MM:SS") & "' "
        sSQL = sSQL & " WHERE iSubjectProfileId= " & cboSubjectId.Text & " "
        sSQL = sSQL & " AND exists ( select 1 from tbSTEExamineeProfile "
        '日別
        sSQL = sSQL & "               where dtSecondExamDay = ( select "
        Select Case iLoopCnt
        Case 0, 1, 2
            'update,xzg,2009/12/21,S----------
            'sTargetDay = Left(prvsSecondExamDay1, 4) & Mid(prvsSecondExamDay1, 6, 2) & Mid(prvsSecondExamDay1, 9, 2)
            sTargetDay = prvsSecondExamDay1
            'update,xzg,2009/12/21,E----------
            sSQL = sSQL & " dtSecondExamDay1 "
        Case 3, 4, 5
            'update,xzg,2009/12/21,S----------
            'sTargetDay = Left(prvsSecondExamDay2, 4) & Mid(prvsSecondExamDay2, 6, 2) & Mid(prvsSecondExamDay2, 9, 2)
            'update,xzg,2009/12/21,E----------
            sTargetDay = prvsSecondExamDay2
            sSQL = sSQL & " dtSecondExamDay2 "
        Case 6, 7, 8
            'update,xzg,2009/12/21,S----------
            'sTargetDay = Left(prvsSecondExamDay3, 4) & Mid(prvsSecondExamDay3, 6, 2) & Mid(prvsSecondExamDay3, 9, 2)
            sTargetDay = prvsSecondExamDay3
            'update,xzg,2009/12/21,E----------
            sSQL = sSQL & " dtSecondExamDay3 "
        End Select
        sSQL = sSQL & "                                         from tbSTESecondExamProfile where iSystemProfileId = ( select top 1 iSystemProfileId from tbSTEsystemProfile where iActiveFlag = 1 ) ) "
        '男女別
        sTargetSex = "-1"
        Select Case iLoopCnt
        Case 0, 3, 6
            sTargetSex = "0"
            sSQL = sSQL & "                 and iSex = 0 "
        Case 1, 4, 7
            sTargetSex = "1"
            sSQL = sSQL & "                 and iSex = 1 "
        End Select
        sSQL = sSQL & "                 and tbSTEScoreProfile.iExamineeProfileID = tbSTEExamineeProfile.iExamineeProfileID ) "

        Call g_obj_Conn.Execute(sSQL)

        lRtn = gflInsChoseiJoken(g_int_CurrentNendo, CInt(cboSubjectId.Text), 2, CDbl(sTargetSex), -1, sTargetDay, -1, sChoseiScore)

        l_Bln_RecordsUpdated = True

LoopEnd:

    Next iLoopCnt

    g_obj_Conn.CommitTrans

'    '表示クリア
'    For iLoopCnt = 0 To 8
'        Me.txtChoseiScore(iLoopCnt).Text = ""
'    Next

    '調整点付加後の平均点表示
    Call lsGetAve

    If l_Bln_RecordsUpdated Then
        lblErrorDetails.Caption = LoadResString(2404)
    Else
        lblErrorDetails.Caption = LoadResString(2427)
    End If
    lblErrorDetails.Visible = True

Exit Sub

ErrorHandler:

'トランザクションを起こす前、閉じた後のエラー対処のため
On Error GoTo ErrorHandler2
    
    g_obj_Conn.RollbackTrans
    
On Error GoTo 0
ErrorHandler2:

    lblErrorDetails.Visible = True
    lblErrorDetails.Caption = LoadResString(1761) & _
        vbCrLf & LoadResString(1125)
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

Dim iLoopCnt As Long

On Error GoTo ErrorHandler

    LoadResStrings Me
    If m_int_ChoseiJoken = 1 Then
        Me.Caption = LoadResString(1012)
    Else
        Me.Caption = LoadResString(1751)
    End If
    Call g_void_SetFontProperties(Me)     ' set the font properties
    m_int_NoOfConditions = 0    ' initialise the no of conditions
    ' select all subjects that come under the selected exam type
    m_str_SQL = "SELECT iSubjectProfileId,vSubjectName FROM tbSTESubjectProfile"

'    ' changed on 14/05/02 to incorporate choosei for Hyotei also
'    If m_int_ChoseiJoken = 1 Then
'        m_str_SQl = m_str_SQl & " WHERE iExamType = 0"
'    ElseIf g_int_ExamType = 1 Then
'        m_str_SQl = m_str_SQl & " WHERE iExamType = " & g_int_ExamType
'    ElseIf g_int_ExamType = 2 Or g_int_ExamType = 3 Or g_int_ExamType = 4 Or g_int_ExamType = 5 Then
'        m_str_SQl = m_str_SQl & " WHERE iExamType = 2 or iExamType = 3 or iExamType = 4 or iExamType = 5"
'    End If
    m_str_SQL = m_str_SQL & " WHERE iSubType = 4"
    
    m_str_SQL = m_str_SQL & " ORDER BY vSubjectName"
    Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQL)

    If Not m_obj_Rst.EOF Then
        m_int_SelectedSubject = m_obj_Rst("iSubjectProfileId")
        ' add the subjects to combo box
        Do While Not m_obj_Rst.EOF
            cboSubject.AddItem m_obj_Rst("vSubjectName")
            cboSubjectId.AddItem m_obj_Rst("iSubjectProfileId")
            m_obj_Rst.MoveNext
        Loop
        cboSubject.ListIndex = 0
    End If
    
    ' release the object variables
    m_obj_Rst.Close
    Set m_obj_Rst = Nothing

'試験日の取得
    m_str_SQL = "Select"
    m_str_SQL = m_str_SQL & "  convert( varchar , dtSecondExamDay1 , 111 ) "
    m_str_SQL = m_str_SQL & " ,convert( varchar , dtSecondExamDay2 , 111 ) "
    m_str_SQL = m_str_SQL & " ,isnull( convert( varchar , dtSecondExamDay3 , 111 ) , '' ) "
    m_str_SQL = m_str_SQL & " From tbSTESecondExamProfile as se "
    m_str_SQL = m_str_SQL & " Where iSystemProfileID = ( select Top 1 iSystemProfileID From tbSteSystemProfile where iActiveFlag = 1 ) "

    Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQL)

    If Not m_obj_Rst.EOF Then
        prvsSecondExamDay1 = Trim(m_obj_Rst.Fields(0))
        prvsSecondExamDay2 = Trim(m_obj_Rst.Fields(1))
        prvsSecondExamDay3 = Trim(m_obj_Rst.Fields(2))
    Else
        prvsSecondExamDay1 = ""
        prvsSecondExamDay2 = ""
        prvsSecondExamDay3 = ""
    End If

    lblDay1.Caption = prvsSecondExamDay1
    lblDay2.Caption = prvsSecondExamDay2
    lblDay3.Caption = prvsSecondExamDay3
    lblDay12.Caption = prvsSecondExamDay1
    lblDay22.Caption = prvsSecondExamDay2
    lblDay32.Caption = prvsSecondExamDay3

    If prvsSecondExamDay3 = "" Then
        For iLoopCnt = 6 To 8
            lblDay3.Visible = False
            Me.txtChoseiScore(iLoopCnt).Visible = False
            lblDay32.Visible = False
            Me.txtAve(iLoopCnt).Visible = False
        Next
    End If

    ' release the object variables
    m_obj_Rst.Close
    Set m_obj_Rst = Nothing

    Call lsGetAve

    Call f_void_ReadAlsoData

'    cmdSubmit.Enabled = False

    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub
'
'Private Sub f_void_InitGrid()
'     vsfSearchGrid.Redraw = flexRDNone
'
'    With vsfSearchGrid
'        .Visible = False
'        .BackColor = &HFFFFFF
'        .BackColorBkg = &HFFFFFF
'        .BackColorFixed = &H8000000F
'        .BackColorSel = &H800000
'        .FixedCols = 0
'        .TextStyleFixed = flexTextFlat
'        .ForeColorFixed = &H80000008
'        .ForeColor = &H800000
''        .CellTextStyle = "0"
'        .GridLines = flexGridFlat
'        .GridLinesFixed = flexGridInset
'        .GridColor = &H808080
'        .AllowUserResizing = flexResizeColumns
'        .Visible = True
'        .Rows = 1
'
'        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'            ' for second exam one additional column is required for the day combo
'            ' If room checkbox is checked, 2 columns for Room id and name
'            If chkRoom.Value = 1 Then
'                .Cols = 10
'            Else
'                .Cols = 8
'            End If
'        Else
'            ' for ist exam, day column is not there, hence one column less
'            .Cols = 7
'        End If
'
'        .Row = 0
'        .Col = 0
'        .ColWidth(0) = 700
'        .Text = LoadResString(1756)   'Sr no  0
'        .CellAlignment = flexAlignRightBottom
'
'        .Col = .Col + 1
'        .ColWidth(1) = 2200
'        .Text = LoadResString(1757)    'subject  1
'
'        .Col = .Col + 1
'        .ColWidth(2) = 2000
'        .Text = LoadResString(1758)  'Raw score from  2
'        .CellAlignment = flexAlignRightBottom
'
'        .Col = .Col + 1
'        .ColWidth(3) = 2000
'        .Text = LoadResString(1759)   'raw score to  3
'        .CellAlignment = flexAlignRightBottom
'
'        .Col = .Col + 1
'        .ColWidth(4) = 1200
'        .Text = LoadResString(1754)   'Sex  4
'
'        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'            ' add the additional column for the day
'            .Col = .Col + 1
'            .ColWidth(7) = 1600
'            .Text = LoadResString(1755)  'Day   Col is 5
'            'new col for roomID
'            If chkRoom.Value = 1 Then
'                .Col = .Col + 1
'                .ColWidth(6) = 0 'hidden column 6 for room Id
'                .Col = .Col + 1
'                .ColWidth(7) = 2000   'Column 7 for Room Desc
'                .Text = LoadResString(2002)
'            End If
'        End If
'
'        .Col = .Col + 1
'        .ColWidth(.Col) = 1000  '5
'        .Text = LoadResString(1760)  'Col 8 Average
'        .CellAlignment = flexAlignRightBottom
'
'        .Col = .Col + 1
'        .ColWidth(.Col) = 1700   '6
'        .Text = LoadResString(1751)  'col 9 choosei score (last column)
'        .CellAlignment = flexAlignRightBottom
'    End With
'        vsfSearchGrid.Redraw = flexRDBuffered
'
'    Exit Sub
'End Sub
'
'Private Sub f_void_InitRawScoreGrid()
'
'    With vsfselectRawScore
'        .Visible = False
'        .BackColor = &HFFFFFF
'        .BackColorBkg = &HFFFFFF
'        .BackColorFixed = &H8000000F
'        .BackColorSel = &H800000
'        .FixedCols = 0
'        .TextStyleFixed = flexTextFlat
'
'        ' change made in com design, arka , 11 apr02
'        '.Font.Name = "ＭＳ Ｐゴシック"
'        '.Font.Name = "Verdana"
'
'        .ForeColorFixed = &H80000008
'        .ForeColor = &H800000
'        '.CellTextStyle = "0"
'        .GridLines = flexGridFlat
'        .GridLinesFixed = flexGridInset
'        .GridColor = &H808080
'        .Visible = True
'
'        .Row = 0
'        .Col = 0
'        .ColWidth(0) = 1200
'        .Text = LoadResString(1769)
'        .CellAlignment = flexAlignRightBottom
'
'        .Col = .Col + 1
'        .ColWidth(1) = 1200
'        .Text = LoadResString(1770)
'        .Editable = flexEDNone
'        .Row = .Row + 1
'        .Col = 0
'        .Text = 0
'        .Col = .Col + 1
'        .Text = 100
'    End With
'    Exit Sub
'End Sub
'
'Private Sub f_void_PopulateGrid(ByVal l_bln_SexFlag As Integer, ByVal l_bln_DayFlag As Integer, ByVal l_dbl_RawScoreFrom As Double, ByVal l_dbl_RawScoreTo As Double, Optional ByVal l_int_RoomNo As Integer, Optional ByVal l_str_RoomName As String)
'    Dim l_dbl_Avg As Double         ' to store the average value calculated
'    On Error GoTo ErrorHandler
'    vsfSearchGrid.Redraw = flexRDNone
'
'    With vsfSearchGrid
'        .Col = 0
'        .Text = .Rows - 1
'
'        .Col = .Col + 1
'        .Text = cboSubject.Text
'
'        .Col = .Col + 1
'        .Text = l_dbl_RawScoreFrom
'
'        .Col = .Col + 1
'        .Text = l_dbl_RawScoreTo
'
'        .Col = .Col + 1
'        If l_bln_SexFlag = 1 Then
'            .Text = LoadResString(1837)
'        ElseIf l_bln_SexFlag = 2 Then
'            .Text = LoadResString(1838)
'        Else
'            .Text = LoadResString(1846)
'        End If
'
'        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'            .Col = .Col + 1
'            Select Case l_bln_DayFlag
'            Case 0
'                .Text = LoadResString(1764)
'            Case 1
'                .Text = LoadResString(1765)
'            Case 2
'                .Text = LoadResString(1766)
'            Case 3
'                .Text = LoadResString(1767)
'            End Select
'            If Not IsEmpty(l_int_RoomNo) Then
'                .Col = .Col + 1
'                .Text = l_int_RoomNo
'                .Col = .Col + 1
'                .Text = l_str_RoomName
'            End If
'        End If
'
'        l_dbl_Avg = f_void_GetAverage(l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
'        .Col = .Cols - 2
'        .Text = l_dbl_Avg
'
'        .Col = .Cols - 1
'        .CellBackColor = &HC0C0FF
'        .Text = 0
'    End With
'    vsfSearchGrid.Redraw = flexRDBuffered
'
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_bln_OnceEntered = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call g_void_CloseChildForm
End Sub
'
'Private Sub vsfSearchGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'    ' this code is written to round off the decimal values to 2 digits precision
'    Dim l_int_ChooseiCol As Integer
'    If chkRoom.Value = 1 Then
'        l_int_ChooseiCol = 9
'    Else
'        l_int_ChooseiCol = 8
'    End If
'    With vsfSearchGrid
'        If Trim(.TextMatrix(Row, Col)) <> "" And .Col = l_int_ChooseiCol Then
'            .TextMatrix(Row, Col) = Round(.TextMatrix(Row, Col), 2)
'        End If
'    End With
'End Sub
'
'Private Sub vsfSearchGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    With vsfSearchGrid
'        If .Redraw <> flexRDNone And Col <> vsfSearchGrid.Cols - 1 Then
'            Cancel = True
'            Exit Sub
'        Else
'            vsfSearchGrid.Editable = flexEDKbdMouse
'        End If
'    End With
'End Sub
'
'Private Sub vsfSearchGrid_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
'    With vsfSearchGrid
'        If .Redraw <> flexRDNone And NewCol <> .Cols - 1 Then
'            Cancel = True
'            .Select NewRow, .Cols - 1
'        End If
'    End With
'End Sub
'
'Private Sub vsfSearchGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'    If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'        ' in second exam, only the 8th column is editable (choosei score)
'        If Col <> IIf(chkRoom.Value = 1, 9, 7) Then
'            KeyAscii = 0
'        ElseIf KeyAscii = 13 Then
'            If vsfSearchGrid.Row < vsfSearchGrid.Rows - 1 Then
'                vsfSearchGrid.Row = vsfSearchGrid.Row + 1
'                vsfSearchGrid.Col = Col
'            End If
'        ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
'            KeyAscii = 0
'        'This is to restrict user from entering more than one "." in the value
'        ElseIf KeyAscii = 46 And InStr(1, vsfSearchGrid.EditText, ".") > 0 Then
'            KeyAscii = 0
'        End If
'    ElseIf g_int_ExamType = 1 Then
'        ' in first exam, only the 8th column is editable (choosei score)
'        If Col <> 6 Then
'            KeyAscii = 0
'        ElseIf KeyAscii = 13 Then
'            If vsfSearchGrid.Row < vsfSearchGrid.Rows - 1 Then
'                vsfSearchGrid.Row = vsfSearchGrid.Row + 1
'                vsfSearchGrid.Col = Col
'            End If
'        ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
'            KeyAscii = 0
'        'This is to restrict user from entering more than one "." in the value
'        ElseIf KeyAscii = 46 And InStr(1, vsfSearchGrid.EditText, ".") > 0 Then
'            KeyAscii = 0
'        End If
'    End If
'End Sub

 Private Function f_void_GetAverage(ByVal l_dbl_RawScoreFrom As Double, ByVal l_dbl_RawScoreTo As Double) As Double
    On Error GoTo ErrorHandler
    
    m_str_SQL = "SELECT Avg(fRawScore) from tbSTEScoreProfile where fRawScore BETWEEN "
    m_str_SQL = m_str_SQL & l_dbl_RawScoreFrom & " AND " & l_dbl_RawScoreTo
    m_str_SQL = m_str_SQL & " AND iSubjectProfileId=" & cboSubjectId.Text
    
    m_obj_Rst.Open m_str_SQL, g_obj_Conn, adOpenStatic, adLockReadOnly
    
    If Trim(m_obj_Rst(0)) <> "" Then
        f_void_GetAverage = Round(m_obj_Rst(0), 2)
    Else
        f_void_GetAverage = 0
    End If
    
    m_obj_Rst.Close
    Set m_obj_Rst = Nothing
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
    f_void_GetAverage = 0
End Function

Private Sub vsfSearchGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Clipboard.Clear
    End If
End Sub
'
'Private Sub vsfselectRawScore_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'    On Error GoTo ErrorHandler
'    With vsfselectRawScore
'            If Col < .Cols - 1 Then
'                .Col = .Col + 1
'            ElseIf Col = .Cols - 1 Then
'                If Trim(.Text) <> "" Then
'                    .Col = 0
'                    If Trim(.Text) <> "" Then
'                        .Row = .Rows - 1    'Go to last row and if its not blank, add a row
'                        .Col = 0
'                        If .Text <> "" Then
'                            If .Rows < 11 Then
'                                .Rows = .Rows + 1
'                                .Row = .Rows - 1
'                                .Col = 0
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'    End With
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Sub
'
'Private Sub vsfselectRawScore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = vbRightButton Then
'        Clipboard.Clear
'    End If
'End Sub
'
'Private Sub vsfselectRawScore_Click()
'    lblErrorDetails.Caption = ""
'    lblErrorDetails.Visible = False
'    If chkRawScore.Value = 1 Then
'        vsfselectRawScore.EditCell
'    End If
'End Sub
'
'Private Sub vsfselectRawScore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'    Dim l_int_PrevCol As Integer
'
'    vsfselectRawScore.Redraw = flexRDDirect
'    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> vbKeyReturn Then
'       KeyAscii = 0
'    End If
'End Sub
'
'Private Function f_bln_ValidateRange() As Integer
'
'    Dim l_int_Rows As Integer 'total rows in grid
'    Dim l_int_Counter As Integer ' current row
'    Dim l_bln_RetVal As Integer  ' return value
'    Dim l_int_PrevColVal As Integer  'previous col value of same row
'    'Dim l_int_PrevRowVal As Integer  ' previous col value of prev row
'    ' 0 means all ok
'    ' 1 means check box checked but no values entered
'    ' 2 means Continuity is missing
'    On Error GoTo ErrorHandler
'    l_bln_RetVal = 0
'
'    l_int_Rows = vsfselectRawScore.Rows
'    vsfselectRawScore.Row = 1
'    vsfselectRawScore.Col = 0
'
'    With vsfselectRawScore
'        If .Text = "" Then
'            l_bln_RetVal = 1
'            f_bln_ValidateRange = l_bln_RetVal
'            Exit Function
'        End If
'        l_int_PrevColVal = vsfselectRawScore.Text
'         For l_int_Counter = 1 To .Rows - 1
'             .Row = l_int_Counter
'             .Col = 0
'             If .Text = "" Then Exit For
'             If .Text <> l_int_PrevColVal + 1 And l_int_Counter > 1 Then l_bln_RetVal = 2
'             If .Text <= l_int_PrevColVal And l_int_Counter > 1 Then l_bln_RetVal = 2
'             l_int_PrevColVal = .Text
'             .Col = 1
'
'             l_int_PrevColVal = .Text
'         Next
'    End With
'    f_bln_ValidateRange = l_bln_RetVal
'    Exit Function
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Function

Private Sub lsGetAve()

Dim sSQL As String
Dim sSQL2 As String
Dim sSQL3 As String
Dim iLoopCnt As Long
Dim oRs As ADODB.Recordset

On Error GoTo ErrorHandler

    sSQL = "SELECT "
    sSQL = sSQL & " isnull( convert( varchar , AVG ( isnull( fRawScore , 0 ) + isnull( fChoseiScore , 0 ) ) ) , '' ) "
    sSQL = sSQL & " FROM tbSTEScoreProfile as sp "
    sSQL = sSQL & " WHERE iSubjectProfileID = " & cboSubjectId.Text
    sSQL = sSQL & " AND exists ( SELECT 1 FROM tbSTEexamineeProfile as ep "
    sSQL = sSQL & " WHERE ep.iExamineeProfileID = sp.iExamineeProfileID "
    sSQL = sSQL & " AND iNendo = ( select top 1 iNendo from tbSTEsystemProfile where iActiveFlag = 1 ) "
    sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    sSQL = sSQL & " AND iAbsentFlag = 0 "

    For iLoopCnt = 0 To 11

        Select Case iLoopCnt
        Case 0, 1, 2
            sSQL2 = " AND dtSecondExamDay = ( select dtSecondExamday1 FROM tbSTESecondExamProfile "
            sSQL2 = sSQL2 & "                  where iSystemProfileID = ( select top 1 iSystemProfileID from tbSTEsystemprofile where iActiveFlag = 1 ) ) "
        Case 3, 4, 5
            sSQL2 = " AND dtSecondExamDay = ( select dtSecondExamday2 FROM tbSTESecondExamProfile "
            sSQL2 = sSQL2 & "                  where iSystemProfileID = ( select top 1 iSystemProfileID from tbSTEsystemprofile where iActiveFlag = 1 ) ) "
        Case 6, 7, 8
            sSQL2 = " AND dtSecondExamDay = ( select dtSecondExamday3 FROM tbSTESecondExamProfile "
            sSQL2 = sSQL2 & "                  where iSystemProfileID = ( select top 1 iSystemProfileID from tbSTEsystemprofile where iActiveFlag = 1 ) ) "
        Case 9, 10, 11
            sSQL2 = ""
        End Select

        Select Case iLoopCnt
        Case 0, 3, 6, 9
            sSQL3 = " AND iSex = 0 ) "
        Case 1, 4, 7, 10
            sSQL3 = " AND iSex = 1 ) "
        Case Else
            sSQL3 = " ) "
        End Select

        Set oRs = g_obj_Conn.Execute(sSQL & sSQL2 & sSQL3)

        If Not oRs.EOF Then
            txtAve(iLoopCnt).Text = Trim(oRs.Fields(0))
        End If

        oRs.Close
        Set oRs = Nothing

    Next iLoopCnt

ErrorHandler:

End Sub

Private Sub txtChoseiScore_KeyPress(Index As Integer, KeyAscii As Integer)
'-100 , -XX.X , XX.X , 100 といった入力ができる。なので、MaxLenは5
    Call NumericPeriodMinus(Me, KeyAscii)
End Sub

Private Sub txtChoseiScore_LostFocus(Index As Integer)

    If txtChoseiScore(Index).Text <> "" Then
        If gf_DblCheck(txtChoseiScore(Index).Text) Then
            txtChoseiScore(Index).Text = Format(CDbl(txtChoseiScore(Index).Text), "##0.0")
        Else
            txtChoseiScore(Index).Text = ""
        End If
    End If

End Sub

Private Sub f_void_ReadAlsoData()

Dim sSQL As String
Dim oRs As ADODB.Recordset

On Error GoTo ErrProc

    m_bln_OnceEntered = False

    sSQL = "SELECT "
    sSQL = sSQL & "  fChoseiStartScore as iSex "
    sSQL = sSQL & ", case dtTaishoBi when sep.dtSecondExamDay1 then 1 "
    sSQL = sSQL & "                  when sep.dtSecondExamDay2 then 2 "
    sSQL = sSQL & "                  when sep.dtSecondExamDay3 then 3 "
    sSQL = sSQL & "                                            else 4 end as dtTaishoBi "
    sSQL = sSQL & ", isnull( STR( fChoseiScore , 5 , 1 ) , '' ) as fChoseiScore "
    sSQL = sSQL & " FROM  tbSTEChoseiJoken as cj "
    sSQL = sSQL & " , tbSTESecondExamProfile as sep "
    sSQL = sSQL & " WHERE iSubjectProfileID = " & cboSubjectId.Text
    sSQL = sSQL & " AND   cj.iNendo = " & g_int_CurrentNendo
    sSQL = sSQL & " AND   iChoseiJokenType = 2 " 'iSexやDateが使用されている
    sSQL = sSQL & " AND   sep.iSystemProfileId = ( select max(iSystemProfileId) from tbSTESystemProfile where iActiveFlag = 1 ) "
    sSQL = sSQL & " ORDER BY dtTaishoBi , fChoseiStartScore "

    Set oRs = g_obj_Conn.Execute(sSQL)

    If oRs.EOF Then
        Set oRs = Nothing
        Exit Sub
    End If

    Do Until oRs.EOF

        Select Case oRs.Fields(1)
        Case 1 '初日
            Select Case oRs.Fields(0)
            Case 0 '男
                txtChoseiScore(0).Text = oRs.Fields(2)
            Case 1 '女
                txtChoseiScore(1).Text = oRs.Fields(2)
            'update,xzg,2009/12/21,S----------
            'Case 2 '区別無し
            Case -1 '区別無し
            'update,xzg,2009/12/21,E----------
                txtChoseiScore(2).Text = oRs.Fields(2)
            End Select
        Case 2 '初日
            Select Case oRs.Fields(0)
            Case 0 '男
                txtChoseiScore(3).Text = oRs.Fields(2)
            Case 1 '女
                txtChoseiScore(4).Text = oRs.Fields(2)
                'update,xzg,2009/12/21,S----------
            'Case 2 '区別無し
            Case -1 '区別無し
            'update,xzg,2009/12/21,E----------
                txtChoseiScore(5).Text = oRs.Fields(2)
            End Select
        Case 3 '初日
            Select Case oRs.Fields(0)
            Case 0 '男
                txtChoseiScore(6).Text = oRs.Fields(2)
            Case 1 '女
                txtChoseiScore(7).Text = oRs.Fields(2)
                'update,xzg,2009/12/21,S----------
            'Case 2 '区別無し
            Case -1 '区別無し
            'update,xzg,2009/12/21,E----------
                txtChoseiScore(8).Text = oRs.Fields(2)
            End Select
        End Select

        oRs.MoveNext

    Loop

    oRs.Close
    Set oRs = Nothing

Exit Sub
ErrProc:

End Sub
