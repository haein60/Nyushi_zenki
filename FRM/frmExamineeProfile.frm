VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExamineeProfile 
   AutoRedraw      =   -1  'True
   ClientHeight    =   9750
   ClientLeft      =   1050
   ClientTop       =   435
   ClientWidth     =   13515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmExamineeProfile.frx":0000
   ScaleHeight     =   9750
   ScaleWidth      =   13515
   Tag             =   "1004"
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.ComboBox cboAdmissionType2 
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
      Height          =   360
      Left            =   9735
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   36
      Tag             =   "[iAdmissionType2]"
      Top             =   7455
      Width           =   1935
   End
   Begin VB.ComboBox cboAdmissionType1 
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
      Height          =   360
      Left            =   9705
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   35
      Tag             =   "[iAdmissionType1]"
      Top             =   7065
      Width           =   1935
   End
   Begin VB.TextBox txtExamineeName 
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
      Height          =   360
      Left            =   3390
      TabIndex        =   3
      Tag             =   "[vExamineeName]"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txtKanaName 
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
      Height          =   360
      Left            =   3360
      TabIndex        =   4
      Tag             =   "[vKanaName]"
      Top             =   2670
      Width           =   2775
   End
   Begin VB.TextBox txtAddress 
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
      Height          =   360
      Left            =   3360
      TabIndex        =   5
      Tag             =   "[vAddress]"
      Top             =   3045
      Width           =   2775
   End
   Begin VB.TextBox txtTelephoneNo 
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
      Height          =   360
      Left            =   3360
      TabIndex        =   7
      Tag             =   "[vTelephone]"
      Top             =   4020
      Width           =   1935
   End
   Begin VB.TextBox txtvEmailAddress 
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
      Height          =   360
      Left            =   3360
      TabIndex        =   8
      Tag             =   "[vEmailAddress]"
      Top             =   4395
      Width           =   1935
   End
   Begin VB.TextBox txtNationality 
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
      Height          =   360
      Left            =   3360
      TabIndex        =   9
      Tag             =   "[vNationality]"
      Top             =   4770
      Width           =   1935
   End
   Begin VB.ComboBox cbobackGroundId 
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
      Height          =   360
      Left            =   3360
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   10
      Tag             =   "[iBackgroundId]"
      Top             =   5160
      Width           =   1935
   End
   Begin VB.ComboBox cboQualificationID 
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
      Height          =   360
      Left            =   3360
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   11
      Tag             =   "[iQualificationId]"
      Top             =   5550
      Width           =   1935
   End
   Begin VB.ComboBox cboScienceSubjprofileID2 
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
      Height          =   360
      Left            =   3360
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   14
      Tag             =   "[iScienceSubjProfileId2]"
      Top             =   6690
      Width           =   1935
   End
   Begin VB.ComboBox cboScienceSubjprofileID1 
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
      Height          =   360
      Left            =   3360
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   13
      Tag             =   "[iScienceSubjProfileId1]"
      Top             =   6315
      Width           =   1935
   End
   Begin VB.ComboBox cboLanguageSubjProfileID 
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
      Height          =   360
      Left            =   3360
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   12
      Tag             =   "[iLanguageSubjProfileId]"
      Top             =   5940
      Width           =   1935
   End
   Begin VB.TextBox txtZipAddress 
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
      Height          =   570
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'êÇíº
      TabIndex        =   6
      Tag             =   "[vAddress]"
      Top             =   3435
      Width           =   2775
   End
   Begin VB.TextBox txtExamineeProfileId 
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
      Height          =   360
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   0
      Tag             =   "[iExamineeProfileId]"
      Top             =   1020
      Width           =   1455
   End
   Begin VB.TextBox txtJukenNumber 
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
      Height          =   360
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "[iJukenNumber]"
      Top             =   1890
      Width           =   1455
   End
   Begin VB.ComboBox cboExamineeStatus 
      Enabled         =   0   'False
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   30
      Tag             =   "[iExamineeStatus]"
      Top             =   5175
      Width           =   1935
   End
   Begin VB.ComboBox cboRoomProfileId 
      Enabled         =   0   'False
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   29
      Tag             =   "[iRoomProfileId]"
      Top             =   4800
      Width           =   1935
   End
   Begin VB.ComboBox cboPreferenceDay1Flag 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   26
      Tag             =   "[iPreferenceDay1Flag]"
      Top             =   3675
      Width           =   1455
   End
   Begin VB.ComboBox cboPreferenceDay2Flag 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   27
      Tag             =   "[iPreferenceDay2Flag]"
      Top             =   4050
      Width           =   1455
   End
   Begin VB.ComboBox cboPreferenceDay3Flag 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   28
      Tag             =   "[iPreferenceDay3Flag]"
      Top             =   4425
      Width           =   1455
   End
   Begin VB.ComboBox cboSuisenFlag 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   25
      Tag             =   "[iSuisenFlagId]"
      Top             =   3300
      Width           =   1455
   End
   Begin VB.ComboBox cboMultipleApplyFlag 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   24
      Tag             =   "[iMultipleApplyFlag]"
      Top             =   2925
      Width           =   1455
   End
   Begin VB.ComboBox cboAbsentFlag 
      Enabled         =   0   'False
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   23
      Tag             =   "[iAbsentFlag]"
      Top             =   2550
      Width           =   1455
   End
   Begin VB.ComboBox cboRejectFlag 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   22
      Tag             =   "[iRejectFlag]"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2463"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11175
      TabIndex        =   18
      Top             =   1020
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2463"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11175
      TabIndex        =   20
      Top             =   1410
      Width           =   375
   End
   Begin VB.ComboBox cboSex 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   21
      Tag             =   "[iSex]"
      Top             =   1785
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSearchGrid 
      Height          =   1440
      Left            =   240
      TabIndex        =   39
      Top             =   8160
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   2540
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtSecondDayExam 
      Enabled         =   0   'False
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
      Height          =   360
      Left            =   3360
      TabIndex        =   16
      Tag             =   "[dtSecondExamDay]"
      Top             =   7515
      Width           =   2175
   End
   Begin VB.ComboBox cboPhysicalConditionId 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   34
      Tag             =   "[iPhysicalConditionId]"
      Top             =   6675
      Width           =   1935
   End
   Begin VB.ComboBox cboParentJobC 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   33
      Tag             =   "[iParentJobCategory]"
      Top             =   6300
      Width           =   1935
   End
   Begin VB.ComboBox cboFamilyID 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   31
      Tag             =   "[iFamilyId]"
      Top             =   5550
      Width           =   1935
   End
   Begin VB.ComboBox cboUniversityType 
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
      Height          =   360
      Left            =   9720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   32
      Tag             =   "[iUniversityType]"
      Top             =   5925
      Width           =   1935
   End
   Begin VB.TextBox txtdtBirthDay 
      DataField       =   "dtBirthDay"
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
      Height          =   360
      Left            =   3360
      TabIndex        =   15
      Tag             =   "[dtBirthDay]"
      Top             =   7095
      Width           =   2175
   End
   Begin VB.TextBox txtHighSchoolId 
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
      Height          =   360
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   19
      Tag             =   "[iHighSchoolId]"
      Top             =   1410
      Width           =   1455
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
      Height          =   360
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   17
      Tag             =   "[iZipCodeId]"
      Top             =   1020
      Width           =   1455
   End
   Begin VB.TextBox txtNendo 
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
      Height          =   360
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "[iNendo]"
      Top             =   1455
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtBirthDay 
      Height          =   255
      Left            =   4680
      TabIndex        =   37
      Top             =   7380
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "MM/DD/YYYY"
      Format          =   54853635
      CurrentDate     =   37223
   End
   Begin MSComCtl2.DTPicker dtSecondDayExam 
      Height          =   255
      Left            =   10680
      TabIndex        =   38
      Top             =   1785
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "MM/DD/YYYY"
      Format          =   54853633
      CurrentDate     =   37223
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   33
      Left            =   9480
      TabIndex        =   109
      Top             =   7455
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1836"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   108
      Tag             =   "1829"
      Top             =   7455
      Width           =   2385
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   34
      Left            =   9465
      TabIndex        =   107
      Top             =   7065
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1835"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6945
      TabIndex        =   106
      Tag             =   "1829"
      Top             =   7065
      Width           =   2385
   End
   Begin VB.Label lblErrIndicator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   18
      Left            =   9480
      TabIndex        =   105
      Top             =   5925
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   32
      Left            =   9480
      TabIndex        =   104
      Top             =   2925
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   31
      Left            =   9480
      TabIndex        =   103
      Top             =   4425
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   30
      Left            =   9480
      TabIndex        =   102
      Top             =   4050
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   29
      Left            =   9480
      TabIndex        =   101
      Top             =   3675
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   25
      Left            =   9480
      TabIndex        =   100
      Top             =   6675
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   23
      Left            =   9480
      TabIndex        =   99
      Top             =   3300
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   21
      Left            =   9480
      TabIndex        =   98
      Top             =   6300
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   20
      Left            =   9480
      TabIndex        =   97
      Top             =   5550
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   17
      Left            =   3120
      TabIndex        =   96
      Top             =   7500
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   16
      Left            =   9480
      TabIndex        =   95
      Top             =   5175
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   15
      Left            =   9480
      TabIndex        =   94
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   9480
      TabIndex        =   93
      Top             =   1410
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   14
      Left            =   9480
      TabIndex        =   92
      Top             =   2550
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   13
      Left            =   9480
      TabIndex        =   91
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   9480
      TabIndex        =   90
      Top             =   1785
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   9480
      TabIndex        =   89
      Top             =   1020
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblNendo 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1804"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   88
      Tag             =   "1804"
      Top             =   1455
      Width           =   2745
   End
   Begin VB.Label lblZipCodeID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1806"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   87
      Tag             =   "1806"
      Top             =   1020
      Width           =   2445
   End
   Begin VB.Label lblSex 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1808"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   86
      Tag             =   "1808"
      Top             =   1785
      Width           =   2445
   End
   Begin VB.Label lblHighSchoolId 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1810"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   85
      Tag             =   "1810"
      Top             =   1410
      Width           =   2445
   End
   Begin VB.Label lbldtBirthDay 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1812"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   84
      Tag             =   "1812"
      Top             =   7095
      Width           =   2745
   End
   Begin VB.Label lblRoomProfileID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1815"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   83
      Tag             =   "1815"
      Top             =   4800
      Width           =   2445
   End
   Begin VB.Label lblAbsentFlag 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1816"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   82
      Tag             =   "16"
      Top             =   2550
      Width           =   2445
   End
   Begin VB.Label lblRejectFlag 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1814"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   81
      Tag             =   "1814"
      Top             =   2160
      Width           =   2445
   End
   Begin VB.Label lblExamineeStatus 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1817"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   80
      Tag             =   "1817"
      Top             =   5175
      Width           =   2445
   End
   Begin VB.Label lblUniversityType 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1832"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   79
      Tag             =   "1832"
      Top             =   5925
      Width           =   2445
   End
   Begin VB.Label lblFamilyID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1831"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   78
      Tag             =   "1831"
      Top             =   5550
      Width           =   2445
   End
   Begin VB.Label lblParentJobCategory 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1833"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6975
      TabIndex        =   77
      Tag             =   "1833"
      Top             =   6300
      Width           =   2445
   End
   Begin VB.Label lblSuisenFlagID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1824"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   76
      Tag             =   "1824"
      Top             =   3300
      Width           =   2445
   End
   Begin VB.Label lblPhysicalConditionID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1834"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   75
      Tag             =   "1834"
      Top             =   6675
      Width           =   2445
   End
   Begin VB.Label lblPreferenceDay1Flag 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1826"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   74
      Tag             =   "1826"
      Top             =   3675
      Width           =   2445
   End
   Begin VB.Label lblPreferenceDay2Flag 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1828"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   73
      Tag             =   "1828"
      Top             =   4050
      Width           =   2445
   End
   Begin VB.Label lbliPreferenceDay3Flag 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1830"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   72
      Tag             =   "1830"
      Top             =   4425
      Width           =   2445
   End
   Begin VB.Label lblMultipleApplyFlag 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1822"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   71
      Tag             =   "1822"
      Top             =   2925
      Width           =   2445
   End
   Begin VB.Label lblSecondDayExam 
      BackStyle       =   0  'ìßñæ
      Caption         =   "1819"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   70
      Tag             =   "1819"
      Top             =   7500
      Width           =   2775
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   69
      Top             =   1455
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   68
      Top             =   7095
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblJukenNumber 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1803"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   67
      Tag             =   "1803"
      Top             =   1890
      Width           =   2745
   End
   Begin VB.Label lblExamineeName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1805"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   66
      Tag             =   "1805"
      Top             =   2280
      Width           =   2745
   End
   Begin VB.Label lblKanaName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1807"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   65
      Tag             =   "1807"
      Top             =   2670
      Width           =   2745
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1809"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   64
      Tag             =   "1809"
      Top             =   3045
      Width           =   2745
   End
   Begin VB.Label lblTelephoneNo 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1811"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   63
      Tag             =   "1811"
      Top             =   4020
      Width           =   2745
   End
   Begin VB.Label lblEmailAddress 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1813"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   62
      Tag             =   "1813"
      Top             =   4395
      Width           =   2745
   End
   Begin VB.Label lblBackgroudID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1821"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   240
      TabIndex        =   61
      Tag             =   "1821"
      Top             =   5160
      Width           =   2745
   End
   Begin VB.Label lblQualificationID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1823"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   60
      Tag             =   "1823"
      Top             =   5550
      Width           =   2745
   End
   Begin VB.Label lblNationality 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1818"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   59
      Tag             =   "1818"
      Top             =   4770
      Width           =   2745
   End
   Begin VB.Label lblScienceSubjprofileID2 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1829"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   225
      TabIndex        =   58
      Tag             =   "1829"
      Top             =   6690
      Width           =   2745
   End
   Begin VB.Label lblScienceSubjprofileID1 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1827"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Tag             =   "1827"
      Top             =   6315
      Width           =   2745
   End
   Begin VB.Label lblLanguageSubjProfileID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1825"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   56
      Tag             =   "1825"
      Top             =   5940
      Width           =   2745
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   55
      Top             =   1890
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   54
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   53
      Top             =   2670
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   52
      Top             =   3045
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   51
      Top             =   4020
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   3120
      TabIndex        =   50
      Top             =   4395
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblExamineeProfileId 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1802"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   49
      Tag             =   "1802"
      Top             =   1020
      Width           =   2745
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   19
      Left            =   3120
      TabIndex        =   48
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   22
      Left            =   3120
      TabIndex        =   47
      Top             =   5550
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   24
      Left            =   3120
      TabIndex        =   46
      Top             =   4770
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   26
      Left            =   3120
      TabIndex        =   45
      Top             =   5940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   27
      Left            =   3120
      TabIndex        =   44
      Top             =   6315
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   28
      Left            =   3120
      TabIndex        =   43
      Top             =   6690
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   42
      Top             =   1020
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   35
      Left            =   12840
      TabIndex        =   41
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrorMsg 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   7860
      Visible         =   0   'False
      Width           =   11895
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmExamineeProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmExamineeProfile
'Author         :   Vishal Kamath
'Created On     :   18/9/01
'Description    :   This form makes a provision for updating data in ExamineeProfile Table.
'Reference      :   FunctionalSpecs OF MaintainExamineeData.doc(Ver1.0)
'**************************************************************************************************
'Modification History
'Mahesh 22/5/2002
'Modified as an exception to module accomodate in the form with background picture
'Made font size 10

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

Public Function fnFillGrid(ByVal StrSearchString As String)
    
    Call Form_Activate  '  - this will initialize the grid , when coming from suisen also
    m_bload = True
    FillGrid (StrSearchString)
End Function

Private Sub chkMultipleApplyFlag_Click()
    SetChange
End Sub

Private Sub chkPreferenceDay1Flag_Click()
    SetChange
End Sub

Private Sub chkPreferenceDay2Flag_Click()
    SetChange
End Sub

Private Sub chkPreferenceday3flag_Click()
    SetChange
End Sub

Private Sub chkSuisenFlagId_Click()
    SetChange
End Sub

Private Sub cboAbsentFlag_Click()
    SetChange
End Sub

Private Sub cboAdmissionType1_Click()
    SetChange
End Sub

Private Sub cboAdmissionType2_Click()
    SetChange
End Sub

Private Sub cbobackGroundId_Click()
    SetChange
End Sub

Private Sub cboExamineeStatus_Click()
    SetChange
End Sub

Private Sub cboFamilyID_Click()
    SetChange
End Sub

Private Sub cboLanguageSubjProfileID_Click()
    SetChange
End Sub

Private Sub cboMultipleApplyFlag_Click()
    SetChange
End Sub

Private Sub cboParentJobC_Click()
    SetChange
End Sub

Private Sub cboPhysicalConditionId_Click()
    SetChange
End Sub

Private Sub cboPreferenceDay1Flag_Click()
    SetChange
End Sub

Private Sub cboPreferenceDay2Flag_Click()
    SetChange
End Sub

Private Sub cboPreferenceDay3Flag_Click()
    SetChange
End Sub

Private Sub cboQualificationID_Click()
    SetChange
End Sub

Private Sub cboRejectFlag_Click()
    SetChange
End Sub

Private Sub cboRoomProfileId_Click()
    SetChange
End Sub

Private Sub cboScienceSubjprofileID1_Click()
    SetChange
End Sub

Private Sub cboScienceSubjprofileID2_Click()
    SetChange
End Sub

Private Sub cboSex_Click()
    SetChange
End Sub

Private Sub cboSuisenFlag_Click()
    SetChange
End Sub

Private Sub cboUniversityType_Click()
    SetChange
End Sub

Private Sub chkAbsentFlag_Click()
    SetChange
End Sub

Private Sub Command1_Click()
    dlgChgHighSchType.Show 1
End Sub

Private Sub Command2_Click()
    dlgChangeZip.Show 1
End Sub

Private Sub dtBirthDay_Change()
    dtBirthDay.Visible = False
    txtdtBirthDay.Text = g_dt_ConvertDate(dtBirthDay.Value)
    txtdtBirthDay.ZOrder 0
End Sub

Private Sub dtBirthDay_LostFocus()
    dtBirthDay.Visible = False
    txtdtBirthDay.Text = g_dt_ConvertDate(dtBirthDay.Value)
    txtdtBirthDay.ZOrder 0
End Sub

Private Sub dtSecondDayExam_Change()
    dtSecondDayExam.Visible = False
    txtSecondDayExam.Text = g_dt_ConvertDate(dtSecondDayExam.Value)
    txtSecondDayExam.ZOrder 0
End Sub

Private Sub dtSecondDayExam_LostFocus()
    dtSecondDayExam.Visible = False
    txtSecondDayExam.Text = g_dt_ConvertDate(dtSecondDayExam.Value)
    txtSecondDayExam.ZOrder 0
End Sub

Private Sub Form_Activate()
 On Error GoTo ErrHandler
    Dim lngRow As Integer
    Dim Index
    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       If Index <> 2 Then
            fMainForm.Toolbar1.Buttons(Index).Enabled = True
        Else
            fMainForm.Toolbar1.Buttons(Index).Enabled = False
        End If
    Next
    If m_bload = False Then
        m_lngSearchGridTop = 5900   '5750 ' changed by team
        
        'initialize the search grid
        InitializeSearchGrid
        'populate the grid with empty row and column headings
        SearchRecords True
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
    On Error GoTo ErrHandler
    ' set the table name
    LoadResStrings Me
    Me.Caption = LoadResString(1004)
    m_TableName = "tbSTEExamineeProfile"
        
    Call g_void_SetFontProperties(Me)     ' set the font properties
    Dim ctl As Control
    Dim l_str_ctlType As String
    'New code of font size 10 added (Mahesh) 22/5/2002
    For Each ctl In Me
        With ctl
        l_str_ctlType = TypeName(ctl)
        Select Case l_str_ctlType
        Case "CommandButton", "ComboBox", "Label", "OptionButton", "CheckBox", "VSFlexGrid", "MSFlexGrid", "MSHFlexGrid", "dXSideBar", "ListBox", "DTPicker"
        .Font.Size = 10
        Case "TextBox"
        .Font.Size = 10
        If .MultiLine = False Then
        .Height = 360
        End If
        End Select
        End With
    Next
    'New Code ends here
    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained
    'Made these fields nonMandatory ->[vTelephone],[vEmailAddress],[iAbsentFlag],[iRejectFlag]
    '[iExamineeStatus],[iUniversityType],[iFamilyId],[iParentJobCategory],[iSuisenFlagId]
    '[iPhysicalConditionId]
    With m_colFieldDetails
        .Add "[iExamineeProfileId]", "[Examinee Profile Id]", 1, True, True, "", "INTEGER", 1800, "", "[iExamineeProfileId]"
        .Add "[iJukenNumber]", "[Juken Number]", 2, True, False, "", "INTEGER", 1500, "", "[iJukenNumber]", "right('0000'+convert(varchar,iJukenNumber),4)", "0000"
        .Add "[iNendo]", "[Nendo]", 3, True, False, "", "INTEGER", 1300, "", "[iNendo]"
        .Add "[vExamineeName]", "[Examinee Name]", 4, True, False, "", "STRING", 1700, "", "[vExamineeName]"
        .Add "[vKanaName]", "[Kana Name]", 5, False, False, "", "STRING", 1500, "", "[vKanaName]"
        'Changed by Mahesh 21/5/2002 Address in now mandatory
        .Add "[vAddress]", "[Address]", 6, True, False, "", "STRING", 1500, "", "[vAddress]"
        'Change ends
        .Add "[iZipCodeId]", "[Zip Code Id]", 7, True, False, "", "INTEGER", 1500, "", "[iZipCodeId]"
        .Add "[iSex]", "[Sex]", 8, True, False, "", "COMBO", 1200, "", "[iSex]"
        .Add "[iHighSchoolId]", "[High School Id]", 9, True, False, "", "INTEGER", 1500, "", "[iHighSchoolId]"
        'Not mandatory Changed on 21/5/2002
        .Add "[vTelephone]", "[Telephone]", 10, False, False, "", "STRING", 1500, "", "[vTelephone]"
        'Not mandatory Changed on 21/5/2002
        .Add "[vEmailAddress]", "[Email Address]", 11, False, False, "", "STRING", 1500, "", "[vEmailAddress]"
        .Add "[dtBirthDay]", "[Birth Day]", 12, True, False, "", "DATE", 1500, "", "[dtBirthDay]", "dbo.usfCpfGetJapanDateFromDt(dtBirthDay)", "gggeeîNmmåéddì˙"
        .Add "[iRoomProfileId]", "[Room Profile Id]", 13, False, False, "", "COMBO", 1500, "", "[iRoomProfileId]"
        'Not mandatory Changed on 21/5/2002
        .Add "[iAbsentFlag]", "[Absent Flag]", 14, False, False, "", "COMBO", 1500, "", "[iAbsentFlag]"
        'Not mandatory Changed on 21/5/2002
        .Add "[iRejectFlag]", "[Reject Flag]", 15, False, False, "", "COMBO", 1500, "", "[iRejectFlag]"
        'Not mandatory Changed on 21/5/2002
        .Add "[iExamineeStatus]", "[Examinee Status]", 16, False, False, "", "COMBO", 1500, "", "[iExamineeStatus]"

''''    .Add "[dtSecondExamDay]", "[Second Exam Day]", 17, False, False, "", "DATE", 1800, "", "[dtSecondExamDay]", "dbo.usfCpfGetJapanDateFromDt(dtSecondExamDay)", "gggeeîNmmåéddì˙"
        .Add "[dtSecondExamDay]", "[Second Exam Day]", 17, False, False, "", "DATE", 1800, "", "[dtSecondExamDay]", "dbo.usfCpfGetJapanDateFromDt(dtSecondExamDay)", "yyyyîNmmåéddì˙"

        'Not mandatory Changed on 21/5/2002
        .Add "[iUniversityType]", "[University Type]", 18, False, False, "", "COMBO", 1500, "", "[iUniversityType]"
        .Add "[iBackgroundId]", "[Background Id]", 19, True, False, "", "COMBO", 1500, "", "[iBackgroundId]"
        'Not mandatory Changed on 21/5/2002
        .Add "[iFamilyId]", "[Family Id]", 20, False, False, "", "COMBO", 1500, "", "[iFamilyId]"
        'Not mandatory Changed on 21/5/2002
        .Add "[iParentJobCategory]", "[Parent Job Category]", 21, False, False, "", "COMBO", 2000, "", "[iParentJobCategory]"
        .Add "[iQualificationId]", "[Qualification Id]", 22, True, False, "", "COMBO", 1500, "", "[iQualificationId]"
        'Not mandatory Changed on 21/5/2002
        .Add "[iSuisenFlagId]", "[Suisen Flag Id]", 23, False, False, "", "COMBO", 2000, "", "[iSuisenFlagId]"
        .Add "[vNationality]", "[Nationality]", 24, True, False, "", "STRING", 1500, "", "[vNationality]"
        'Not mandatory Changed on 21/5/2002
        .Add "[iPhysicalConditionId]", "[Physical Condition Id]", 25, False, False, "", "COMBO", 2000, "", "[iPhysicalConditionId]"
        .Add "[iLanguageSubjProfileId]", "[Language Subj Profile Id]", 26, True, False, "", "COMBO", 2300, "", "[iLanguageSubjProfileId]"
        .Add "[iScienceSubjProfileId1]", "[Science Subj Profile Id 1]", 27, True, False, "", "COMBO", 2100, "", "[iScienceSubjProfileId1]"
        .Add "[iScienceSubjProfileId2]", "[Language Subj Profile Id 2]", 28, True, False, "", "COMBO", 2100, "", "[iScienceSubjProfileId2]"
        .Add "[iPreferenceDay1Flag]", "[Preference Day1 Flag]", 29, True, False, "", "COMBO", 2000, "", "[iPreferenceDay1Flag]"
        .Add "[iPreferenceDay2Flag]", "[iPreference Day2 Flag]", 30, True, False, "", "COMBO", 2000, "", "[iPreferenceDay2Flag]"
        .Add "[iPreferenceDay3Flag]", "[iPreference Day3 Flag]", 31, True, False, "", "COMBO", 2000, "", "[iPreferenceDay3Flag]"
        .Add "[iMultipleApplyFlag]", "[Multiple Apply Flag]", 32, True, False, "", "COMBO", 1700, "", "[iMultipleApplyFlag]"
        .Add "[iAdmissionType1]", "[Admission Type1]", 33, False, False, "", "COMBO", 1700, "", "[iAdmissionType1]"     ' change made on 31/07/02
        .Add "[iAdmissionType2]", "[Admission Type2]", 34, False, False, "", "COMBO", 1700, "", "[iAdmissionType2]"     ' change made on 31/07/02
    End With
    
    ' functions to populate the combo boxes
    Call f_void_AddUnivType
    Call f_void_AddBackgroundID
    Call f_void_AddFamilyID
    Call f_void_AddParentJobCategory
    Call f_void_AddQualificationID
    Call f_void_PhysicalConditionID
    Call f_void_AddLanguageSubjProfileID
    Call f_void_AddScienceSubjProfileID1
    Call f_void_AddScienceSubjProfileID2
    Call f_void_AddSex
    Call f_void_AddRoomProfileID
    Call f_void_AddExamineeStatus
    Call f_void_AddAdmissionType ' change made on 31/07/02
    
    ' added on 271101
    
    With cboMultipleApplyFlag
        .AddItem ""
        .AddItem LoadResString(2058)
        .AddItem LoadResString(2059)
        m_ComboDetails.Add 0, LoadResString(2058), "[iMultipleApplyFlag]"
        m_ComboDetails.Add 1, LoadResString(2059), "[iMultipleApplyFlag]"
    End With
    
    With cboSuisenFlag
        .AddItem ""
        .AddItem LoadResString(2060)
        .AddItem LoadResString(2061)
        m_ComboDetails.Add 0, LoadResString(2060), "[iSuisenFlagId]"
        m_ComboDetails.Add 1, LoadResString(2061), "[iSuisenFlagId]"
    End With
    
    With cboPreferenceDay1Flag
        .AddItem ""
        .AddItem LoadResString(2451)
        .AddItem LoadResString(2452)
        m_ComboDetails.Add 1, LoadResString(2451), "[iPreferenceDay1Flag]"
        m_ComboDetails.Add 0, LoadResString(2452), "[iPreferenceDay1Flag]"
    End With
    
    With cboPreferenceDay2Flag
        .AddItem ""
        .AddItem LoadResString(2451)
        .AddItem LoadResString(2452)
        m_ComboDetails.Add 1, LoadResString(2451), "[iPreferenceDay2Flag]"
        m_ComboDetails.Add 0, LoadResString(2452), "[iPreferenceDay2Flag]"
    End With
    
    With cboPreferenceDay3Flag
        .AddItem ""
        .AddItem LoadResString(2451)
        .AddItem LoadResString(2452)
        m_ComboDetails.Add 1, LoadResString(2451), "[iPreferenceDay3Flag]"
        m_ComboDetails.Add 0, LoadResString(2452), "[iPreferenceDay3Flag]"
    End With
    
    With cboRejectFlag
        .AddItem ""
        .AddItem LoadResString(2459)
        .AddItem LoadResString(2460)
        m_ComboDetails.Add 0, LoadResString(2459), "[iRejectFlag]"
        m_ComboDetails.Add 1, LoadResString(2460), "[iRejectFlag]"
    End With
    
    With cboAbsentFlag
        .AddItem ""
        .AddItem LoadResString(2062)
        .AddItem LoadResString(2063)
        m_ComboDetails.Add 0, LoadResString(2062), "[iAbsentFlag]"
        m_ComboDetails.Add 1, LoadResString(2063), "[iAbsentFlag]"
    End With
    
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
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    On Error GoTo ErrorHandler
    
    AssignValues True
    
    If Len(Trim(txtZipCodeId.Text)) <> 0 Then
        l_str_Sql = "SELECT vPrefectureName, vCityName, vAddress1, vAddress2 FROM tbSTEZipCodeMaster"
        l_str_Sql = l_str_Sql & " WHERE vZipCodeName='" & Trim(txtZipCodeId.Text) & "'"
        l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
        If Not l_obj_Rst.EOF Then
            txtZipAddress.Text = l_obj_Rst("vPrefectureName") & "," & l_obj_Rst("vCityName") & "," & _
            l_obj_Rst("vAddress1") & "," & l_obj_Rst("vAddress2")
        End If
        l_obj_Rst.Close
        Set l_obj_Rst = Nothing
    End If
    Call g_void_HighlightRow(hfgSearchGrid.Row, f_int_PrevRow)
    f_int_PrevRow = hfgSearchGrid.Row
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
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
    If Trim(cboScienceSubjprofileID1.Text) = Trim(cboScienceSubjprofileID2.Text) Then
        MsgBox LoadResString(2467), vbInformation
        cboScienceSubjprofileID1.SetFocus
        ExtraValidation = False
    Else
        ExtraValidation = True
    End If
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Private Sub txtAddress_Change()
    SetChange
End Sub

Private Sub txtdtBirthDay_Change()
    SetChange
End Sub

Private Sub txtEmployeeProfileID_Change()
    SetChange
End Sub

Private Sub txtEmployeeProfileID_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtdtBirthDay_GotFocus()
    ' position date picker control over cell
    With txtdtBirthDay
        dtBirthDay.Move .Left, .Top, .Width, .Height
        If Trim(.Text) <> "" Then
            ' initialize value, save original in tag in case user hits escape
            'dtBirthDay.Value = .Text
            'dtBirthDay.Tag = .Text
            
            ' changed the above two lines in Comdesign , arka 9Apr 2002
            
            dtBirthDay.Value = g_dt_ConvertDate(.Text)
            dtBirthDay.Tag = g_dt_ConvertDate(.Text)
        Else
            dtBirthDay.Tag = #1/1/1900#
        End If
        ' show and activate date picker control
        dtBirthDay.Visible = True
        dtBirthDay.SetFocus
    End With
    ' make it drop down the calendar
    SendKeys "{f4}"
End Sub

Private Sub txtExamineeName_Change()
    SetChange
End Sub

Private Sub txtExamineeProfileId_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtExamineeStatus_Change()
    SetChange
End Sub

Private Sub txtHighSchoolId_Change()
    SetChange
End Sub

Private Sub txtJukenNumber_Change()
    SetChange
End Sub

Private Sub txtJukenNumber_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> Asc("_") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtKanaName_Change()
    SetChange
End Sub

Private Sub txtMultipleApplyFlag_Change()
    SetChange
End Sub

Private Sub txtNationality_Change()
    SetChange
End Sub

Private Sub txtNendo_Change()
    SetChange
End Sub

Private Sub txtPreferenceDay1Flag_Change()
    SetChange
End Sub

Private Sub txtPreferenceDay2Flag_Change()
    SetChange
End Sub

Private Sub txtPreferenceDay3Flag_Change()
    SetChange
End Sub

Private Sub txtRejectFlag_Change()
    SetChange
End Sub

Private Sub txtNendo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRoomProfileId_Change()
    SetChange
End Sub

Private Sub txtSex_Change()
    SetChange
End Sub

Private Sub txtSuisenFlagId_Change()
    SetChange
End Sub

Private Sub txtSecondDayExam_Change()
    SetChange
End Sub

Private Sub txtSecondDayExam_GotFocus()
    ' position date picker control over cell
    With txtSecondDayExam
        dtSecondDayExam.Move .Left, .Top, .Width, .Height
        If Trim(.Text) <> "" Then
            ' initialize value, save original in tag in case user hits escape
            
            'dtSecondDayExam.Value = .Text
            'dtSecondDayExam.Tag = .Text
            
          ' changed in comdesign , arka 9Apr 2002
          
            dtSecondDayExam.Value = g_dt_ConvertDate(.Text)
            dtSecondDayExam.Tag = g_dt_ConvertDate(.Text)
        Else
            dtSecondDayExam.Tag = #1/1/1900#
        End If
        ' show and activate date picker control
        dtSecondDayExam.Visible = True
        dtSecondDayExam.SetFocus
    End With
    ' make it drop down the calendar
    SendKeys "{f4}"
End Sub

Private Sub txtTelephoneNo_Change()
    SetChange
End Sub

Private Sub txtTelephoneNo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtvEmailAddress_Change()
    SetChange
End Sub

Private Sub txtZipCodeId_Change()
    SetChange
End Sub

Private Sub f_void_AddUnivType()
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iLookUpTableId,vName from tbSTELookUpTable WHERE iLookUpTableType = 1"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboUniversityType.AddItem ""
    
    While Not l_obj_Rst.EOF
       cboUniversityType.AddItem l_obj_Rst.Fields("vName")
       m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboUniversityType.Tag
       l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddBackgroundID()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iLookUpTableId,vName from tbSTELookUpTable WHERE iLookUpTableType =2"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
        
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
         
    ' add a blank row to the combobox
    cbobackGroundId.AddItem ""
    
    While Not l_obj_Rst.EOF
       cbobackGroundId.AddItem l_obj_Rst.Fields("vName")
       m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cbobackGroundId.Tag
       l_obj_Rst.MoveNext
    Wend
     l_obj_Rst.Close
     Set l_obj_Rst = Nothing
     Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddFamilyID()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iLookUpTableId,vName from tbSTELookUpTable WHERE iLookUpTableType =3"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboFamilyID.AddItem ""

    While Not l_obj_Rst.EOF
       cboFamilyID.AddItem l_obj_Rst.Fields("vName")
       m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboFamilyID.Tag
       l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddParentJobCategory()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iLookUpTableId,vName from tbSTELookUpTable WHERE iLookUpTableType =4"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboParentJobC.AddItem ""
    
    While Not l_obj_Rst.EOF
       cboParentJobC.AddItem l_obj_Rst.Fields("vName")
       m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboParentJobC.Tag
       l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddQualificationID()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iLookUpTableId,vName from tbSTELookUpTable WHERE iLookUpTableType =5"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboQualificationID.AddItem ""
    
    While Not l_obj_Rst.EOF
       cboQualificationID.AddItem l_obj_Rst.Fields("vName")
       m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboQualificationID.Tag
       l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_PhysicalConditionID()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iLookUpTableId,vName from tbSTELookUpTable WHERE iLookUpTableType =6"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboPhysicalConditionId.AddItem ""
    
    While Not l_obj_Rst.EOF
       cboPhysicalConditionId.AddItem l_obj_Rst.Fields("vName")
       m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboPhysicalConditionId.Tag
       l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub


Private Sub f_void_AddLanguageSubjProfileID()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iLookUpTableId,vName from tbSTELookUpTable WHERE iLookUpTableType =7"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboLanguageSubjProfileID.AddItem ""
    
    While Not l_obj_Rst.EOF
       cboLanguageSubjProfileID.AddItem l_obj_Rst.Fields("vName")
       m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboLanguageSubjProfileID.Tag
       l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddScienceSubjProfileID1()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iLookUpTableId,vName from tbSTELookUpTable WHERE iLookUpTableType =8"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
        l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboScienceSubjprofileID1.AddItem ""
    
    While Not l_obj_Rst.EOF
        cboScienceSubjprofileID1.AddItem l_obj_Rst.Fields("vName")
        m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboScienceSubjprofileID1.Tag
        l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddScienceSubjProfileID2()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iLookUpTableId,vName from tbSTELookUpTable WHERE iLookUpTableType =8"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
        l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboScienceSubjprofileID2.AddItem ""
    
    While Not l_obj_Rst.EOF
        cboScienceSubjprofileID2.AddItem l_obj_Rst.Fields("vName")
        m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboScienceSubjprofileID2.Tag
        l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddSex()
    On Error GoTo ErrorHandler
    cboSex.AddItem ""
    cboSex.AddItem LoadResString(1837)
    cboSex.AddItem LoadResString(1838)
    m_ComboDetails.Add 0, LoadResString(1837), "[iSex]"
    m_ComboDetails.Add 1, LoadResString(1838), "[iSex]"
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddRoomProfileID()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iRoomProfileId,vRoomName from tbSTERoomProfile"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboRoomProfileId.AddItem ""
    
    While Not l_obj_Rst.EOF
       cboRoomProfileId.AddItem l_obj_Rst.Fields("vRoomName")
       m_ComboDetails.Add l_obj_Rst.Fields("iRoomProfileId"), l_obj_Rst.Fields("vRoomName"), cboRoomProfileId.Tag
       l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddExamineeStatus()
    cboExamineeStatus.AddItem ""
    cboExamineeStatus.AddItem LoadResString(2446)
    cboExamineeStatus.AddItem LoadResString(2447)
    cboExamineeStatus.AddItem LoadResString(2448)
    cboExamineeStatus.AddItem LoadResString(2449)
    cboExamineeStatus.AddItem LoadResString(2450)
    m_ComboDetails.Add 0, LoadResString(2446), "[iExamineeStatus]"
    m_ComboDetails.Add 1, LoadResString(2447), "[iExamineeStatus]"
    m_ComboDetails.Add 2, LoadResString(2448), "[iExamineeStatus]"
    m_ComboDetails.Add 3, LoadResString(2449), "[iExamineeStatus]"
    m_ComboDetails.Add 6, LoadResString(2450), "[iExamineeStatus]"
End Sub

' change made on 31/07/02
Private Sub f_void_AddAdmissionType()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iAdmissionType,vAdmissionName from tbSTEAdmissionType "
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboAdmissionType1.AddItem ""
    cboAdmissionType2.AddItem ""
    
    While Not l_obj_Rst.EOF
       cboAdmissionType1.AddItem l_obj_Rst.Fields("vAdmissionName")
       cboAdmissionType2.AddItem l_obj_Rst.Fields("vAdmissionName")
       m_ComboDetails.Add l_obj_Rst.Fields("iAdmissionType"), l_obj_Rst.Fields("vAdmissionName"), cboAdmissionType1.Tag
       m_ComboDetails.Add l_obj_Rst.Fields("iAdmissionType"), l_obj_Rst.Fields("vAdmissionName"), cboAdmissionType2.Tag
       l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
