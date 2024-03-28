VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExamineeCheck 
   AutoRedraw      =   -1  'True
   Caption         =   "frmExamineeCheck : "
   ClientHeight    =   10755
   ClientLeft      =   1050
   ClientTop       =   435
   ClientWidth     =   15300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmExamineeCheck.frx":0000
   ScaleHeight     =   10755
   ScaleWidth      =   15300
   Tag             =   "1004"
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.TextBox txtZipPref 
      Enabled         =   0   'False
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
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   106
      TabStop         =   0   'False
      Tag             =   "[iZipCodeId]"
      Top             =   1005
      Width           =   1215
   End
   Begin VB.TextBox txtHighSchoolCode 
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   105
      TabStop         =   0   'False
      Tag             =   "[iZipCodeId]"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtZipCode 
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
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   104
      TabStop         =   0   'False
      Tag             =   "[iZipCodeId]"
      Top             =   1492
      Width           =   1215
   End
   Begin VB.TextBox txtExamineeProfileID 
      Enabled         =   0   'False
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
      Left            =   12000
      MaxLength       =   10
      TabIndex        =   101
      Tag             =   "[iExamineeProfileID]"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cmbDepartment 
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
      Left            =   6960
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   96
      TabStop         =   0   'False
      Tag             =   "[iDepartment]"
      Top             =   3330
      Width           =   1815
   End
   Begin VB.TextBox txtDepartment 
      DataField       =   "iDepartment"
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   6240
      MaxLength       =   1
      TabIndex        =   13
      Top             =   3330
      Width           =   495
   End
   Begin VB.ComboBox cmbCourse 
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
      Left            =   2400
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   93
      TabStop         =   0   'False
      Tag             =   "[iCourse]"
      Top             =   3330
      Width           =   1815
   End
   Begin VB.TextBox txtCourse 
      DataField       =   "iCourse"
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3330
      Width           =   495
   End
   Begin VB.TextBox txtSuisenFlag 
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   9645
      TabIndex        =   20
      Top             =   4935
      Width           =   855
   End
   Begin VB.TextBox txtParentJobC 
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   5445
      TabIndex        =   19
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtFamilyID 
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   1320
      TabIndex        =   18
      Top             =   4935
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "ämíË"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12330
      TabIndex        =   28
      Top             =   5970
      Width           =   1170
   End
   Begin VB.TextBox txtAddress 
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
      IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
      Left            =   5880
      TabIndex        =   1
      Tag             =   "[vAddress]"
      Top             =   1005
      Width           =   6015
   End
   Begin VB.TextBox txtNendo 
      Enabled         =   0   'False
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
      Left            =   1665
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "[iNendo]"
      Top             =   1005
      Width           =   1215
   End
   Begin VB.TextBox txtSex 
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   5040
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2482
      Width           =   495
   End
   Begin VB.TextBox txtAdmissionType2 
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   15
      Top             =   4095
      Width           =   495
   End
   Begin VB.TextBox txtAdmissionType 
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   14
      Top             =   4095
      Width           =   495
   End
   Begin VB.TextBox txtUnivName 
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
      IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
      Left            =   5160
      TabIndex        =   16
      Tag             =   "[vUnivName]"
      Top             =   4095
      Width           =   2655
   End
   Begin VB.TextBox txtAge 
      DataField       =   "dtBirthDay"
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
      Left            =   10080
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2482
      Width           =   495
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
      Left            =   12000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'êÇíº
      TabIndex        =   78
      Top             =   2025
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   240
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "[iAdmissionType1]"
      Top             =   4530
      Width           =   1815
   End
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
      Left            =   2160
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "[iAdmissionType2]"
      Top             =   4530
      Width           =   1815
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
      IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
      Left            =   1665
      TabIndex        =   7
      Tag             =   "[vExamineeName]"
      Top             =   2482
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
      IMEMode         =   6  'îºäp∂¿∂≈
      Left            =   1665
      TabIndex        =   6
      Tag             =   "[vKanaName]"
      Top             =   2017
      Width           =   2775
   End
   Begin VB.TextBox txtNationality 
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
      IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
      Left            =   4440
      TabIndex        =   3
      Tag             =   "[vNationality]"
      Top             =   1492
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
      Left            =   9975
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   23
      Tag             =   "[iScienceSubjProfileId2]"
      Top             =   5430
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
      Left            =   6660
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   22
      Tag             =   "[iScienceSubjProfileId1]"
      Top             =   5430
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
      Left            =   2040
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   21
      Tag             =   "[iLanguageSubjProfileId]"
      Top             =   5430
      Width           =   1935
   End
   Begin VB.TextBox txtJukenNumber 
      Enabled         =   0   'False
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
      Left            =   1665
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "[iJukenNumber]"
      Top             =   1492
      Width           =   1455
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
      Left            =   1815
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   24
      Tag             =   "[iPreferenceDay1Flag]"
      Top             =   6120
      Width           =   975
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
      Left            =   3135
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   25
      Tag             =   "[iPreferenceDay2Flag]"
      Top             =   6120
      Width           =   975
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
      Left            =   4575
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   26
      Tag             =   "[iPreferenceDay3Flag]"
      Top             =   6120
      Width           =   975
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
      ItemData        =   "frmExamineeCheck.frx":3AD3
      Left            =   10560
      List            =   "frmExamineeCheck.frx":3AD5
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "[iSuisenFlagId]"
      Top             =   4935
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
      Left            =   7815
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   27
      Tag             =   "[iMultipleApplyFlag]"
      Top             =   6120
      Width           =   2310
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9120
      TabIndex        =   5
      Top             =   1500
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
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
      Left            =   2760
      TabIndex        =   11
      Top             =   2925
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
      Left            =   5640
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "[iSex]"
      Top             =   2482
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSearchGrid 
      Height          =   2670
      Left            =   240
      TabIndex        =   38
      Top             =   6645
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   4710
      _Version        =   393216
      ForeColor       =   16777088
      ForeColorSel    =   8421631
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Left            =   6360
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "[iParentJobCategory]"
      Top             =   4935
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
      Left            =   2160
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "[iFamilyId]"
      Top             =   4935
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
      Left            =   9360
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   17
      Tag             =   "[iUniversityType]"
      Top             =   4095
      Width           =   1095
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
      Left            =   7440
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "[dtBirthDay]"
      Top             =   2482
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
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "[iHighSchoolId]"
      Top             =   2970
      Visible         =   0   'False
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
      Left            =   12000
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "[iZipCodeId]"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtBirthDay 
      Height          =   360
      Left            =   12000
      TabIndex        =   100
      Top             =   2655
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      CustomFormat    =   "yyyyîNMMåéDDì˙"
      Format          =   54657027
      CurrentDate     =   37223
   End
   Begin VB.Label lblErrIndicatorDel 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   27
      Left            =   12615
      TabIndex        =   103
      Top             =   900
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblExamineeProfileID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "îNìx"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   12030
      TabIndex        =   102
      Tag             =   "1804"
      Top             =   900
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblHighSchoolCode 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "highschool code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   163
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1680
      TabIndex        =   99
      Tag             =   "1810"
      Top             =   2940
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblDepartment 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "äwâ»"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4800
      TabIndex        =   98
      Top             =   3390
      Width           =   1065
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   26
      Left            =   6000
      TabIndex        =   97
      Top             =   3390
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblCourse 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "â€íˆ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   95
      Top             =   3390
      Width           =   1065
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   25
      Left            =   1425
      TabIndex        =   94
      Top             =   3420
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHighSchoolName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "highschool name"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   6480
      TabIndex        =   92
      Tag             =   "1810"
      Top             =   2985
      Width           =   2550
   End
   Begin VB.Label lblHighSchoolType 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "highschool type"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   4800
      TabIndex        =   91
      Tag             =   "1810"
      Top             =   2970
      Width           =   1620
   End
   Begin VB.Label lblHighSchoolPref 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "highschool pref"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   3180
      TabIndex        =   90
      Tag             =   "1810"
      Top             =   2970
      Width           =   1695
   End
   Begin VB.Label lblPrefDay 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ñ ê⁄äÛñ]ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   4530
      TabIndex        =   89
      Tag             =   "1826"
      Top             =   5865
      Width           =   1140
   End
   Begin VB.Label lblPrefDay 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ñ ê⁄äÛñ]ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   3090
      TabIndex        =   88
      Tag             =   "1826"
      Top             =   5865
      Width           =   1140
   End
   Begin VB.Label lblPrefDay 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ñ ê⁄äÛñ]ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   1785
      TabIndex        =   87
      Tag             =   "1826"
      Top             =   5865
      Width           =   1110
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   4230
      TabIndex        =   86
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "èZèä"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3375
      TabIndex        =   85
      Tag             =   "1809"
      Top             =   1050
      Width           =   825
   End
   Begin VB.Label lblNendo 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "îNìx"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   210
      TabIndex        =   84
      Tag             =   "1804"
      Top             =   1065
      Width           =   480
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   83
      Top             =   1058
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   24
      Left            =   4935
      TabIndex        =   82
      Top             =   4170
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUnivName 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ëÂäwñº"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4080
      TabIndex        =   81
      Tag             =   "1812"
      Top             =   4155
      Width           =   705
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   9840
      TabIndex        =   80
      Top             =   2535
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblAge 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "îNóÓ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9990
      TabIndex        =   79
      Tag             =   "1812"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   23
      Left            =   3240
      TabIndex        =   36
      Top             =   4155
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "åªòQãÊï™ÇQ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2190
      TabIndex        =   37
      Tag             =   "1829"
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   22
      Left            =   1425
      TabIndex        =   77
      Top             =   4155
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "åªòQãÊï™ÇP"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   76
      Tag             =   "1829"
      Top             =   4155
      Width           =   1065
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   10
      Left            =   9120
      TabIndex        =   75
      Top             =   4185
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   21
      Left            =   7635
      TabIndex        =   74
      Top             =   6180
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   20
      Left            =   4335
      TabIndex        =   73
      Top             =   6180
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   19
      Left            =   2955
      TabIndex        =   72
      Top             =   6180
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   18
      Left            =   1575
      TabIndex        =   71
      Top             =   6180
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   13
      Left            =   9450
      TabIndex        =   70
      Top             =   5010
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   5205
      TabIndex        =   69
      Top             =   4995
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   1080
      TabIndex        =   68
      Top             =   4995
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   67
      Top             =   2955
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   4800
      TabIndex        =   66
      Top             =   2535
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   7680
      TabIndex        =   65
      Top             =   1545
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblZipCodeID 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "óXï÷î‘çÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6480
      TabIndex        =   64
      Tag             =   "1806"
      Top             =   1545
      Width           =   1125
   End
   Begin VB.Label lblSex 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ê´ï "
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4995
      TabIndex        =   63
      Tag             =   "1808"
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label lblHighSchoolId 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "çÇçZID"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   62
      Tag             =   "1810"
      Top             =   2940
      Width           =   1005
   End
   Begin VB.Label lbldtBirthDay 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ê∂îNåéì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7440
      TabIndex        =   61
      Tag             =   "1812"
      Top             =   2145
      Width           =   945
   End
   Begin VB.Label lblUniversityType 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ëÂäwãÊï™"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7920
      TabIndex        =   60
      Tag             =   "1832"
      Top             =   4140
      Width           =   1155
   End
   Begin VB.Label lblFamilyID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "â∆ë∞ID"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   59
      Tag             =   "1831"
      Top             =   4995
      Width           =   885
   End
   Begin VB.Label lblParentJobCategory 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "óºêeédéñãÊï™"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4260
      TabIndex        =   58
      Tag             =   "1833"
      Top             =   4935
      Width           =   975
   End
   Begin VB.Label lblSuisenFlagID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "êÑëEÉtÉâÉOID"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   8505
      TabIndex        =   57
      Tag             =   "1824"
      Top             =   4935
      Width           =   975
   End
   Begin VB.Label lblPreferenceDay1Flag 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ñ ê⁄äÛñ]ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   255
      TabIndex        =   56
      Tag             =   "1826"
      Top             =   6165
      Width           =   1125
   End
   Begin VB.Label lblMultipleApplyFlag 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ïπäËÉtÉâÉO"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6675
      TabIndex        =   55
      Tag             =   "1822"
      Top             =   6165
      Width           =   1020
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   7200
      TabIndex        =   54
      Top             =   2535
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblJukenNumber 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "éÛå±î‘çÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Tag             =   "1803"
      Top             =   1545
      Width           =   1065
   End
   Begin VB.Label lblExamineeName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "éÛå±é“éÅñº"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Tag             =   "1805"
      Top             =   2535
      Width           =   1065
   End
   Begin VB.Label lblKanaName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "Ç©Ç»éÅñº"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Tag             =   "1807"
      Top             =   2070
      Width           =   825
   End
   Begin VB.Label lblNationality 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "çëê–"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3585
      TabIndex        =   50
      Tag             =   "1818"
      Top             =   1485
      Width           =   600
   End
   Begin VB.Label lblScienceSubjprofileID2 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ëIëóùâ»ÇQ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8820
      TabIndex        =   49
      Tag             =   "1827"
      Top             =   5505
      Width           =   1065
   End
   Begin VB.Label lblScienceSubjprofileID1 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ëIëóùâ»ÇP"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5520
      TabIndex        =   48
      Tag             =   "1827"
      Top             =   5505
      Width           =   1110
   End
   Begin VB.Label lblLanguageSubjProfileID 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ëIëäOçëåÍ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   47
      Tag             =   "1825"
      Top             =   5490
      Width           =   1425
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   46
      Top             =   1545
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   45
      Top             =   2535
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   44
      Top             =   2070
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   14
      Left            =   4230
      TabIndex        =   43
      Top             =   1515
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   15
      Left            =   1800
      TabIndex        =   42
      Top             =   5490
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   16
      Left            =   6510
      TabIndex        =   41
      Top             =   5490
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   17
      Left            =   9825
      TabIndex        =   40
      Top             =   5490
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
      TabIndex        =   39
      Top             =   9360
      Visible         =   0   'False
      Width           =   13290
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmExamineeCheck"
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
Private prvbNoDirtyCheck As Boolean

Private Sub Form_Load()

    On Error GoTo ErrHandler

    Dim ctl           As Control
    Dim l_str_ctlType As String


    gbExamCheckNewShow = False


    ' set the table name
    LoadResStrings Me

    Me.Caption = "frmExamineeCheck : éÛå±é“ÉfÅ[É^ÇÃï“èW"
    m_TableName = "tbSTEExamineeProfile"
        
    Call g_void_SetFontProperties(Me)     'set the font properties

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


#If 0 Then '2021.11.14 del jhi
    With m_colFieldDetails
        .Add "[iJukenNumber]", "[Juken Number]", 1, True, False, "", "INTEGER", 1500, "", "[iJukenNumber]", "right('0000'+convert(varchar,iJukenNumber),4)", "0000"
        .Add "[iNendo]", "[Nendo]", 2, True, False, "", "INTEGER", 1300, "", "[iNendo]"
        .Add "[vExamineeName]", "[Examinee Name]", 3, True, False, "", "STRING", 1700, "", "[vExamineeName]"
        .Add "[vKanaName]", "[Kana Name]", 4, False, False, "", "STRING", 1500, "", "[vKanaName]"
        .Add "[vAddress]", "[Address]", 5, True, False, "", "STRING", 1500, "", "[vAddress]"
        .Add "[iZipCodeId]", "[Zip Code Id]", 6, True, False, "", "INTEGER", 1500, "", "[iZipCodeId]"
        .Add "[iSex]", "[Sex]", 7, True, False, "", "COMBO", 1200, "", "[iSex]"
        .Add "[iHighSchoolId]", "[High School Id]", 8, True, False, "", "INTEGER", 1500, "", "[iHighSchoolId]"
        .Add "[dtBirthDay]", "[Birth Day]", 9, True, False, "", "DATE", 1500, "", "[dtBirthDay]", "dbo.usfCpfGetJapanDateFromDt(dtBirthDay)", "gggeeîNmmåéddì˙"
        .Add "[iUniversityType]", "[University Type]", 10, False, False, "", "COMBO", 1500, "", "[iUniversityType]"
        .Add "[iFamilyId]", "[Family Id]", 11, False, False, "", "COMBO", 1500, "", "[iFamilyId]"
        .Add "[iParentJobCategory]", "[Parent Job Category]", 12, False, False, "", "COMBO", 2000, "", "[iParentJobCategory]"
        .Add "[iSuisenFlagId]", "[Suisen Flag Id]", 13, False, False, "", "COMBO", 2000, "", "[iSuisenFlagId]"
        .Add "[vNationality]", "[Nationality]", 14, False, False, "", "STRING", 1500, "", "[vNationality]"
        .Add "[iLanguageSubjProfileId]", "[Language Subj Profile Id]", 15, True, False, "", "COMBO", 2300, "", "[iLanguageSubjProfileId]"
        .Add "[iScienceSubjProfileId1]", "[Science Subj Profile Id 1]", 16, True, False, "", "COMBO", 2100, "", "[iScienceSubjProfileId1]"
        .Add "[iScienceSubjProfileId2]", "[Language Subj Profile Id 2]", 17, True, False, "", "COMBO", 2100, "", "[iScienceSubjProfileId2]"
        .Add "[iPreferenceDay1Flag]", "[Preference Day1 Flag]", 18, True, False, "", "COMBO", 2000, "", "[iPreferenceDay1Flag]"
        .Add "[iPreferenceDay2Flag]", "[iPreference Day2 Flag]", 19, True, False, "", "COMBO", 2000, "", "[iPreferenceDay2Flag]"
        .Add "[iPreferenceDay3Flag]", "[iPreference Day3 Flag]", 20, True, False, "", "COMBO", 2000, "", "[iPreferenceDay3Flag]"
        .Add "[iMultipleApplyFlag]", "[Multiple Apply Flag]", 21, True, False, "", "COMBO", 1700, "", "[iMultipleApplyFlag]"
        .Add "[iAdmissionType1]", "[Admission Type1]", 22, True, False, "", "COMBO", 1700, "", "[iAdmissionType1]"           'change made on 31/07/02
        .Add "[iAdmissionType2]", "[Admission Type2]", 23, True, False, "", "COMBO", 1700, "", "[iAdmissionType2]"           'change made on 31/07/02
        .Add "[vUnivName]", "[UnivName]", 24, False, False, "", "STRING", 1700, "", "[vUnivName]"                            'change made on 31/07/02
        .Add "[iCourse]", "[Course]", 25, True, False, "", "COMBO", 1700, "", "[iCourse]"                                    'change made on 17/12/04
        .Add "[iDepartment]", "[Department]", 26, True, False, "", "COMBO", 1700, "", "[iDepartment]"                        'change made on 17/12/04
        .Add "[iExamineeProfileId]", "[ExamineeProfileId]", 27, True, True, "", "INTEGER", 0, "", "[iExamineeProfileId]"     'change made on 17/12/04
    End With
#End If

#If 1 Then '2021.11.14 add jhi
    With m_colFieldDetails
        .Add "[iJukenNumber]", "[éÛå±î‘çÜ]", 1, True, False, "", "INTEGER", 950, "", "[iJukenNumber]", "right('0000'+convert(varchar,iJukenNumber),4)", "0000"
        .Add "[iNendo]", "[îNìx]", 2, True, False, "", "INTEGER", 520, "", "[iNendo]"
        .Add "[vExamineeName]", "[éÛå±é“éÅñº]", 3, True, False, "", "STRING", 1300, "", "[vExamineeName]"
        .Add "[vKanaName]", "[Ç©Ç»éÅñº]", 4, False, False, "", "STRING", 1300, "", "[vKanaName]"
        .Add "[vAddress]", "[èZèä]", 5, True, False, "", "STRING", 2500, "", "[vAddress]"
        .Add "[iZipCodeId]", "[óXï÷î‘çÜ]", 6, True, False, "", "INTEGER", 900, "", "[iZipCodeId]"
        .Add "[iSex]", "[ê´ï ]", 7, True, False, "", "COMBO", 300, "", "[iSex]"
        .Add "[iHighSchoolId]", "[çÇçZID]", 8, True, False, "", "INTEGER", 800, "", "[iHighSchoolId]"
        .Add "[dtBirthDay]", "[ê∂îNåéì˙]", 9, True, False, "", "DATE", 1500, "", "[dtBirthDay]", "dbo.usfCpfGetJapanDateFromDt(dtBirthDay)", "yyyyîNmmåéddì˙"
        .Add "[iUniversityType]", "[ëÂäwãÊï™]", 10, False, False, "", "COMBO", 1500, "", "[iUniversityType]"
        .Add "[iFamilyId]", "[â∆ë∞ID]", 11, False, False, "", "COMBO", 1500, "", "[iFamilyId]"
        .Add "[iParentJobCategory]", "[óºêeédéñãÊï™]", 12, False, False, "", "COMBO", 2000, "", "[iParentJobCategory]"
        .Add "[iSuisenFlagId]", "[êÑëEÉtÉâÉOID]", 13, False, False, "", "COMBO", 2000, "", "[iSuisenFlagId]"
        .Add "[vNationality]", "[çëê–]", 14, False, False, "", "STRING", 900, "", "[vNationality]"
        .Add "[iLanguageSubjProfileId]", "[ëIëäOçëåÍ]", 15, True, False, "", "COMBO", 2300, "", "[iLanguageSubjProfileId]"
        .Add "[iScienceSubjProfileId1]", "[ëIëóùâ»1]", 16, True, False, "", "COMBO", 2100, "", "[iScienceSubjProfileId1]"
        .Add "[iScienceSubjProfileId2]", "[ëIëóùâ»2]", 17, True, False, "", "COMBO", 2100, "", "[iScienceSubjProfileId2]"
        .Add "[iPreferenceDay1Flag]", "[ñ ê⁄äÛñ]ì˙1]", 18, True, False, "", "COMBO", 2000, "", "[iPreferenceDay1Flag]"
        .Add "[iPreferenceDay2Flag]", "[ñ ê⁄äÛñ]ì˙2]", 19, True, False, "", "COMBO", 2000, "", "[iPreferenceDay2Flag]"
        .Add "[iPreferenceDay3Flag]", "[ñ ê⁄äÛñ]ì˙3]", 20, True, False, "", "COMBO", 2000, "", "[iPreferenceDay3Flag]"
        .Add "[iMultipleApplyFlag]", "[ïπäËÉtÉâÉO]", 21, True, False, "", "COMBO", 1700, "", "[iMultipleApplyFlag]"
        .Add "[iAdmissionType1]", "[åªòQãÊï™1]", 22, True, False, "", "COMBO", 1700, "", "[iAdmissionType1]"             'change made on 31/07/02
        .Add "[iAdmissionType2]", "[åªòQãÊï™2]", 23, True, False, "", "COMBO", 1700, "", "[iAdmissionType2]"             'change made on 31/07/02
        .Add "[vUnivName]", "[ëÂäwñº]", 24, False, False, "", "STRING", 1500, "", "[vUnivName]"                          'change made on 31/07/02
        .Add "[iCourse]", "[â€íˆ]", 25, True, False, "", "COMBO", 1700, "", "[iCourse]"                                  'change made on 17/12/04
        .Add "[iDepartment]", "[äwâ»]", 26, True, False, "", "COMBO", 1700, "", "[iDepartment]"                          'change made on 17/12/04
        .Add "[iExamineeProfileId]", "[ExamineeProfileId]", 27, True, True, "", "INTEGER", 0, "", "[iExamineeProfileId]" 'change made on 17/12/04
    End With

#End If


    
    ' functions to populate the combo boxes
    Call f_void_AddUnivType
    Call f_void_AddFamilyID
    Call f_void_AddParentJobCategory
    Call f_void_AddLanguageSubjProfileID
    Call f_void_AddScienceSubjProfileID1
    Call f_void_AddScienceSubjProfileID2
    Call f_void_AddSex
    Call f_void_AddAdmissionType ' change made on 31/07/02
    Call f_void_AddCourse
    Call f_void_AddDepartment
    

    'added on 2021.11.27

    With cboMultipleApplyFlag
        .AddItem ""
        .AddItem LoadResString(2058)
        .AddItem LoadResString(2059)
        m_ComboDetails.Add 0, LoadResString(2058), "[iMultipleApplyFlag]"
        m_ComboDetails.Add 1, LoadResString(2059), "[iMultipleApplyFlag]"
    End With

    With cboSuisenFlag
        .AddItem ""
        .ItemData(.NewIndex) = -1
        .AddItem LoadResString(2060)
        .ItemData(.NewIndex) = 0
        .AddItem LoadResString(2061)
        .ItemData(.NewIndex) = 1
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

'ì˙ïtï\é¶

    Dim sSQL As String
    Dim oRs  As New ADODB.Recordset

    sSQL = ""
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "  isnull( substring( convert( varchar , dtSecondExamDay1 , 111 ) , 6 , 5 ) , '' ) "
    sSQL = sSQL & " ,isnull( substring( convert( varchar , dtSecondExamDay2 , 111 ) , 6 , 5 ) , '' ) "
    sSQL = sSQL & " ,isnull( substring( convert( varchar , dtSecondExamDay3 , 111 ) , 6 , 5 ) , '' ) "
    sSQL = sSQL & " FROM tbSTEsecondexamprofile "

    Set oRs = g_obj_Conn.Execute(sSQL)

    If Not oRs.EOF Then
        If oRs.Fields(0) <> "" Then
            lblPrefDay(0).Caption = Format(oRs.Fields(0), "MM/DD")
        Else
            lblPrefDay(0).Caption = ""
        End If
        If oRs.Fields(1) <> "" Then
            lblPrefDay(1).Caption = Format(oRs.Fields(1), "MM/DD")
        Else
            lblPrefDay(1).Caption = ""
        End If
        If oRs.Fields(2) <> "" Then
            lblPrefDay(2).Caption = Format(oRs.Fields(2), "MM/DD")
        Else
            lblPrefDay(2).Caption = ""
        End If
    End If

    oRs.Close
    Set oRs = Nothing

    lblHighSchoolCode.FontSize = 12
    lblHighSchoolName.FontSize = 12
    lblHighSchoolPref.FontSize = 12
    lblHighSchoolType.FontSize = 12
'    lblZipCode.FontSize = 12

    m_bload = False

    'add,2007/11/08,S------
    'ñ‚ëËì_ëŒâûå¬êlñºÇÃóìÇëÂÇ´Ç≠Ç∑ÇÈ
    Me.txtExamineeName.FontSize = 12
    'add,2007/11/08,E------

    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbInformation, "ÉGÉâÅ[" ''''LoadResString(1729)

End Sub

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
    Call s_clickCmb2ChangTxt(cboAdmissionType1, txtAdmissionType)
    txtAdmissionType2.Text = txtAdmissionType.Text
    cboAdmissionType2.ListIndex = cboAdmissionType1.ListIndex
End Sub

Private Sub cboAdmissionType2_Click()
    SetChange
    Call s_clickCmb2ChangTxt(cboAdmissionType2, txtAdmissionType2)
End Sub

Private Sub cbobackGroundId_Click()
    SetChange
End Sub

Private Sub cboExamineeStatus_Click()
    SetChange
End Sub

Private Sub cboFamilyID_Click()
    SetChange
    Call s_clickCmb2ChangTxtGrp(cboFamilyID, txtFamilyID)
End Sub

Private Sub cboLanguageSubjProfileID_Click()
    SetChange
End Sub

Private Sub cboMultipleApplyFlag_Click()
    SetChange
End Sub

Private Sub cboParentJobC_Click()
    SetChange
    Call s_clickCmb2ChangTxtGrp(cboParentJobC, txtParentJobC)
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
'M,Fï\é¶Ç≈ÅAì¡éÍÇ»ÇΩÇﬂÅAÇ±Ç±ÇÕå¬ï èàóù
    If cboSex.ListIndex <> -1 Then
        If cboSex.ItemData(cboSex.ListIndex) = 0 Then
            txtSex.Text = "M"
        ElseIf cboSex.ItemData(cboSex.ListIndex) = 1 Then
            txtSex.Text = "F"
        Else
            txtSex.Text = ""
        End If
    Else
        txtSex.Text = ""
    End If
End Sub

Private Sub cboSuisenFlag_Click()
    SetChange
    Call s_clickCmb2ChangTxt(cboSuisenFlag, txtSuisenFlag)
End Sub

Private Sub cboUniversityType_Click()
    SetChange
End Sub

Private Sub chkAbsentFlag_Click()
    SetChange
End Sub

Private Sub cmbCourse_Click()
    SetChange
    Call s_clickCmb2ChangTxt(cmbCourse, txtCourse)
    'add,2007/11/08,S----------
    'í êMÅEíËéûêßÇÃèÍçáÅAåªòQãÊï™ÇÕÇXÇ…Ç∑ÇÈ
    If txtCourse = 2 Or txtCourse = 3 Then
        cboAdmissionType1.ListIndex = cboAdmissionType1.ListCount - 1
    End If
    'add,2007/11/08,E----------
End Sub

Private Sub cmbDepartment_Click()
    SetChange
    Call s_clickCmb2ChangTxt(cmbDepartment, txtDepartment)
End Sub

Private Sub cmdUpdate_Click()
    lblErrorMsg.Caption = ""
    Call ValidateAndSaveData
    If lblErrorMsg.Caption = "ÉfÅ[É^ÇÕï€ë∂Ç≥ÇÍÇ‹ÇµÇΩ" Then
        lblErrorMsg.Caption = ""
        If hfgSearchGrid.Row + 1 < hfgSearchGrid.Rows Then
            hfgSearchGrid.Row = hfgSearchGrid.Row + 1
            Call hfgSearchGrid_DblClick
            hfgSearchGrid.TopRow = hfgSearchGrid.Row
        ElseIf Me.hfgSearchGrid.Row + 1 = Me.hfgSearchGrid.Rows Then
            frmExamCheckPara.Show
            frmExamCheckPara.ZOrder 0
            prvbNoDirtyCheck = True
            Unload Me
        End If
    End If
End Sub

Private Sub Command1_Click()
    Set dlgChgHighSchTypeForCheck.goParentForm = Me
    dlgChgHighSchTypeForCheck.Show 1
End Sub

Private Sub Command2_Click()
    Set dlgChangeZipForCheck.goParentForm = Me
    dlgChangeZipForCheck.Show 1
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
' On Error GoTo ErrHandler
'    Dim lngRow As Integer
'    Dim Index
'    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
'       If Index <> 2 Then
'            fMainForm.Toolbar1.Buttons(Index).Enabled = True
'        Else
'            fMainForm.Toolbar1.Buttons(Index).Enabled = False
'        End If
'    Next
'    If m_bload = False Then
'        m_lngSearchGridTop = 5900   '5750 ' changed by team
'
'        'initialize the search grid
'        InitializeSearchGrid
'        'populate the grid with empty row and column headings
'        SearchRecords False
'        fMainForm.mnuToolsSave.Enabled = False
'        fMainForm.mnuToolsCancel.Enabled = False
'
'        fMainForm.Toolbar1.Buttons("Save").Enabled = False
'        fMainForm.Toolbar1.Buttons("Cancel").Enabled = False
'        If Me.hfgSearchGrid.Rows > 1 Then
'            hfgSearchGrid.Row = 1
'            Call hfgSearchGrid_DblClick
'        End If
'    End If
'    m_bload = True
'    fMainForm.mnuTools.Enabled = True
'
'    Exit Sub
'ErrHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
    On Error GoTo ErrorHandler
    fMainForm.mnuTools.Enabled = False  ' disable tools menu
    Dim Index
    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next
    If m_bload = False Then
        m_lngSearchGridTop = 5900   '5750 ' changed by team

        'initialize the search grid
        InitializeSearchGrid
        'populate the grid with empty row and column headings
        SearchRecords False

        If Me.hfgSearchGrid.Rows > 1 Then
            hfgSearchGrid.Row = 1
            Call hfgSearchGrid_DblClick
        End If
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If prvbNoDirtyCheck = False Then
        If CheckDirty = False Then
            Cancel = 1
            Exit Sub
        End If
    End If
    gbExamCheckNewShow = True
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
        l_str_Sql = "SELECT vZipCodeName , vPrefectureName, vCityName, vAddress1, vAddress2 FROM tbSTEZipCodeMaster"
        l_str_Sql = l_str_Sql & " WHERE iZipCodeId='" & Trim(txtZipCodeId.Text) & "'"
        l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
        If Not l_obj_Rst.EOF Then
'            lblZipCode.Caption = l_obj_Rst("vZipCodeName")
            txtZipCode.Text = l_obj_Rst("vZipCodeName")
            txtZipAddress.Text = l_obj_Rst("vPrefectureName") & "," & l_obj_Rst("vCityName") & "," & _
                                    l_obj_Rst("vAddress1") & "," & l_obj_Rst("vAddress2")
            txtZipPref.Text = l_obj_Rst("vPrefectureName")
        End If
        l_obj_Rst.Close
        Set l_obj_Rst = Nothing
    End If
    
'add,xzg,2006/12/21,S------------------
'çÇçZèÓïÒÇçƒï\é¶Ç∑ÇÈëOÇ…ÅAèâä˙âªÇçsÇ§
            lblHighSchoolCode.Caption = ""
            lblHighSchoolName.Caption = ""
            lblHighSchoolPref.Caption = ""
            lblHighSchoolType.Caption = ""
'add,xzg,2006/12/21,E------------------

    If Len(Trim(txtHighSchoolID.Text)) <> 0 Then
        'HighScoolProfile DataGet
        'iZipCodeId -> ZipCodeMaster.vPrefectureName -> KenMei
        'vHighSchoolName -> KoukouMei
        'iHighSchoolRecommendation -> KouritsuShiritsuKubun
        'iHighSchoolRecommendation -> Katei(Zennnishi , yakann )
        'iHighSchoolRecommendation -> gakka
        l_str_Sql = "SELECT vHighSchoolCode, vHighSchoolName FROM tbSTEHighSchoolType"
        l_str_Sql = l_str_Sql & " WHERE iHighSchoolID='" & Trim(txtHighSchoolID.Text) & "'"
        l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
        If Not l_obj_Rst.EOF Then
            lblHighSchoolCode.Caption = l_obj_Rst("vHighSchoolCode")
            txtHighSchoolCode.Text = lblHighSchoolCode.Caption
            lblHighSchoolName.Caption = l_obj_Rst("vHighSchoolName")
            lblHighSchoolPref.Caption = f_str_HighSchoolPref(lblHighSchoolCode.Caption)
            lblHighSchoolType.Caption = f_str_HighSchoolType(lblHighSchoolCode.Caption)
        Else
            lblHighSchoolCode.Caption = ""
            txtHighSchoolCode.Text = ""
            lblHighSchoolName.Caption = ""
            lblHighSchoolPref.Caption = ""
            lblHighSchoolType.Caption = ""
        End If
        l_obj_Rst.Close
        Set l_obj_Rst = Nothing
    Else
        lblHighSchoolCode.Caption = ""
        txtHighSchoolCode.Text = ""
        lblHighSchoolName.Caption = ""
        lblHighSchoolPref.Caption = ""
        lblHighSchoolType.Caption = ""
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

Private Sub txtAdmissionType_Change()
    SetChange
    Call txtAdmissionType_GotFocus
End Sub

Private Sub txtAdmissionType_GotFocus()
    txtAdmissionType.SelStart = 0
    txtAdmissionType.SelLength = Len(txtAdmissionType.Text)
End Sub

Private Sub txtAdmissionType_LostFocus()
    If Not IsNumeric(txtAdmissionType.Text) Then
        txtAdmissionType.Text = "0"
    ElseIf CInt(txtAdmissionType.Text) > cboAdmissionType1.ListCount Or CInt(txtAdmissionType.Text) < 1 Then
        txtAdmissionType.Text = "0"
    Else
        cboAdmissionType1.ListIndex = CInt(txtAdmissionType.Text) + 1
        txtAdmissionType2.Text = txtAdmissionType.Text
        cboAdmissionType2.ListIndex = cboAdmissionType1.ListIndex
    End If
    SetChange
End Sub

Private Sub txtAdmissionType_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Or _
        KeyAscii = vbKey1 Or KeyAscii = vbKey2 Or _
        KeyAscii = vbKey3 Or KeyAscii = vbKey4 Or _
        KeyAscii = vbKey5 Or KeyAscii = vbKey6 Or _
        KeyAscii = vbKey7 Or KeyAscii = vbKey8 Or _
        KeyAscii = vbKey9 Or KeyAscii = vbKey0) _
            Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtAdmissionType2_Change()
    SetChange
    Call txtAdmissionType2_GotFocus
End Sub

Private Sub txtAdmissionType2_GotFocus()
    txtAdmissionType2.SelStart = 0
    txtAdmissionType2.SelLength = Len(txtAdmissionType2.Text)
End Sub

Private Sub txtAdmissionType2_LostFocus()
    If Not IsNumeric(txtAdmissionType2.Text) Then
        txtAdmissionType2.Text = "0"
    ElseIf CInt(txtAdmissionType2.Text) > cboAdmissionType2.ListCount Or CInt(txtAdmissionType2.Text) < 1 Then
        txtAdmissionType2.Text = "0"
    Else
        cboAdmissionType2.ListIndex = CInt(txtAdmissionType2.Text) + 1
    End If
    SetChange
End Sub

Private Sub txtAdmissionType2_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Or _
        KeyAscii = vbKey1 Or KeyAscii = vbKey2 Or _
        KeyAscii = vbKey3 Or KeyAscii = vbKey4 Or _
        KeyAscii = vbKey5 Or KeyAscii = vbKey6 Or _
        KeyAscii = vbKey7 Or KeyAscii = vbKey8 Or _
        KeyAscii = vbKey9 Or KeyAscii = vbKey0) _
            Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtAge_Change()
    SetChange
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Or _
        KeyAscii = vbKey1 Or KeyAscii = vbKey2 Or _
        KeyAscii = vbKey3 Or KeyAscii = vbKey4 Or _
        KeyAscii = vbKey5 Or KeyAscii = vbKey6 Or _
        KeyAscii = vbKey7 Or KeyAscii = vbKey8 Or _
        KeyAscii = vbKey9 Or KeyAscii = vbKey0) _
            Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtCourse_Change()
    SetChange
    Call txtCourse_GotFocus
End Sub

Private Sub txtCourse_GotFocus()
    txtCourse.SelStart = 0
    txtCourse.SelLength = Len(txtCourse.Text)
End Sub

Private Sub txtCourse_LostFocus()
Dim bChkErr As Boolean
    If Trim(txtCourse) <> "" Then
        bChkErr = True
        For l_int_Counter = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
            If Trim(Me.m_ComboDetails.Item(l_int_Counter).Value) = Trim(txtCourse) And UCase(Trim(Me.m_ComboDetails.Item(l_int_Counter).GroupId)) = UCase(cmbCourse.Tag) Then
                If Me.m_ComboDetails.Item(l_int_Counter).Description <> "" Then
                    cmbCourse.Text = Me.m_ComboDetails.Item(l_int_Counter).Description
                    bChkErr = False
                    Exit For
                End If
            End If
        Next
        If bChkErr Then
            txtCourse.Text = ""
            cmbCourse.ListIndex = 0
        End If
    Else
        cmbCourse.ListIndex = 0
    End If
End Sub

Private Sub txtDepartment_Change()
    SetChange
    Call txtDepartment_GotFocus
End Sub

Private Sub txtDepartment_GotFocus()
    txtDepartment.SelStart = 0
    txtDepartment.SelLength = Len(txtDepartment.Text)
End Sub

Private Sub txtDepartment_LostFocus()
Dim bChkErr As Boolean
    If Trim(txtDepartment) <> "" Then
        bChkErr = True
        For l_int_Counter = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
            If Trim(Me.m_ComboDetails.Item(l_int_Counter).Value) = Trim(txtDepartment) And UCase(Trim(Me.m_ComboDetails.Item(l_int_Counter).GroupId)) = UCase(cmbDepartment.Tag) Then
                If Me.m_ComboDetails.Item(l_int_Counter).Description <> "" Then
                    cmbDepartment.Text = Me.m_ComboDetails.Item(l_int_Counter).Description
                    bChkErr = False
                    Exit For
                End If
            End If
        Next
        If bChkErr Then
            txtDepartment.Text = ""
            cmbDepartment.ListIndex = 0
        End If
    Else
        cmbDepartment.ListIndex = 0
    End If
End Sub

Private Sub txtdtBirthDay_Change()
Dim iAge As Integer
    If IsDate(txtdtBirthDay.Text) Then
        iAge = DateDiff("yyyy", CDate(txtdtBirthDay.Text), Now)
'update,xzg,2006/12/21,S------------------------
'ê∂îNåéì˙ÇÊÇËÅAîNóÓÇåvéZÇ∑ÇÈ
        'txtAge.Text = Trim(str(iAge))
        If Format(Now, "mm/dd") >= Format(CDate(txtdtBirthDay.Text), "mm/dd") Then
            txtAge.Text = Trim(str(iAge))
        Else
            txtAge.Text = Trim(str(iAge - 1))
            If txtAge.Text < 0 Then
                txtAge.Text = 0
            End If
        End If
'update,xzg,2006/12/21,E------------------------
    End If
    SetChange
End Sub

Private Sub txtdtBirthDay_LostFocus()
Dim iAge As Integer
    If IsDate(txtdtBirthDay.Text) Then
        iAge = DateDiff("yyyy", CDate(txtdtBirthDay.Text), Now)
        'update,xzg,2007/01/15,S------------------------
'ê∂îNåéì˙ÇÊÇËÅAîNóÓÇåvéZÇ∑ÇÈ
        'txtAge.Text = Trim(str(iAge))
        If Format(Now, "mm/dd") >= Format(CDate(txtdtBirthDay.Text), "mm/dd") Then
            txtAge.Text = Trim(str(iAge))
        Else
            txtAge.Text = Trim(str(iAge - 1))
            If txtAge.Text < 0 Then
                txtAge.Text = 0
            End If
        End If
'update,xzg,2007/01/15,E------------------------
    End If
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

Private Sub txtFamilyID_Change()
    SetChange
    Call txtFamilyID_GotFocus
End Sub

Private Sub txtFamilyID_GotFocus()
    txtFamilyID.SelStart = 0
    txtFamilyID.SelLength = Len(txtFamilyID.Text)
End Sub

Private Sub txtFamilyID_LostFocus()
Dim bChkErr As Boolean
    If Trim(txtFamilyID) <> "" Then
        bChkErr = True
        For l_int_Counter = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
            If Trim(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).GroupValue) = Trim(txtFamilyID) And UCase(Trim(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).GroupId)) = UCase(cboFamilyID.Tag) Then      'UCase(fMainForm.ActiveForm.m_colFieldDetails.Item(i_int_lngCol).DBFieldName) Then
                If fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Description <> "" Then
                    cboFamilyID.Text = fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Description
                    bChkErr = False
                    Exit For
                End If
            End If
        Next
        If bChkErr Then
            txtFamilyID.Text = ""
            cboFamilyID.ListIndex = 0
        End If
    Else
        cboFamilyID.ListIndex = 0
    End If
End Sub

Private Sub txtParentJobC_Change()
    SetChange
    Call txtParentJobC_GotFocus
End Sub

Private Sub txtParentJobC_GotFocus()
    txtParentJobC.SelStart = 0
    txtParentJobC.SelLength = Len(txtParentJobC.Text)
End Sub

Private Sub txtParentJobC_LostFocus()
Dim bChkErr As Boolean
    If Trim(txtParentJobC) <> "" Then
        bChkErr = True
        For l_int_Counter = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
            If Trim(Me.m_ComboDetails.Item(l_int_Counter).GroupValue) = Trim(txtParentJobC) And UCase(Trim(Me.m_ComboDetails.Item(l_int_Counter).GroupId)) = UCase(cboParentJobC.Tag) Then
                If Me.m_ComboDetails.Item(l_int_Counter).Description <> "" Then
                    cboParentJobC.Text = Me.m_ComboDetails.Item(l_int_Counter).Description
                    bChkErr = False
                    Exit For
                End If
            End If
        Next
        If bChkErr Then
            txtParentJobC.Text = ""
            cboParentJobC.ListIndex = 0
        End If
    Else
        cboParentJobC.ListIndex = 0
    End If
End Sub

Private Sub txtSex_Change()
    SetChange
    Call txtSex_GotFocus
End Sub

Private Sub txtSex_GotFocus()
    txtSex.SelStart = 0
    txtSex.SelLength = Len(txtSex.Text)
End Sub

Private Sub txtSuisenFlag_Change()
    SetChange
    Call txtSuisenFlag_GotFocus
End Sub

Private Sub txtSuisenFlag_GotFocus()
    txtSuisenFlag.SelStart = 0
    txtSuisenFlag.SelLength = Len(txtSuisenFlag.Text)
End Sub

Private Sub txtSuisenFlag_LostFocus()
Dim bChkErr As Boolean
    If Trim(txtSuisenFlag) <> "" Then
        bChkErr = True
        For l_int_Counter = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
            If Trim(Me.m_ComboDetails.Item(l_int_Counter).Value) = Trim(txtSuisenFlag) And UCase(Trim(Me.m_ComboDetails.Item(l_int_Counter).GroupId)) = UCase(cboSuisenFlag.Tag) Then
                If Me.m_ComboDetails.Item(l_int_Counter).Description <> "" Then
                    cboSuisenFlag.Text = Me.m_ComboDetails.Item(l_int_Counter).Description
                    bChkErr = False
                    Exit For
                End If
            End If
        Next
        If bChkErr Then
            txtSuisenFlag.Text = ""
            cboSuisenFlag.ListIndex = 0
        End If
    Else
        cboSuisenFlag.ListIndex = 0
    End If
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

Private Sub txtSex_LostFocus()
    If txtSex.Text <> "M" And txtSex.Text <> "F" Then
        txtSex.Text = ""
        cboSex.ListIndex = -1
    Else
        If txtSex.Text = "M" Then
            cboSex.ListIndex = 1
        Else
            cboSex.ListIndex = 2
        End If
    End If
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

Private Sub txtSex_KeyPress(KeyAscii As Integer)

Dim sChkChar As String

    If Not (KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab) Then
        sChkChar = UCase(Chr(KeyAscii))
        If sChkChar = "F" Or sChkChar = "M" Then
            KeyAscii = Asc(sChkChar)
        Else
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtUnivName_Change()
    SetChange
End Sub

Private Sub txtZipCodeId_Change()
    SetChange
End Sub

Private Sub f_void_AddUnivType()
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iLookUpTableId,vName,iValue from tbSTELookUpTable WHERE iLookUpTableType = 1"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboUniversityType.AddItem ""
    
    While Not l_obj_Rst.EOF
        cboUniversityType.AddItem l_obj_Rst.Fields("vName")
        cboUniversityType.ItemData(cboUniversityType.NewIndex) = l_obj_Rst.Fields("iLookUpTableId")
        m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboUniversityType.Tag, "", l_obj_Rst.Fields("iValue")
        l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddCourse()
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iCourse,vCourseName from tbSTECourse order by iCourse "
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If

    ' add a blank row to the combobox
    cmbCourse.AddItem ""
    
    While Not l_obj_Rst.EOF
       cmbCourse.AddItem l_obj_Rst.Fields("vCourseName")
       cmbCourse.ItemData(cmbCourse.NewIndex) = l_obj_Rst.Fields("iCourse")
       m_ComboDetails.Add l_obj_Rst.Fields("iCourse"), l_obj_Rst.Fields("vCourseName"), cmbCourse.Tag
       l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_AddDepartment()
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iDepartment,vDepartmentName from tbSTEDepartment order by iDepartment "
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cmbDepartment.AddItem ""
    
    While Not l_obj_Rst.EOF
       cmbDepartment.AddItem l_obj_Rst.Fields("vDepartmentName")
       cmbDepartment.ItemData(cmbDepartment.NewIndex) = l_obj_Rst.Fields("iDepartment")
       m_ComboDetails.Add l_obj_Rst.Fields("iDepartment"), l_obj_Rst.Fields("vDepartmentName"), cmbDepartment.Tag
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
    
    l_str_Sql = "Select iLookUpTableId,vName,iValue from tbSTELookUpTable WHERE iLookUpTableType =2"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
        
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
         
    ' add a blank row to the combobox
    cbobackGroundId.AddItem ""
    
    While Not l_obj_Rst.EOF
        cbobackGroundId.AddItem l_obj_Rst.Fields("vName")
        cbobackGroundId.ItemData(cbobackGroundId.NewIndex) = l_obj_Rst.Fields("iLookUpTableId")
        m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cbobackGroundId.Tag, "", l_obj_Rst.Fields("iValue")
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
    
    l_str_Sql = "Select iLookUpTableId,vName,iValue from tbSTELookUpTable WHERE iLookUpTableType =3"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboFamilyID.AddItem ""

    While Not l_obj_Rst.EOF
        cboFamilyID.AddItem l_obj_Rst.Fields("vName")
        cboFamilyID.ItemData(cboFamilyID.NewIndex) = l_obj_Rst.Fields("iLookUpTableId")
        m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboFamilyID.Tag, "", l_obj_Rst.Fields("iValue")
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
    
    l_str_Sql = "Select iLookUpTableId,vName,iValue from tbSTELookUpTable WHERE iLookUpTableType =4"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboParentJobC.AddItem ""
    
    While Not l_obj_Rst.EOF
        cboParentJobC.AddItem l_obj_Rst.Fields("vName")
        cboParentJobC.ItemData(cboParentJobC.NewIndex) = l_obj_Rst.Fields("iLookUpTableId")
        m_ComboDetails.Add l_obj_Rst.Fields("iLookUpTableId"), l_obj_Rst.Fields("vName"), cboParentJobC.Tag, "", l_obj_Rst.Fields("iValue")
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
    
    l_str_Sql = "Select iLookUpTableId,vName,iValue from tbSTELookUpTable WHERE iLookUpTableType =5"
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
    
    l_str_Sql = "Select iLookUpTableId,vName,iValue from tbSTELookUpTable WHERE iLookUpTableType =6"
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
    
    l_str_Sql = "Select iSubjectProfileID,vSubjectName from tbSTESubjectProfile WHERE iSubType = 1 "
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboLanguageSubjProfileID.AddItem ""
    
    While Not l_obj_Rst.EOF
       cboLanguageSubjProfileID.AddItem l_obj_Rst.Fields("vSubjectName")
       m_ComboDetails.Add l_obj_Rst.Fields("iSubjectProfileID"), l_obj_Rst.Fields("vSubjectName"), cboLanguageSubjProfileID.Tag
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

    l_str_Sql = "Select iSubjectProfileID,vSubjectName from tbSTESubjectProfile WHERE iSubType = 2 "
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
        l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboScienceSubjprofileID1.AddItem ""
    
    While Not l_obj_Rst.EOF
        cboScienceSubjprofileID1.AddItem l_obj_Rst.Fields("vSubjectName")
        m_ComboDetails.Add l_obj_Rst.Fields("iSubjectProfileID"), l_obj_Rst.Fields("vSubjectName"), cboScienceSubjprofileID1.Tag
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

    l_str_Sql = "Select iSubjectProfileID,vSubjectName from tbSTESubjectProfile WHERE iSubType = 2 "
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Not l_obj_Rst.EOF Then
        l_obj_Rst.MoveFirst
    End If
    
    ' add a blank row to the combobox
    cboScienceSubjprofileID2.AddItem ""
    
    While Not l_obj_Rst.EOF
        cboScienceSubjprofileID2.AddItem l_obj_Rst.Fields("vSubjectName")
        m_ComboDetails.Add l_obj_Rst.Fields("iSubjectProfileID"), l_obj_Rst.Fields("vSubjectName"), cboScienceSubjprofileID2.Tag
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
    cboSex.ItemData(cboSex.NewIndex) = -1
    cboSex.AddItem LoadResString(1837)
    cboSex.ItemData(cboSex.NewIndex) = 0
    cboSex.AddItem LoadResString(1838)
    cboSex.ItemData(cboSex.NewIndex) = 1
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
Dim iAdmissionType As Long
Dim vAdmissionName As String
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
        iAdmissionType = l_obj_Rst.Fields("iAdmissionType")
        vAdmissionName = l_obj_Rst.Fields("vAdmissionName")
        cboAdmissionType1.AddItem vAdmissionName
        cboAdmissionType1.ItemData(cboAdmissionType1.NewIndex) = iAdmissionType
        cboAdmissionType2.AddItem vAdmissionName
        cboAdmissionType2.ItemData(cboAdmissionType1.NewIndex) = iAdmissionType
        m_ComboDetails.Add iAdmissionType, vAdmissionName, cboAdmissionType1.Tag
        m_ComboDetails.Add iAdmissionType, vAdmissionName, cboAdmissionType2.Tag
        l_obj_Rst.MoveNext
    Wend
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Public Function f_str_HighSchoolPref(psHighSchoolCode As String) As String

Dim sRtnStr As String

    Select Case Left(psHighSchoolCode, 2)
    Case "01": sRtnStr = "ñkäCìπ"
    Case "02": sRtnStr = "ê¬êX"
    Case "03": sRtnStr = "ä‚éË"
    Case "04": sRtnStr = "ã{èÈ"
    Case "05": sRtnStr = "èHìc"
    Case "06": sRtnStr = "éRå`"
    Case "07": sRtnStr = "ïüìá"
    Case "08": sRtnStr = "àÔèÈ"
    Case "09": sRtnStr = "ì»ñÿ"
    Case "10": sRtnStr = "åQîn"
    Case "11": sRtnStr = "çÈã "
    Case "12": sRtnStr = "êÁót"
    Case "13": sRtnStr = "ìåãû"
    Case "14": sRtnStr = "ê_ìﬁêÏ"
    Case "15": sRtnStr = "êVäÉ"
    Case "16": sRtnStr = "ïxéR"
    Case "17": sRtnStr = "êŒêÏ"
    Case "18": sRtnStr = "ïüà‰"
    Case "19": sRtnStr = "éRóú"
    Case "20": sRtnStr = "í∑ñÏ"
    Case "21": sRtnStr = "äÚïå"
    Case "22": sRtnStr = "ê√â™"
    Case "23": sRtnStr = "à§ím"
    Case "24": sRtnStr = "éOèd"
    Case "25": sRtnStr = "é†âÍ"
    Case "29": sRtnStr = "ìﬁó«"
    Case "26": sRtnStr = "ãûìs"
    Case "30": sRtnStr = "òaâÃéR"
    Case "27": sRtnStr = "ëÂç„"
    Case "28": sRtnStr = "ï∫å…"
    Case "31": sRtnStr = "íπéÊ"
    Case "32": sRtnStr = "ìáç™"
    Case "33": sRtnStr = "â™éR"
    Case "34": sRtnStr = "çLìá"
    Case "35": sRtnStr = "éRå˚"
    Case "36": sRtnStr = "ìøìá"
    Case "37": sRtnStr = "çÅêÏ"
    Case "38": sRtnStr = "à§ïQ"
    Case "39": sRtnStr = "çÇím"
    Case "40": sRtnStr = "ïüâ™"
    Case "41": sRtnStr = "ç≤âÍ"
    Case "42": sRtnStr = "í∑çË"
    Case "43": sRtnStr = "åFñ{"
    Case "44": sRtnStr = "ëÂï™"
    Case "45": sRtnStr = "ã{çË"
    Case "46": sRtnStr = "é≠éôìá"
    Case "47": sRtnStr = "â´ìÍ"
    Case Else: sRtnStr = "äCäOÅEÇªÇÃëº"
    End Select
    f_str_HighSchoolPref = sRtnStr

End Function

Public Function f_str_HighSchoolType(psHighSchoolCode As String) As String

Dim sRtnStr As String

    Select Case Mid(psHighSchoolCode, 1, 2)
    Case "51": sRtnStr = "çëóß"
    Case "52": sRtnStr = "äOçë"
    Case "53": sRtnStr = "éwíË"
    Case "54": sRtnStr = "îFíË"
    Case "55": sRtnStr = "ç›äO"
    Case Else
        Select Case Mid(psHighSchoolCode, 3, 1)
        Case "0": sRtnStr = "çëóß"
        Case "1", "2", "3", "4": sRtnStr = "åˆóß"
        Case Else: sRtnStr = "éÑóß"
        End Select
    End Select
    f_str_HighSchoolType = sRtnStr

End Function

Private Sub s_clickCmb2ChangTxt(poTrgCmb As Object, poTrgTxt As Object)
    poTrgTxt.Text = ""
    If poTrgCmb.ListIndex <> -1 Then
        For l_int_Counter = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
            If fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Value = poTrgCmb.ItemData(poTrgCmb.ListIndex) And UCase(Trim(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).GroupId)) = UCase(poTrgCmb.Tag) Then
                If Trim(str(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Value)) <> "" Then
                    poTrgTxt.Text = Trim(str(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Value))
                    Exit For
                End If
            End If
        Next
    Else
        poTrgTxt.Text = ""
    End If
End Sub

Private Sub s_clickCmb2ChangTxtGrp(poTrgCmb As Object, poTrgTxt As Object)
    poTrgTxt.Text = ""
    If poTrgCmb.ListIndex <> -1 Then
        For l_int_Counter = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
            If fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Value = poTrgCmb.ItemData(poTrgCmb.ListIndex) And UCase(Trim(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).GroupId)) = UCase(poTrgCmb.Tag) Then
                If Trim(str(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).GroupValue)) <> "" Then
                    poTrgTxt.Text = Trim(str(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).GroupValue))
                    Exit For
                End If
            End If
        Next
    Else
        poTrgTxt.Text = ""
    End If
End Sub

