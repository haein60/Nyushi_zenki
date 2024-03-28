VERSION 5.00
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   Caption         =   "2057"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   13485
   Tag             =   "1006"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClear 
      Caption         =   "2415"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   31
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "2414"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   30
      Tag             =   "2414"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox txtJukenNoFrom 
      Height          =   315
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   27
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtJukenNoTo 
      Height          =   315
      Left            =   7560
      MaxLength       =   4
      TabIndex        =   26
      Top             =   3480
      Width           =   1695
   End
   Begin VB.ComboBox cboPhysicalConditionId 
      Height          =   315
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2850
      Width           =   1695
   End
   Begin VB.ComboBox cboParentJobCategory 
      Height          =   315
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ComboBox cbobackGroundId 
      Height          =   315
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox cboLanguageSubjProfile 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ComboBox cboQualificationId 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ComboBox cboFamilyId 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Frame fraSuisen 
      Caption         =   "2057"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   5280
      TabIndex        =   9
      Tag             =   "2057"
      Top             =   4200
      Width           =   3975
      Begin VB.OptionButton optSuisen 
         Caption         =   "2060"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Tag             =   "2060"
         Top             =   480
         Width           =   2655
      End
      Begin VB.OptionButton optSuisen 
         Caption         =   "2061"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Tag             =   "2061"
         Top             =   960
         Width           =   2655
      End
   End
   Begin VB.Frame fraAbsentFlag 
      Caption         =   "1816"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   240
      TabIndex        =   8
      Tag             =   "1816"
      Top             =   4200
      Width           =   3735
      Begin VB.OptionButton optAbsentFlag 
         Caption         =   "2062"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   25
         Tag             =   "2062"
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optAbsentFlag 
         Caption         =   "2063"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   24
         Tag             =   "2063"
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.TextBox txtNendo 
      Height          =   315
      Left            =   7560
      MaxLength       =   4
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtNationality 
      Height          =   315
      Left            =   7560
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtExamineeStatus 
      Height          =   315
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtHighschoolCode 
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblJukenNoFrom 
      Caption         =   "1952"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   29
      Tag             =   "1952"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblJukenNumberTo 
      Caption         =   "1955"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5280
      TabIndex        =   28
      Tag             =   "1955"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblPhysicalConditionId 
      Caption         =   "1834"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5280
      TabIndex        =   23
      Tag             =   "1834"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblParentJobCategory 
      Caption         =   "1833"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Tag             =   "1833"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblBackGroundId 
      Caption         =   "1821"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5280
      TabIndex        =   21
      Tag             =   "1821"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblLanguageSubjProfileId 
      Caption         =   "1825"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Tag             =   "1825"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblQualificationId 
      Caption         =   "1823"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Tag             =   "1823"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label lblFamilyId 
      Caption         =   "1831"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Tag             =   "1831"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblNendo 
      Caption         =   "1804"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Tag             =   "1804"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblNationality 
      Caption         =   "1818"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Tag             =   "1818"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblExamineeStatus 
      Caption         =   "1817"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Tag             =   "1817"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblHighschoolCode 
      Caption         =   "1103"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Tag             =   "1103"
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmSearch
'Author         :   Vishal Kamath
'Created On     :
'Description    :   This form makes a provision for Searching the Favoured (Suisen)records.
'Reference      :   Functional Specs Of Maintain Examinee Data Ver 1.0
'**************************************************************************************************

Private Sub f_void_AddUnivType()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    l_str_sql = "SELECT iValue,vName FROM tbSTELookUpTable WHERE iLookUpTableType =1"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
         If Not l_obj_Rst.EOF Then
            l_obj_Rst.MoveFirst
         End If
         While Not l_obj_Rst.EOF
            cboUniversityType.AddItem l_obj_Rst.Fields("vName")
            l_obj_Rst.MoveNext
         Wend
     l_obj_Rst.Close
     Set l_obj_Rst = Nothing
End Sub

Private Sub cmdClear_Click()
'    lblErrorDetails.Caption = ""
    cbobackGroundId.ListIndex = -1
    cboFamilyId.ListIndex = -1
    cboLanguageSubjProfile.ListIndex = -1
    cboParentJobCategory.ListIndex = -1
    cboPhysicalConditionId.ListIndex = -1
    cboQualificationId.ListIndex = -1
    
    txtExamineeStatus.Text = ""
    txtHighschoolCode.Text = ""
    txtNendo.Text = ""
    txtNationality.Text = ""
    txtJukenNoFrom.Text = ""
    txtJukenNoTo.Text = ""
    
    optAbsentFlag(0).Value = False
    optAbsentFlag(1).Value = False
    optSuisen(0).Value = False
    optSuisen(1).Value = False
    
    
End Sub

Private Sub cmdOK_Click()
Dim l_str_sql As String
l_str_sql = ""
   
    
    If txtHighschoolCode.Text <> "" Then
        If l_str_sql <> "" Then
            l_str_sql = l_str_sql & " AND  iHighSchoolId = (SELECT iHighSchoolId FROM "
            l_str_sql = l_str_sql & " tbSTEHighSchoolType WHERE vHighSchoolCode = '" & Trim(txtHighschoolCode.Text) & "')"
        Else
        l_str_sql = l_str_sql & " iHighSchoolId = (SELECT iHighSchoolId FROM "
        l_str_sql = l_str_sql & " tbSTEHighSchoolType WHERE vHighSchoolCode = '" & Trim(txtHighschoolCode.Text) & "')"
        End If
    End If
    
    If txtExamineeStatus.Text <> "" Then
        If l_str_sql <> "" Then
            l_str_sql = l_str_sql & " AND  iExamineeStatus = " & Trim(txtExamineeStatus.Text)
        Else
            l_str_sql = l_str_sql & " iExamineeStatus = " & Trim(txtExamineeStatus.Text)
        End If
    End If
        
    If cbobackGroundId.Text <> "" Then
        If l_str_sql <> "" Then
            l_str_sql = l_str_sql & " AND  iBackGroundId = (SELECT iValue FROM tbSTELookUptable "
            l_str_sql = l_str_sql & " WHERE vname='" & Trim(cbobackGroundId.Text) & "'"
            l_str_sql = l_str_sql & " AND iLookUpTableType=2 )"
        Else
            l_str_sql = l_str_sql & " iBackGroundId = (SELECT iValue FROM tbSTELookUptable "
            l_str_sql = l_str_sql & " WHERE vname='" & Trim(cbobackGroundId.Text) & "'"
            l_str_sql = l_str_sql & " AND iLookUpTableType=2 )"
        End If
    End If
        
    If cboFamilyId.Text <> "" Then
        If l_str_sql <> "" Then
            l_str_sql = l_str_sql & " AND  iFamilyId = (SELECT iValue FROM tbSTELookUptable "
            l_str_sql = l_str_sql & " WHERE vname='" & Trim(cboFamilyId.Text) & "'"
            l_str_sql = l_str_sql & " AND iLookUpTableType=3 )"
        Else
            l_str_sql = l_str_sql & " iFamilyId = (SELECT iValue FROM tbSTELookUptable "
            l_str_sql = l_str_sql & " WHERE vname='" & Trim(cboFamilyId.Text) & "'"
            l_str_sql = l_str_sql & " AND iLookUpTableType=3 )"
        End If
    End If
        
    If cboParentJobCategory.Text <> "" Then
        If l_str_sql <> "" Then
            l_str_sql = l_str_sql & " AND  iFamilyId = (SELECT iValue FROM tbSTELookUptable "
            l_str_sql = l_str_sql & " WHERE vname='" & Trim(cboParentJobCategory.Text) & "'"
            l_str_sql = l_str_sql & " AND iLookUpTableType=4 )"
        Else
            l_str_sql = l_str_sql & " iFamilyId = (SELECT iValue FROM tbSTELookUptable "
            l_str_sql = l_str_sql & " WHERE vname='" & Trim(cboParentJobCategory.Text) & "'"
            l_str_sql = l_str_sql & " AND iLookUpTableType=4)"
        End If
    End If
        
    If cboQualificationId.Text <> "" Then
        If l_str_sql <> "" Then
            l_str_sql = l_str_sql & " AND  iFamilyId = (SELECT iValue FROM tbSTELookUptable "
            l_str_sql = l_str_sql & " WHERE vname='" & Trim(cboQualificationId.Text) & "'"
            l_str_sql = l_str_sql & " AND iLookUpTableType=5 )"
        Else
            l_str_sql = l_str_sql & " iFamilyId = (SELECT iValue FROM tbSTELookUptable "
            l_str_sql = l_str_sql & " WHERE vname='" & Trim(cboQualificationId.Text) & "'"
            l_str_sql = l_str_sql & " AND iLookUpTableType=5)"
        End If
    End If
        
        
    If cboPhysicalConditionId.Text <> "" Then
        If l_str_sql <> "" Then
            l_str_sql = l_str_sql & " AND  iFamilyId = (SELECT iValue FROM tbSTELookUptable "
            l_str_sql = l_str_sql & " WHERE vname='" & Trim(cboPhysicalConditionId.Text) & "'"
            l_str_sql = l_str_sql & " AND iLookUpTableType=6 )"
        Else
            l_str_sql = l_str_sql & " iFamilyId = (SELECT iValue FROM tbSTELookUptable "
            l_str_sql = l_str_sql & " WHERE vname='" & Trim(cboPhysicalConditionId.Text) & "'"
            l_str_sql = l_str_sql & " AND iLookUpTableType=6)"
        End If
    End If
        
    If txtNationality.Text <> "" Then
        If l_str_sql <> "" Then
            l_str_sql = l_str_sql & " AND  vNationality LIKE '%" & Trim(txtNationality.Text) & "%'"
        Else
            l_str_sql = l_str_sql & " vNationality LIKE '%" & Trim(txtNationality.Text) & "%'"
        End If
    End If
    
    If txtNendo.Text <> "" Then
        If l_str_sql <> "" Then
            l_str_sql = l_str_sql & " AND  iNendo =" & Trim(txtNendo.Text)
        Else
            l_str_sql = l_str_sql & " iNendo =" & Trim(txtNendo.Text)
        End If
    End If
    
    If cboLanguageSubjProfile.Text <> "" Then
        If l_str_sql <> "" Then
            l_str_sql = l_str_sql & " AND  iLanguageSubjProfileId = (SELECT iSubjectProfileId FROM tbSTESubjectProfile "
            l_str_sql = l_str_sql & " WHERE vSubjectName='" & Trim(cboLanguageSubjProfile.Text) & "'" & "AND iExamType = " & g_int_ExamType & ")"
        Else
            l_str_sql = l_str_sql & " iLanguageSubjProfileId = (SELECT iSubjectProfileId FROM tbSTESubjectProfile "
            l_str_sql = l_str_sql & " WHERE vSubjectName='" & Trim(cboLanguageSubjProfile.Text) & "'" & "AND iExamType = " & g_int_ExamType & ")"
            
        End If
    End If
       
    If optAbsentFlag(0).Value Then
            'Present
            If l_str_sql <> "" Then
                l_str_sql = l_str_sql & " AND iAbsentFlag =0"
            Else
                l_str_sql = l_str_sql & " iAbsentFlag =0"
            End If
    ElseIf optAbsentFlag(1).Value Then
            If l_str_sql <> "" Then
                l_str_sql = l_str_sql & " AND iAbsentFlag =1"
            Else
                l_str_sql = l_str_sql & " iAbsentFlag =1"
            End If
    End If
    
    If optSuisen(0).Value Then
            'Present
            If l_str_sql <> "" Then
                l_str_sql = l_str_sql & " AND iSuisenFlagId =0"
            Else
                l_str_sql = l_str_sql & " iSuisenFlagId =0"
            End If
    ElseIf optSuisen(1).Value Then
            If l_str_sql <> "" Then
                l_str_sql = l_str_sql & " AND iSuisenFlagId =1"
            Else
                l_str_sql = l_str_sql & " iSuisenFlagId =1"
            End If
    End If
    
    If txtJukenNoFrom.Text <> "" And txtJukenNoTo.Text <> "" Then
            If l_str_sql <> "" Then
                l_str_sql = l_str_sql & " AND iJukenNumber BETWEEN " & txtJukenNoFrom.Text & " AND " & txtJukenNoTo.Text
            Else
                l_str_sql = l_str_sql & "iJukenNumber BETWEEN " & txtJukenNoFrom.Text & " AND " & txtJukenNoTo.Text
            End If
    End If
    
    Dim l_bln_RetVal As Boolean
    l_bln_RetVal = frmExamineeProfile.f_bln_Search(l_str_sql)
    
    If l_bln_RetVal Then
        Me.Hide
        frmExamineeProfile.Show
    End If
End Sub



Private Sub Form_Activate()
    fMainForm.mnuTools.Enabled = False
    Dim index
    For index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(index).Enabled = False
    Next
'    lblErrorDetails.Caption = ""
End Sub

Private Sub Form_Load()
    LoadResStrings Me
    Me.Caption = LoadResString(2057)
    optSuisen(0).Caption = LoadResString(2060)
    optSuisen(1).Caption = LoadResString(2061)
    
    
    fMainForm.mnuTools.Enabled = False
    'Call f_void_AddUnivType
    Call f_void_AddBackgroundID
    Call f_void_AddFamilyID
    Call f_void_AddParentJobCategory
    Call f_void_AddQualificationID
    Call f_void_PhysicalConditionID
    Call f_void_AddLanguageSubjProfileID
End Sub

Private Sub f_void_AddBackgroundID()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_sql = "Select iValue,vName from tbSTELookUpTable WHERE iLookUpTableType = 2"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    ' add a blank row to the combobox
    cbobackGroundId.AddItem ""
    While Not l_obj_Rst.EOF
       cbobackGroundId.AddItem l_obj_Rst.Fields("vName")
       l_obj_Rst.MoveNext
    Wend
    cbobackGroundId.ListIndex = 0
     l_obj_Rst.Close
     Set l_obj_Rst = Nothing
     Exit Sub
     
ErrorHandler:
     MsgBox Err.Description
End Sub
Private Sub f_void_AddFamilyID()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_sql = "Select iValue,vName from tbSTELookUpTable WHERE iLookUpTableType = 3"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    ' add a blank row to the combobox
    cboFamilyId.AddItem ""
    While Not l_obj_Rst.EOF
       cboFamilyId.AddItem l_obj_Rst.Fields("vName")
       l_obj_Rst.MoveNext
    Wend
    cboFamilyId.ListIndex = 0
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
End Sub
Private Sub f_void_AddParentJobCategory()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_sql = "Select iValue,vName from tbSTELookUpTable WHERE iLookUpTableType = 4"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    ' add a blank row to the combobox
    cboParentJobCategory.AddItem ""
    While Not l_obj_Rst.EOF
       cboParentJobCategory.AddItem l_obj_Rst.Fields("vName")
       l_obj_Rst.MoveNext
    Wend
    cboParentJobCategory.ListIndex = 0
    
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub f_void_AddQualificationID()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_sql = "Select iValue,vName from tbSTELookUpTable WHERE iLookUpTableType = 5"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    ' add a blank row to the combobox
    cboQualificationId.AddItem ""
    While Not l_obj_Rst.EOF
       cboQualificationId.AddItem l_obj_Rst.Fields("vName")
       l_obj_Rst.MoveNext
    Wend
    cboQualificationId.ListIndex = 0
    
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub f_void_PhysicalConditionID()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_sql = "Select iValue,vName from tbSTELookUpTable WHERE iLookUpTableType = 6"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
    If Not l_obj_Rst.EOF Then
       l_obj_Rst.MoveFirst
    End If
    ' add a blank row to the combobox
    cboPhysicalConditionId.AddItem ""
    While Not l_obj_Rst.EOF
       cboPhysicalConditionId.AddItem l_obj_Rst.Fields("vName")
       l_obj_Rst.MoveNext
    Wend
    cboPhysicalConditionId.ListIndex = 0
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub f_void_AddLanguageSubjProfileID()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_sql = "Select iValue,vName from tbSTELookUpTable WHERE iLookUpTableType =7"
    
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
    If Not l_obj_Rst.EOF Then
        l_obj_Rst.MoveFirst
    End If
    ' add a blank row to the combobox
    cboLanguageSubjProfile.AddItem ""
   
    While Not l_obj_Rst.EOF
        cboLanguageSubjProfile.AddItem l_obj_Rst.Fields("vName")
        l_obj_Rst.MoveNext
    Wend
    cboLanguageSubjProfile.ListIndex = 0
    
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub txtExamineeStatus_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtJukenNoFrom_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
 End If
End Sub
Private Sub txtJukenNoTo_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
 End If
End Sub

Private Sub txtNendo_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
 End If
End Sub
