VERSION 5.00
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Nyushi"
   ClientHeight    =   8535
   ClientLeft      =   2280
   ClientTop       =   1740
   ClientWidth     =   11145
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin DXSIDEBARLibCtl.dxSideBar dxSideBar1 
      Align           =   3  'Align Left
      Height          =   8535
      Left            =   0
      OleObjectBlob   =   "frmMDI.frx":0000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuApplyPhase 
         Caption         =   "Apply Phase"
         Begin VB.Menu mnuOCR 
            Caption         =   "Read From OCR Data"
         End
         Begin VB.Menu mnuMaintainExamineeData 
            Caption         =   "Maintain Examinee Data"
         End
         Begin VB.Menu mnuHyotei 
            Caption         =   "Hyotei"
         End
         Begin VB.Menu mnuSuisen 
            Caption         =   "Suisen"
         End
         Begin VB.Menu mnuFixDataApply 
            Caption         =   "Fix Data"
         End
      End
      Begin VB.Menu mnu1stExam 
         Caption         =   "1st Exam"
         Begin VB.Menu mnuRoomAllocation 
            Caption         =   "Room Allocation"
         End
         Begin VB.Menu mnuInputAbsenteeRecord 
            Caption         =   "Input Absentee Record"
         End
         Begin VB.Menu mnuInputRawScore 
            Caption         =   "Input Raw Score"
         End
         Begin VB.Menu mnuInputChoseiScore 
            Caption         =   "Input Choosei Score"
         End
         Begin VB.Menu mnuInputPassedPersonData 
            Caption         =   "Input Passed Person Data"
         End
         Begin VB.Menu mnuDistribution 
            Caption         =   "Distribution Of Passed Examinee to groups"
         End
      End
      Begin VB.Menu mnuEnterRefuse 
         Caption         =   "Enter/&Refuse"
      End
      Begin VB.Menu mnu2ndExam 
         Caption         =   "2nd Exam"
         Begin VB.Menu mnuInputAbsenteeRecord2 
            Caption         =   "Input Absentee Record"
         End
         Begin VB.Menu mnuInputRawScore2 
            Caption         =   "Input Raw Score"
         End
         Begin VB.Menu mnuInputChoseiScore2 
            Caption         =   "Input Choosei Score"
         End
         Begin VB.Menu mnuInputPassedPersonData2 
            Caption         =   "Input Passed Person Data"
         End
         Begin VB.Menu mnuWaitList2 
            Caption         =   "Input Waiting List"
         End
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master Maintenance"
      Begin VB.Menu mnuHighSchoolType 
         Caption         =   "High School Type"
      End
      Begin VB.Menu mnuZipCode 
         Caption         =   "Zip Code"
      End
      Begin VB.Menu mnuRoomProfile 
         Caption         =   "Room Profile"
      End
      Begin VB.Menu mnuSubjectProfile 
         Caption         =   "Subject Profile"
      End
      Begin VB.Menu mnuInterviewerProfile 
         Caption         =   "Interviewer Profile"
      End
      Begin VB.Menu mnuInterviewRoomProfile 
         Caption         =   "Interview Room Profile"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuToolsSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuToolsDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuToolsCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuToolsClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuToolsNew 
         Caption         =   "New"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    ' to check the status of database connection
    Dim l_bln_Conn As Boolean
    
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    
    ' to store the current phase
    Dim l_int_CurrentPhase As Integer
    
    ' open database connection
    l_bln_Conn = g_void_OpenConnection()
    
    If Not l_bln_Conn Then
        ' there is an error in opening the database connection, exit the precedure
        MsgBox "There is an error opening the connection. Please try after sometime.", vbCritical, "Connection Error"
        Call mnuExit_Click
    End If
    
    l_str_Sql = "Select iCurrentPhase from tbSTESystemProfile where iActiveFlag=1"
    l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
        Exit Sub
    End If
    
    If Not l_obj_Rst.EOF Then
         l_int_CurrentPhase = l_obj_Rst("iCurrentPhase")
    End If
    
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    
    Select Case l_int_CurrentPhase
    
    Case 0
        mnuApplyPhase.Enabled = True
        mnu1stExam.Enabled = False
        mnu2ndExam.Enabled = False
        mnuEnterRefuse.Enabled = False
        
        g_int_ExamType = 0
        
    Case 1
        mnuApplyPhase.Enabled = False
        mnu1stExam.Enabled = True
        mnu2ndExam.Enabled = False
        mnuEnterRefuse.Enabled = False
        
        g_int_ExamType = 1
        
    Case 2
        mnuApplyPhase.Enabled = False
        mnu1stExam.Enabled = False
        mnu2ndExam.Enabled = True
        mnuEnterRefuse.Enabled = False
        
        g_int_ExamType = 2
        
    Case 3
        mnuApplyPhase.Enabled = False
        mnu1stExam.Enabled = False
        mnu2ndExam.Enabled = False
        mnuEnterRefuse.Enabled = True
    Case Else
        mnuApplyPhase.Enabled = False
        mnu1stExam.Enabled = False
        mnu2ndExam.Enabled = False
        mnuEnterRefuse.Enabled = False
    End Select
    
    '******** code from template ************
    LoadResStrings Me
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    LoadNewDoc
    frmRoomProfile.Show
    
    
End Sub

Private Sub mnuExit_Click()
    Dim l_frm As Form
    For Each l_frm In Forms
        Unload l_frm
    Next
    End
End Sub

Private Sub mnuFixDataApply_Click()
    Dim l_int_Response As Integer
    Dim l_frm As Form
    Dim l_str_Sql As String
    
    l_int_Response = MsgBox("This action will freeze the data for Apply phase and you will not be able to make futher changes to the data in this phase. Are you sure to proceed?", vbYesNo + vbQuestion, "Fix Data")
    If l_int_Response = vbYes Then
        l_str_Sql = "Update tbSTESystemProfile set"
        l_str_Sql = l_str_Sql & " iCurrentPhase=1 where iActiveFlag=1"

        g_obj_Conn.Execute (l_str_Sql)
        
        ' check error
        If Err.Number <> 0 Then
            MsgBox Err.Number & vbCrLf & Err.Description
            Exit Sub
        Else
            MsgBox "Data for Apply Phase is freezed. Enter 1st Exam Phase", vbInformation
            mnuApplyPhase.Enabled = False
            mnu1stExam.Enabled = True
            g_int_ExamType = 1
            For Each l_frm In Forms
                If l_frm.Name <> "frmLogo" And l_frm.Name <> "MDIForm1" Then
                    Unload l_frm
                End If
            Next
        End If
    End If
End Sub

Private Sub mnuHighSchoolType_Click()
    Load frmHighSchoolType
    frmHighSchoolType.Show
End Sub

Private Sub mnuHyotei_Click()
    Load frmRawScore
    frmRawScore.Show
End Sub

Private Sub mnuInputAbsenteeRecord_Click()
    Load frmAbsentRecord
    frmAbsentRecord.Show
End Sub

Private Sub mnuInputChoseiScore_Click()
    Load frmChooseiScore
    frmChooseiScore.Show
End Sub

Private Sub mnuInputPassedPersonData_Click()
    Load frmPassPersonData
    frmPassPersonData.Show
End Sub

Private Sub mnuInputPassedPersonData2_Click()
Load frmPassPersonData
frmPassPersonData.Show
End Sub

Private Sub mnuInputRawScore_Click()
    Load frmRawScore
    frmRawScore.Show
End Sub

Private Sub mnuInterviewerProfile_Click()
    Load frmInterviewerProfile
    frmInterviewerProfile.Show
End Sub

Private Sub mnuInterviewRoomProfile_Click()
    Load frmInterViewRoomProfile
    frmInterViewRoomProfile.Show
End Sub

Private Sub mnuMaintainExamineeData_Click()
    Load frmExamineeProfile
    frmExamineeProfile.Show
    frmExamineeProfile.ZOrder 0
End Sub

Private Sub mnuOCR_Click()
    Load frmBrowse
    frmBrowse.Show
End Sub

Private Sub mnuRoomAllocation_Click()
    Load frmRoomAlloc
    frmRoomAlloc.Show
End Sub

Private Sub mnuRoomProfile_Click()
    Load frmRoomProfile
    frmRoomProfile.Show
End Sub

Private Sub mnuSubjectProfile_Click()
    Load frmSubjectProfile
    frmSubjectProfile.Show
End Sub

Private Sub mnuSuisen_Click()
    Load frmSearch
    frmSearch.Show
End Sub

Private Sub mnuZipCode_Click()
    Load frmZipCode
    frmZipCode.Show
End Sub

'******** code from template ************
Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Document " & lDocumentCount
    frmD.Show
End Sub

'******** code from template ************
Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuToolsCancel_Click()
    'this menu will be enabled only for the data entry form- and dirty mode
    'this will be enabled only if save is enabled
    Call CancelData
End Sub

Private Sub mnuToolsClear_Click()
    'this menu will be enabled only for the data entry form- in short depends on the mode
    Call ClearData
End Sub

Private Sub mnuToolsDelete_Click()
    'this menu will be enabled only for the exiting data- in short depends on the mode
    Call DeleteData
End Sub

Private Sub mnuToolsDupicate_Click()
    'this menu is enabled only for the data entry form
    Call DuplicateData
End Sub

Private Sub mnuToolsNew_Click()
    'this menu will be enabled only for the data entry form
    Call NewData
End Sub

Private Sub mnuToolsSave_Click()
    'this menu is enabled only for the data entry form
    Call ValidateAndSaveData
End Sub

Private Sub mnuToolsSearch_Click()
    Call SearchRecords
End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

