VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmDistribution 
   Caption         =   "Distribution of Passed Examinee"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin VB.Frame fraDay 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8055
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   12735
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgDay 
         Height          =   3975
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7011
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSComctlLib.TabStrip tbsDistribution 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   15690
      TabWidthStyle   =   2
      TabFixedWidth   =   3528
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Day 1"
            Key             =   "day1"
            Object.Tag             =   "day1"
            Object.ToolTipText     =   "Day 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Day 2"
            Key             =   "day2"
            Object.Tag             =   "Day 2"
            Object.ToolTipText     =   "Day 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Day 3"
            Key             =   "day3"
            Object.Tag             =   "day3"
            Object.ToolTipText     =   "Day 3"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmDistribution
'Author         :   Dileep Cherian
'Created On     :
'Description    :   This screen is used to provide mechanism for distributing examinees,
'                   who have passed the first examination, for the interview and report
'                   in second examination.
'Reference      :   Functional Specs Of Distribution Of Examinee Ver 1.0
'**************************************************************************************************
Dim f_int_Day As Integer

Private Sub Form_Activate()
    fMainForm.mnuTools.Enabled = False
    Dim index
    For index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(index).Enabled = False
    Next
End Sub

Private Sub Form_Load()
    Dim l_obj_Cmd As New ADODB.Command
    Dim l_obj_Rst As New ADODB.Recordset
    Dim l_obj_rst1 As New ADODB.Recordset
    Dim l_obj_rst2 As New ADODB.Recordset
    Dim intCount As Integer
    
    On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    l_obj_Cmd.ActiveConnection = g_obj_Conn
    l_obj_Cmd.CommandType = adCmdStoredProc
    l_obj_Cmd.CommandText = "UspCTMAllocateRoom"
    l_obj_Cmd.Execute
    
    Screen.MousePointer = vbDefault
    
    f_int_Day = 1
    Call f_void_SizeTab
    Call Initialize_Grid
    Call Init_grid
    Call f_void_PopulateGrid
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub f_void_SizeTab()
    With tbsDistribution
        .Left = 0
        .Top = 1000
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
    
    With fraDay
        .Top = tbsDistribution.ClientTop
        .Left = tbsDistribution.ClientLeft
        .Width = tbsDistribution.ClientWidth
        .Height = tbsDistribution.ClientHeight
    End With
End Sub

Private Sub Form_Resize()
    Call f_void_SizeTab
End Sub

Private Sub Initialize_Grid()
    On Error GoTo ErrorHandler
    With hfgDay
        .Visible = False
        .BackColor = &HFFFFFF
        .BackColorBkg = &HFFFFFF
        .BackColorFixed = &H8000000F
        .BackColorSel = &H800000
        .FixedCols = 0
        .Font.Bold = False
        .Font.Name = "Verdana"
        .Font.Size = 8
        .ForeColorFixed = &H80000008
        .ForeColor = &H800000
        .CellTextStyle = "0"
        .GridLinesFixed = flexGridInset
        .GridColor = &H808080
        .Visible = True
    End With

    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub Init_grid()
    With hfgDay
        .Rows = 2
        .Cols = 3
        .FixedRows = 1
        .Row = 0
        .Col = 0
        .colwidth(0) = 800
        .Text = "Sr No"
        .CellAlignment = flexAlignRightBottom
        .Col = .Col + 1
        .colwidth(1) = 1500
        .Text = "Room Name"
        .CellAlignment = flexAlignLeftBottom
        .Col = .Col + 1
        .colwidth(2) = 5000
        .Text = "Juken Numbers"
        .CellAlignment = flexAlignLeftBottom
    End With
End Sub

Private Sub f_void_PopulateGrid()
    Dim l_obj_Rst As New ADODB.Recordset
    Dim l_obj_rst1 As New ADODB.Recordset
    Dim l_obj_rst2 As New ADODB.Recordset
    Dim l_str_sql As String
    Dim l_str_JukenNo As String
    Dim l_int_counter As Integer
    
    On Error GoTo ErrorHandler
    
    l_str_sql = "SELECT DISTINCT r.iRoomProfileId FROM tbSTEExamineeRoomProfile r, tbSTEExamineeProfile e"
    l_str_sql = l_str_sql & " WHERE e.iExamineeStatus = 1 AND e.IExamineeProfileId = r.iExamineeProfileId"
    l_str_sql = l_str_sql & " AND e.iAbsentFlag = 0 "
    l_str_sql = l_str_sql & " AND e.iNendo = (SELECT iNendo FROM tbSTESystemProfile"
    l_str_sql = l_str_sql & " WHERE iActiveFlag = 1)"
    Select Case f_int_Day
    Case 1
        l_str_sql = l_str_sql & " AND e.dtSecondExamDay = (SELECT dtSecondExamDay1 FROM tbSTESecondExamProfile"
        l_str_sql = l_str_sql & " WHERE iSystemProfileId = (SELECT iSystemProfileId FROM tbSTESystemProfile"
        l_str_sql = l_str_sql & " WHERE iActiveFlag = 1))"
    Case 2
        l_str_sql = l_str_sql & " AND e.dtSecondExamDay = (SELECT dtSecondExamDay2 FROM tbSTESecondExamProfile"
        l_str_sql = l_str_sql & " WHERE iSystemProfileId = (SELECT iSystemProfileId FROM tbSTESystemProfile"
        l_str_sql = l_str_sql & " WHERE iActiveFlag = 1))"
    Case 3
        l_str_sql = l_str_sql & " AND e.dtSecondExamDay = (SELECT dtSecondExamDay3 FROM tbSTESecondExamProfile"
        l_str_sql = l_str_sql & " WHERE iSystemProfileId = (SELECT iSystemProfileId FROM tbSTESystemProfile"
        l_str_sql = l_str_sql & " WHERE iActiveFlag = 1))"
    End Select
    
    l_obj_Rst.Open l_str_sql, g_obj_Conn, adOpenStatic, adLockReadOnly
'    MsgBox l_obj_Rst.RecordCount
    If Not l_obj_Rst.EOF Then
        Do While Not l_obj_Rst.EOF
            l_str_sql = "SELECT e.iJukenNumber FROM tbSTEExamineeProfile e, tbSTEExamineeRoomProfile r"
            l_str_sql = l_str_sql & " WHERE r.iRoomProfileId = " & l_obj_Rst("iRoomProfileId")
            l_str_sql = l_str_sql & " AND r.iExamineeProfileId = e.iExamineeProfileId"
            l_str_sql = l_str_sql & " AND e.iExamineeStatus = 1"
            l_str_sql = l_str_sql & " AND e.iAbsentFlag = 0"
'            l_str_sql = l_str_sql & " AND r.iRoomFlag = 0"
            
            Select Case f_int_Day
            Case 1
                l_str_sql = l_str_sql & " AND e.dtSecondExamDay = (SELECT dtSecondExamDay1 FROM tbSTESecondExamProfile"
                l_str_sql = l_str_sql & " WHERE iSystemProfileId = (SELECT iSystemProfileId FROM tbSTESystemProfile"
                l_str_sql = l_str_sql & " WHERE iActiveFlag = 1))"
            Case 2
                l_str_sql = l_str_sql & " AND e.dtSecondExamDay = (SELECT dtSecondExamDay2 FROM tbSTESecondExamProfile"
                l_str_sql = l_str_sql & " WHERE iSystemProfileId = (SELECT iSystemProfileId FROM tbSTESystemProfile"
                l_str_sql = l_str_sql & " WHERE iActiveFlag = 1))"
            Case 3
                l_str_sql = l_str_sql & " AND e.dtSecondExamDay = (SELECT dtSecondExamDay3 FROM tbSTESecondExamProfile"
                l_str_sql = l_str_sql & " WHERE iSystemProfileId = (SELECT iSystemProfileId FROM tbSTESystemProfile"
                l_str_sql = l_str_sql & " WHERE iActiveFlag = 1))"
            End Select
            
            l_obj_rst1.Open l_str_sql, g_obj_Conn, adOpenStatic, adLockReadOnly
            l_int_counter = l_int_counter + 1
            
            If Not l_obj_rst1.EOF Then
                l_str_JukenNo = ""
                Do While Not l_obj_rst1.EOF
                    l_str_JukenNo = l_str_JukenNo & l_obj_rst1("iJukenNumber") & ", "
                    l_obj_rst1.MoveNext
                Loop
            End If
            
            l_str_sql = "SELECT vRoomName FROM tbSTERoomProfile WHERE iRoomProfileId = " & l_obj_Rst("iRoomProfileId")
            l_obj_rst2.Open l_str_sql, g_obj_Conn, adOpenStatic, adLockReadOnly
            
            If Trim(l_str_JukenNo) <> "" Then
                l_str_JukenNo = Left(Trim(l_str_JukenNo), Len(Trim(l_str_JukenNo)) - 1)
            End If
            
            With hfgDay
                .Rows = .Rows + 1
                .Row = l_int_counter ' + 1
                .Col = 0
                .Text = CStr(l_int_counter)
                .Col = .Col + 1
                .Text = l_obj_rst2("vRoomName")
                .Col = .Col + 1
                .Text = l_str_JukenNo
                .CellAlignment = flexAlignLeftBottom
            End With
            l_obj_rst2.Close
            Set l_obj_rst2 = Nothing
            l_obj_rst1.Close
            Set l_obj_rst1 = Nothing
            
            l_obj_Rst.MoveNext
        Loop
        hfgDay.Rows = hfgDay.Rows - 1
    End If
    'hfgday.Rows = hfgday.Rows - 1
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub tbsDistribution_Click()
    f_int_Day = tbsDistribution.SelectedItem.index
    hfgDay.Rows = 2
    Call f_void_PopulateGrid
End Sub
