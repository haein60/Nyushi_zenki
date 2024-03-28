VERSION 5.00
Begin VB.Form frmSpecialInterview 
   Caption         =   "Form1"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmSpecialInterview.frx":0000
   ScaleHeight     =   9750
   ScaleWidth      =   10620
   WindowState     =   2  '最大化
   Begin VB.TextBox txtTotal 
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
      Left            =   9015
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   8040
      Width           =   1230
   End
   Begin VB.ComboBox cboDayValues 
      Height          =   315
      Left            =   8040
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cboRoomId 
      Height          =   315
      Left            =   8280
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox cboDestInterviewId 
      Height          =   315
      Left            =   3360
      TabIndex        =   15
      Text            =   "Combo2"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox cboSourceInterviewId 
      Height          =   315
      Left            =   3360
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CheckBox chkAbsentees 
      Caption         =   "Check1"
      Height          =   240
      Left            =   3045
      TabIndex        =   0
      Top             =   1170
      Width           =   240
   End
   Begin VB.ComboBox cboDestInterview 
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
      Left            =   3045
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   2280
      Width           =   2370
   End
   Begin VB.ComboBox cboSourceInterview 
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
      Left            =   3045
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1680
      Width           =   2370
   End
   Begin VB.CommandButton cmdDistributionLogic 
      Caption         =   "2462"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3000
      Width           =   4455
   End
   Begin VB.ComboBox cboSourceRoom 
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
      Left            =   7710
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      Top             =   1680
      Width           =   2505
   End
   Begin VB.ListBox lstSource 
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
      Height          =   5190
      Left            =   6105
      MultiSelect     =   2  '拡張
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   4140
   End
   Begin VB.ComboBox cboSourceDay 
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
      Left            =   7725
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   1080
      Width           =   2490
   End
   Begin VB.Label lblTotalDayRoom 
      Caption         =   "2475"
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
      Height          =   360
      Left            =   240
      TabIndex        =   20
      Top             =   8040
      Width           =   8625
   End
   Begin VB.Label lblErrorDetails 
      Caption         =   "Label6"
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
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   3600
      Width           =   5655
   End
   Begin VB.Label Label5 
      Caption         =   "2478"
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
      Height          =   465
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   2550
   End
   Begin VB.Label Label2 
      Caption         =   "2477"
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
      Height          =   465
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   2550
   End
   Begin VB.Label Label1 
      Caption         =   "2476"
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
      Height          =   465
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   2550
   End
   Begin VB.Label lblSourceCapacity 
      BackStyle       =   0  '透明
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
      Left            =   9810
      TabIndex        =   10
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "1505"
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
      Left            =   6420
      TabIndex        =   9
      Top             =   2280
      Width           =   3180
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "1503"
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
      Left            =   6360
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblSourceDay 
      BackStyle       =   0  '透明
      Caption         =   "1755"
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
      Left            =   6360
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmSpecialInterview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmSpecialInterview
'Author         :   Dileep Cherian
'Created On     :   16/05/02
'Description    :   This form cna be used to allocate examinees for any special interview that might happen
'Reference      :   Functional Specs Of Special Interview Ver 1.0
'***************************************************************************************************
Option Explicit
Dim f_dt_SourceDay As Date              ' to store the selected source day
Dim f_int_SourceDayMax As Integer       ' to store the max capacity of the selected source dayon day
Dim f_int_SourceRoomMax As Integer      ' max capacity of selected source room
Dim f_int_SourceRoomCount As Integer    ' count of existing examinees in selected source room
Dim f_bln_Flag As Boolean               ' flag to clear the error label
    
Private Sub cboDestInterview_Click()
    If Not f_bln_Flag Then lblErrorDetails.Caption = ""
    cboDestInterviewId.ListIndex = cboDestInterview.ListIndex
    f_bln_Flag = False
End Sub

Private Sub cboSourceDay_Click()
    On Error GoTo ErrorHandler
    cboDayValues.ListIndex = cboSourceDay.ListIndex
    Call l_void_PopulateRoomCombo(cboSourceDay.Text)
    txtTotal.Text = lstSource.ListCount
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cboSourceInterview_Click()
    ' populate the destination interview list based on the & _
        selection of source interview
    Dim l_str_sqlDestInterview As String
    Dim l_obj_rsDestInterview As New ADODB.Recordset
    Dim l_str_SubjectIds As String
    Dim l_int_Counter As Integer
    
    On Error GoTo ErrorHandler
    
    If Not f_bln_Flag Then lblErrorDetails.Caption = ""
    cboSourceInterviewId.ListIndex = cboSourceInterview.ListIndex
    
    l_str_SubjectIds = ""
    For l_int_Counter = 0 To cboSourceInterviewId.ListCount - 1
        l_str_SubjectIds = l_str_SubjectIds & cboSourceInterviewId.List(l_int_Counter) & ","
    Next
    
    If l_str_SubjectIds <> "" Then
        l_str_SubjectIds = Left(l_str_SubjectIds, Len(l_str_SubjectIds) - 1)
    End If
    
    l_str_sqlDestInterview = "SELECT iSubjectProfileId, vSubjectName FROM tbSTESubjectProfile" & _
        " WHERE iSubjectProfileId NOT IN(" & l_str_SubjectIds & ")" & _
        " AND iExamType IN(2,4)"
        
    l_obj_rsDestInterview.Open l_str_sqlDestInterview, g_obj_Conn
    
    cboDestInterview.Clear
    cboDestInterviewId.Clear
    Do While Not l_obj_rsDestInterview.EOF
        cboDestInterviewId.AddItem l_obj_rsDestInterview.Fields("iSubjectProfileId").Value
        cboDestInterview.AddItem l_obj_rsDestInterview.Fields("vSubjectName").Value
        l_obj_rsDestInterview.MoveNext
    Loop
    
    l_obj_rsDestInterview.Close
    Set l_obj_rsDestInterview = Nothing
    
    If cboDestInterview.ListCount > 0 Then
        cboDestInterview.ListIndex = 0
        cmdDistributionLogic.Enabled = True
        cmdDistributionLogic.Enabled = True
    Else
        cmdDistributionLogic.Enabled = False
        f_bln_Flag = False
    End If
    Call f_void_PopulateDayCombo
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cboSourceRoom_Click()
    ' display the maximum capacity of the selected room & _
        on the max capacity label
    Dim l_str_sqlRoomMax As String
    Dim l_obj_rsRoomMax As New ADODB.Recordset
    On Error GoTo ErrorHandler
    
    lblErrorDetails.Caption = ""
    cboRoomID.ListIndex = cboSourceRoom.ListIndex
    
    l_str_sqlRoomMax = "SELECT iMaxCapacity FROM tbSTERoomProfile" & _
        " WHERE vRoomName='" & cboSourceRoom.Text & "'"
    
    l_obj_rsRoomMax.Open l_str_sqlRoomMax, g_obj_Conn
    If Not l_obj_rsRoomMax.EOF Then
        f_int_SourceRoomMax = l_obj_rsRoomMax.Fields("iMaxCapacity").Value
    Else
        f_int_SourceRoomMax = 0
    End If
    l_obj_rsRoomMax.Close
    Set l_obj_rsRoomMax = Nothing
    
    lblSourceCapacity.Caption = CStr(f_int_SourceRoomMax)
    
    Call l_void_PopulateList(cboSourceRoom.Text, cboDayValues.Text)
    txtTotal.Text = lstSource.ListCount
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub cmdDistributionLogic_Click()
    Dim l_str_sqlExaminee As String
    Dim l_obj_rsExaminee As New ADODB.Recordset
    Dim l_str_ExamineeList As String
    Dim l_str_ExamineeArray() As String
    Dim l_int_CurrentDestIntwPos As Integer
    Dim l_int_counter1 As Integer
    Dim l_int_counter2 As Integer
    Dim l_str_sqlEaxmineeRoom As String
    Dim l_obj_rsExamineeRoom As New ADODB.Recordset
    Dim l_str_sqlDelete As String           ' delete statement
    Dim l_str_sqlInsert As String           ' insert statement
    Dim l_str_CheckExisting As String       ' to check whether data is existing for a subject
    Dim l_obj_rsCheck As New ADODB.Recordset ' to check whether data is existing for a subject
    Dim l_str_sqlTableIdMapping As String
    Dim l_obj_rsTableIdMapping As New ADODB.Recordset
    Dim l_int_ExamineeRoomProfileId As Long
    
    On Error GoTo ErrorHandler
    
    l_int_CurrentDestIntwPos = cboDestInterviewId.ListIndex
    
    If chkAbsentees.Value = Checked Then
        l_str_sqlExaminee = "SELECT iExamineeProfileId, iRoomProfileId FROM tbSTEExamineeRoomProfile" & _
            " WHERE iSubjectProfileId =" & cboSourceInterviewId.Text
    Else
        l_str_sqlExaminee = "SELECT iExamineeProfileId, iRoomProfileId" & _
            " FROM tbSTEExamineeRoomProfile WHERE iExamineeProfileId NOT IN(" & _
            " SELECT iExamineeProfileId FROM tbSTEScoreProfile" & _
            " WHERE iAbsentFlag=1" & _
            " AND iSubjectProfileId =" & cboSourceInterviewId.Text & ")" & _
            " AND iSubjectProfileId =" & cboSourceInterviewId.Text
    End If
    l_obj_rsExaminee.Open l_str_sqlExaminee, g_obj_Conn, adOpenStatic, adLockReadOnly
        
    ' see whether any record exists in the examineeroomprofile table
'    l_str_sqlEaxmineeRoom = "SELECT iExamineeRoomProfileId FROM tbSTEExamineeRoomProfile" & _
'        " ORDER BY iExamineeRoomProfileId"
'    l_obj_rsExamineeRoom.Open l_str_sqlEaxmineeRoom, g_obj_Conn, adOpenStatic, adLockReadOnly
'
'    If l_obj_rsExamineeRoom.EOF Then
'        ' if no record exists, pick the 1st id from the tableidmapping table
'        l_str_sqlTableIdMapping = " SELECT iTableCounterIdMapping FROM tbSTETableIdMapping" & _
'            " WHERE vTableName='tbSTEExamineeRoomProfile'"
'        l_obj_rsTableIdMapping.Open l_str_sqlTableIdMapping, g_obj_Conn
'        If l_obj_rsTableIdMapping.EOF Then
'            ' no id in tableidmapping table, so initialize to 1
'            l_int_ExamineeRoomProfileId = 1
'        Else
'            l_int_ExamineeRoomProfileId = l_obj_rsTableIdMapping.Fields("iTableCounterIdMapping").Value
'        End If
'        l_obj_rsTableIdMapping.Close
'        Set l_obj_rsTableIdMapping = Nothing
'    Else
'        ' get the new id by adding 1 to the highest existing id
'        l_obj_rsExamineeRoom.MoveLast
'        l_int_ExamineeRoomProfileId = l_obj_rsExamineeRoom.Fields("iExamineeRoomProfileId").Value + 1
'    End If
'
'    l_obj_rsExamineeRoom.Close
'    Set l_obj_rsExamineeRoom = Nothing
Dim bRtn As Boolean
    bRtn = getNewId("tbSTEExamineeRoomProfile", "iExamineeRoomProfileId", l_int_ExamineeRoomProfileId)

    g_obj_Conn.BeginTrans
        
    l_str_CheckExisting = "SELECT iSubjectProfileId FROM tbSTEExamineeRoomProfile" & _
        " WHERE iSubjectProfileId=" & cboDestInterviewId.Text
    l_obj_rsCheck.Open l_str_CheckExisting, g_obj_Conn
    If Not l_obj_rsCheck.EOF Then
        l_str_sqlDelete = "DELETE FROM tbSTEExamineeRoomProfile" & _
            " WHERE iSubjectProfileId =" & cboDestInterviewId.Text
        g_obj_Conn.Execute l_str_sqlDelete
    End If
    l_obj_rsCheck.Close
    Set l_obj_rsCheck = Nothing
    
    l_obj_rsExaminee.MoveFirst
    Do While Not l_obj_rsExaminee.EOF
        l_str_sqlInsert = "INSERT INTO tbSTEExamineeRoomProfile VALUES(" & _
            l_int_ExamineeRoomProfileId & "," & _
            l_obj_rsExaminee.Fields("iExamineeProfileId").Value & "," & _
            l_obj_rsExaminee.Fields("iRoomProfileId").Value & "," & _
            cboDestInterviewId.Text & ",'" & _
            Format(Date, "MM/DD/YYYY") & "','" & _
            Format(Date, "MM/DD/YYYY") & "')"
        g_obj_Conn.Execute l_str_sqlInsert
        
        l_int_ExamineeRoomProfileId = l_int_ExamineeRoomProfileId + 1
        l_obj_rsExaminee.MoveNext
    Loop
            
    g_obj_Conn.CommitTrans
    lblErrorDetails.Caption = "データを保存しました。" '''''LoadResString(1121)
    f_bln_Flag = True
    l_obj_rsExaminee.Close
    Set l_obj_rsExaminee = Nothing
    
    Call l_void_PopulateSourceInterviews
    Exit Sub
ErrorHandler:
On Error GoTo ErrProc

Dim sErrMsg As String

    sErrMsg = Err.Description

    g_obj_Conn.RollbackTrans

On Error GoTo 0

ErrProc:

    lblErrorDetails.Caption = LoadResString(1957)
    MsgBox sErrMsg, vbInformation, LoadResString(1729)
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
    On Error GoTo ErrorHandler
    
    LoadResStrings Me
    Me.Caption = LoadResString(2432)
    g_void_SetFontProperties Me
    ' code added to display bigger error messages
    lblErrorDetails.Height = 2000
    lblErrorDetails.WordWrap = True
    txtTotal.Text = lstSource.ListCount
    cboSourceInterviewId.Visible = False
    cboDestInterviewId.Visible = False
    cboRoomID.Visible = False
    cboDayValues.Visible = False
    cmdDistributionLogic.Enabled = False
    lblErrorDetails.Caption = ""
    Call l_void_PopulateSourceInterviews
    
    lblTotalDayRoom.Caption = LoadResString(2488)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub l_void_PopulateSourceInterviews()
    ' populate the source room interview list
    Dim l_str_sqlSourceInterview As String
    Dim l_obj_rsSourceInterview As New ADODB.Recordset
    
    l_str_sqlSourceInterview = "SELECT DISTINCT a.iSubjectProfileId, a.vSubjectName" & _
        " FROM tbSTESubjectProfile a, tbSTEExamineeRoomProfile b" & _
        " WHERE a.iSubjectProfileId = b.iSubjectProfileId" & _
        " AND a.iExamType IN(2,4)"
    l_obj_rsSourceInterview.Open l_str_sqlSourceInterview, g_obj_Conn
    
    cboSourceInterview.Clear
    cboSourceInterviewId.Clear
    Do While Not l_obj_rsSourceInterview.EOF
        cboSourceInterview.AddItem l_obj_rsSourceInterview.Fields("vSubjectName").Value
        cboSourceInterviewId.AddItem l_obj_rsSourceInterview.Fields("iSubjectProfileId").Value
        l_obj_rsSourceInterview.MoveNext
    Loop
    
    l_obj_rsSourceInterview.Close
    Set l_obj_rsSourceInterview = Nothing
    
    If cboSourceInterview.ListCount > 0 Then
        lblErrorDetails.Caption = ""
        cboSourceInterview.ListIndex = 0
    Else
        lblErrorDetails.Caption = LoadResString(2498)
        Exit Sub
    End If
    Call f_void_PopulateDayCombo
End Sub

Private Sub f_void_PopulateDayCombo()
    Dim l_str_SqlDay As String
    Dim l_obj_rsDay As New ADODB.Recordset
    
    ' get the current selected day and room, and their capacities
    l_str_SqlDay = "SELECT dtSecondExamDay1, dtSecondExamDay2, dtSecondExamDay3" & _
        " FROM tbSTESecondExamProfile" & _
        " WHERE iSystemProfileId = (" & _
        " SELECT iSystemProfileId FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag = 1)"
    
    l_obj_rsDay.Open l_str_SqlDay, g_obj_Conn
    If Not l_obj_rsDay.EOF Then
        ' populate the day combo
        With cboSourceDay
            .Clear
            cboDayValues.Clear
            .AddItem LoadResString(2424)    ' Day1
            cboDayValues.AddItem l_obj_rsDay.Fields("dtSecondExamDay1").Value    ' Day1
            .AddItem LoadResString(2425)    ' Day2
            cboDayValues.AddItem l_obj_rsDay.Fields("dtSecondExamDay2").Value    ' Day2
            If Not IsNull(l_obj_rsDay.Fields("dtSecondExamDay3").Value) Then
                .AddItem LoadResString(2426)    ' Day3
                cboDayValues.AddItem l_obj_rsDay.Fields("dtSecondExamDay3").Value    ' Day3
            End If
            .ListIndex = 0
        End With
    End If
            
    l_obj_rsDay.Close
    Set l_obj_rsDay = Nothing
End Sub

Private Sub l_void_PopulateRoomCombo(ByVal l_str_vDay As String)
    ' fill the room combo based on the day selected in the day combo
    Dim l_obj_RsRoom As New ADODB.Recordset    ' recordset object
    Dim l_str_sqlRoom As String                 ' SQL string
    Dim l_int_NoOfRooms As Integer          ' to store the number of rooms
    Dim l_int_Counter As Integer            ' counter
        
    cboSourceRoom.Clear
    
    ' get the current selected day and room, and their capacities
    l_str_sqlRoom = "SELECT iNumberOfRoomDay1, iNumberOfRoomDay2, iNumberOfRoomDay3," & _
        " dtSecondExamDay1, dtSecondExamDay2, dtSecondExamDay3," & _
        " iNumberOfExamineeDay1, iNumberOfExamineeDay2, iNumberOfExamineeDay3" & _
        " FROM tbSTESecondExamProfile" & _
        " WHERE iSystemProfileId = (" & _
        " SELECT iSystemProfileId FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag = 1)"
    
    l_obj_RsRoom.Open l_str_sqlRoom, g_obj_Conn
    If Not l_obj_RsRoom.EOF Then
        Label4.Visible = True
        lblSourceCapacity.Visible = True
        Select Case UCase(l_str_vDay)
        Case UCase(LoadResString(2424)) ' day 1
            f_dt_SourceDay = l_obj_RsRoom.Fields("dtSecondExamDay1").Value
            f_int_SourceDayMax = l_obj_RsRoom.Fields("iNumberOfExamineeDay1").Value
            l_int_NoOfRooms = l_obj_RsRoom.Fields("iNumberOfRoomDay1").Value
        Case UCase(LoadResString(2425)) ' day 2
            f_dt_SourceDay = l_obj_RsRoom.Fields("dtSecondExamDay2").Value
            f_int_SourceDayMax = l_obj_RsRoom.Fields("iNumberOfExamineeDay2").Value
            l_int_NoOfRooms = l_obj_RsRoom.Fields("iNumberOfRoomDay2").Value
        Case UCase(LoadResString(2426)) ' day 3
            f_dt_SourceDay = l_obj_RsRoom.Fields("dtSecondExamDay3").Value
            f_int_SourceDayMax = l_obj_RsRoom.Fields("iNumberOfExamineeDay3").Value
            l_int_NoOfRooms = l_obj_RsRoom.Fields("iNumberOfRoomDay3").Value
        End Select
    Else
        Label4.Visible = False
        lblSourceCapacity.Visible = False
        l_obj_RsRoom.Close
        Set l_obj_RsRoom = Nothing
        Exit Sub
    End If
            
    l_obj_RsRoom.Close
    Set l_obj_RsRoom = Nothing
    
    ' to check whether the max capacity of the room is reached or not
    ' change made on 31/07/02
    l_str_sqlRoom = "SELECT iRoomProfileId, vRoomName FROM tbSTERoomProfile" & _
        " WHERE iInterviewRoomFlag = 0" & _
        " ORDER BY iRoomProfileId"
    l_obj_RsRoom.Open l_str_sqlRoom, g_obj_Conn
    
    l_int_Counter = 1
    Do While Not l_obj_RsRoom.EOF And l_int_Counter <= l_int_NoOfRooms
        cboRoomID.AddItem l_obj_RsRoom.Fields("iRoomProfileId").Value
        cboSourceRoom.AddItem l_obj_RsRoom.Fields("vRoomName").Value
        l_int_Counter = l_int_Counter + 1
        l_obj_RsRoom.MoveNext
    Loop
    If cboSourceRoom.ListCount > 0 Then cboSourceRoom.ListIndex = 0
    l_obj_RsRoom.Close
    Set l_obj_RsRoom = Nothing
    
    ' to check whether the max capacity of the day is reached or not
    l_str_sqlRoom = "SELECT r.iExamineeProfileId FROM tbSTEExamineeRoomProfile r, tbSTEExamineeProfile e"
    Select Case l_str_vDay
    Case LoadResString(2424)
        l_str_sqlRoom = l_str_sqlRoom & " WHERE CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay1,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"
    Case LoadResString(2425)
        l_str_sqlRoom = l_str_sqlRoom & " WHERE CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay2,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"
    Case LoadResString(2426)
        l_str_sqlRoom = l_str_sqlRoom & " WHERE CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay3,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"
    End Select
        
    l_str_sqlRoom = l_str_sqlRoom & "  AND e.iExamineeProfileId = r.iExamineeProfileId"
    
    l_obj_RsRoom.Open l_str_sqlRoom, g_obj_Conn, adOpenStatic, adLockReadOnly
        
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub l_void_PopulateList(ByVal l_str_vRoom As String, ByVal l_dt_dtDay As Date)
    ' populate the list box based on selection made in the day and room combos
    Dim l_obj_Rst As New ADODB.Recordset        ' recordset object
    Dim l_obj_rst1 As New ADODB.Recordset       ' recordset object
    Dim l_str_Sql As String                     ' SQL string
    
    On Error GoTo ErrorHandler

    lstSource.Clear
    f_int_SourceRoomCount = 0
    
    l_str_Sql = l_str_Sql & " SELECT iExamineeProfileId FROM tbSTEExamineeRoomProfile" & _
        " WHERE iRoomProfileId=(" & _
        " SELECT iRoomProfileId FROM tbSTERoomProfile" & _
        " WHERE vRoomName='" & l_str_vRoom & "')" & _
        " AND iSubjectProfileId = " & cboSourceInterviewId.Text
    
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    
    Do While Not l_obj_Rst.EOF
        l_str_Sql = "SELECT iJukenNumber, vExamineeName, iPreferenceDay1Flag, iPreferenceDay2Flag, iPreferenceDay3Flag" & _
            " FROM tbSTEExamineeProfile" & _
            " WHERE iExamineeProfileId=" & l_obj_Rst("iExamineeProfileId") & _
            " AND dtSecondExamDay='" & Format(l_dt_dtDay, "MM/DD/YYYY") & "'" & _
            " AND iExamineeStatus = " & gclExamineeStatus_1stPass & _
            " AND iNendo=" & g_int_CurrentNendo
        l_obj_rst1.Open l_str_Sql, g_obj_Conn
        If Not l_obj_rst1.EOF Then
            lstSource.AddItem g_str_LPad(l_obj_rst1.Fields("iJukenNumber").Value, Len(l_obj_rst1.Fields("iJukenNumber").Value)) & _
                " - " & l_obj_rst1("vExamineeName") & _
                " - " & l_obj_rst1.Fields("iPreferenceDay1Flag").Value & _
                " - " & l_obj_rst1.Fields("iPreferenceDay2Flag").Value & _
                " - " & l_obj_rst1.Fields("iPreferenceDay3Flag").Value
            f_int_SourceRoomCount = f_int_SourceRoomCount + 1
        End If
        l_obj_rst1.Close
        Set l_obj_rst1 = Nothing
        
        l_obj_Rst.MoveNext
    Loop
    
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    
    l_str_Sql = "SELECT iMaxCapacity FROM tbSTERoomProfile WHERE vRoomName='" & _
        cboSourceRoom.Text & "'"
    
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    If Not l_obj_Rst.EOF Then
        f_int_SourceRoomMax = l_obj_Rst("ImaxCapacity")
    Else
        f_int_SourceRoomMax = 0
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call g_void_CloseChildForm
End Sub
