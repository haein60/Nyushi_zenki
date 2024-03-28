VERSION 5.00
Begin VB.Form frmReportChecker 
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmReportChecker.frx":0000
   ScaleHeight     =   10335
   ScaleWidth      =   13080
   WindowState     =   2  'Å‘å‰»
   Begin VB.ComboBox cboDay 
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
      Left            =   9000
      Style           =   2  'ÄÞÛ¯ÌßÀÞ³Ý Ø½Ä
      TabIndex        =   2
      Top             =   1080
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
      Left            =   5760
      Style           =   2  'ÄÞÛ¯ÌßÀÞ³Ý Ø½Ä
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ListBox lstAllInterviewers 
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
      Height          =   6300
      ItemData        =   "frmReportChecker.frx":3AD3
      Left            =   240
      List            =   "frmReportChecker.frx":3AD5
      MultiSelect     =   2  'Šg’£
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   3615
   End
   Begin VB.ListBox lstselectedInterviewers 
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
      Height          =   6300
      ItemData        =   "frmReportChecker.frx":3AD7
      Left            =   7080
      List            =   "frmReportChecker.frx":3AD9
      MultiSelect     =   2  'Šg’£
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton cmdSelectall 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4830
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4830
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4830
      TabIndex        =   8
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeselectall 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4830
      TabIndex        =   7
      Top             =   6600
      Width           =   1215
   End
   Begin VB.ComboBox cboRoomID 
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
      Left            =   4560
      Style           =   2  'ÄÞÛ¯ÌßÀÞ³Ý Ø½Ä
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cboRoom 
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
      Left            =   2160
      Style           =   2  'ÄÞÛ¯ÌßÀÞ³Ý Ø½Ä
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblErrorDetails 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   10455
   End
   Begin VB.Label lblDay 
      BackStyle       =   0  '“§–¾
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
      Left            =   7920
      TabIndex        =   14
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  '“§–¾
      Caption         =   "1954"
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
      Height          =   435
      Left            =   4440
      TabIndex        =   13
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblSelectedInterviewers 
      Caption         =   "2304"
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
      Left            =   7080
      TabIndex        =   12
      Top             =   2280
      Width           =   3585
   End
   Begin VB.Label lblAllInterviewers 
      Caption         =   "2303"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   3585
   End
   Begin VB.Label lblRoom 
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
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1785
   End
End
Attribute VB_Name = "frmReportChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmInterviewerRoom
'Author         :   Vishal Kamath
'Created On     :   30/10/01
'Description    :   This form makes a provision for mapping interviewers to rooms.
'Reference      :   FunctionalSpecs OFInterviewerRoommapping Ver1.1.doc
'**************************************************************************************************
'Modification History:
'12/5/02 Included Day ComboBox
'14/5/02 Insertion in TbsteSubjectQuestionProfile
'**************************************************************************************************
' Ammendments - NyushiChangesSummary.doc ver 1.0
'20/5/02 - Changed to add day combo and subsequent changes

Private f_obj_RsInterviewerID As New ADODB.Recordset            'Recordset object variable to open Interviewer table
Private f_str_InterviewerName As String                         'String variable to hold the SQL string
Private f_int_InterviewRoomProfileID As Integer                 'to store the interviewer profile id
Private f_bln_SelectAll As Boolean                              'Shows the status of the Select All button
Private f_bln_Select As Boolean                                 'Shows the status of the Select  button
Private f_bln_DeSelect As Boolean                               'Shows the status of the DeSelectAll button
Private f_bln_DeSelectAll As Boolean                            'Shows the status of the DeSelect  button
Private f_int_RoomProfileID As Integer                          'Stores the RoomProfileID
Public m_int_IntRpt As Integer                                  'form variable variable which indicated whether the form has to be instantiated for the "interview" or "report"

Private Sub cboDay_Click()
    Call f_void_InterviewerRooms
    Call f_void_AllInterviewers
    Call f_void_CheckButtonStatus
End Sub

'refresh the all interviewers and selected interviewers list on change of the subject
Private Sub cboSubject_Click()
    Call f_void_InterviewerRooms
    Call f_void_AllInterviewers
    Call f_void_CheckButtonStatus
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    fMainForm.mnuTools.Enabled = False          'disable the tools menu as this is not a master maintenance screen
    Dim index As Integer
    For index = 1 To fMainForm.Toolbar1.Buttons.Count   'disable all the toolbar button also
       fMainForm.Toolbar1.Buttons(index).Enabled = False
    Next
    If m_int_IntRpt = 0 Then
        'its for the interview
        Me.Caption = LoadResString(1051)
    Else
        ' its for the report
        Me.Caption = LoadResString(1053)
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    LoadResStrings Me               'load captions from resource file
    Call g_void_SetFontProperties(Me)     ' set the font properties
    Call f_void_LoadRoom            'populate room combo
    f_void_AllInterviewers          'populate list of all interviewers
    cmdDeselect.Enabled = False
    cmdSelect.Enabled = False
    Call f_void_PopulateSubject     'populate subject combo
    Call f_void_CheckButtonStatus
    Call l_void_PopulateDayCombo
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Public Sub f_void_LoadRoom()        'populate the room names
    Dim l_obj_RsRoom As New ADODB.Recordset
    Dim l_str_sqlRoom As String
    
    On Error GoTo ErrorHandler
    
    l_str_sqlRoom = "SELECT iRoomProfileid,vRoomName FROM tbSTERoomProfile" & _
        " WHERE iMaxCapacity > 0 "
    
    If m_int_IntRpt = 0 Then    ' change made on 31/07/02
        l_str_sqlRoom = l_str_sqlRoom & " AND iInterviewRoomFlag = 0"
    Else
        l_str_sqlRoom = l_str_sqlRoom & " AND iInterviewRoomFlag = 1"
    End If
    
    l_str_sqlRoom = l_str_sqlRoom & " ORDER BY iRoomProfileId"
    
    l_obj_RsRoom.Open l_str_sqlRoom, g_obj_Conn
    Do While Not l_obj_RsRoom.EOF
        cboRoomId.AddItem l_obj_RsRoom.Fields("iRoomProfileid").Value       'hidden combo to keep the id's of rooms
        cboRoom.AddItem l_obj_RsRoom.Fields("vRoomName").Value              'combo which displays the rooms names
        l_obj_RsRoom.MoveNext
    Loop
    
    If cboRoom.ListCount > 0 Then
        cboRoom.ListIndex = 0
        cboRoomId.ListIndex = 0
        lblErrorDetails.Caption = ""
    Else
        lblErrorDetails.Caption = LoadResString(2010)
        Unload Me
    End If
    l_obj_RsRoom.Close
    Set l_obj_RsRoom = Nothing
    Exit Sub
ErrorHandler:
        MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub
   
Private Sub cboRoom_Click()             'refresh the list boxes when a new room is selected
    cboRoomId.ListIndex = cboRoom.ListIndex
    lstselectedInterviewers.Clear
    f_int_RoomProfileID = cboRoomId.List(cboRoomId.ListIndex)
    f_void_InterviewerRooms
    f_int_RoomProfileID = cboRoomId.List(cboRoomId.ListIndex)
    Call f_void_AllInterviewers
    Call f_void_CheckButtonStatus
End Sub

Public Sub f_void_AllInterviewers()
    'loads all the Interviewers
    Dim l_str_Sql As String
    Dim l_obj_RsInterviewers  As ADODB.Recordset
    Dim l_int_Count As Integer
    Dim l_str_SelectedIntwr() As String
    
    On Error GoTo ErrorHandler
    lstAllInterviewers.Clear
    For l_int_Count = 0 To lstselectedInterviewers.ListCount - 1
        ReDim Preserve l_str_SelectedIntwr(l_int_Count)
        l_str_SelectedIntwr(l_int_Count) = Trim(lstselectedInterviewers.List(l_int_Count))
    Next
    
    Set l_obj_RsInterviewers = New ADODB.Recordset
    l_str_Sql = "SELECT iInterviewerProfileId,vInterviewerName FROM tbSTEInterviewerProfile"
    If l_int_Count > 0 Then ' select only those interviewers who are not assigned to any rooms for the selected subject
        l_str_Sql = l_str_Sql & " WHERE vInterviewerName NOT IN('" & Join(l_str_SelectedIntwr, "','")
        l_str_Sql = l_str_Sql & "')"
    End If
    l_obj_RsInterviewers.Open l_str_Sql, g_obj_Conn
    Do While Not l_obj_RsInterviewers.EOF
        lstAllInterviewers.AddItem l_obj_RsInterviewers.Fields("vInterviewerName").Value
        l_obj_RsInterviewers.MoveNext
    Loop
    l_obj_RsInterviewers.Close
    Set l_obj_RsInterviewers = Nothing
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Public Sub f_void_InterviewerRooms()
    'Loads the Interviewers already mapped to the selected rooms
    Dim l_str_Sql As String
    Dim l_obj_RsInterviewerRoom As ADODB.Recordset
    On Error GoTo ErrorHandler
    lstselectedInterviewers.Clear
    
    Set l_obj_RsInterviewerRoom = New ADODB.Recordset
    l_str_Sql = "SELECT i.iInterviewerProfileId,i.vInterviewerName,a.iRoomProfileId" & _
        " FROM tbSTEInterviewerprofile i,tbSTEInterviewRoomProfile a ,tbSTERoomProfile r" & _
        " WHERE i.iInterviewerProfileId = a.iInterviewerProfileId" & _
        " AND r.iRoomProfileId = a.iRoomProfileId" & _
        " AND a.iRoomProfileId = " & f_int_RoomProfileID & _
        " AND a.iDayFlag = " & cboDay.ListIndex & _
        " AND a.iSubjectProfileId =(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & cboSubject.Text & "')"
    
    If m_int_IntRpt = 0 Then    ' change made on 31/07/02
        l_str_Sql = l_str_Sql & " AND iInterviewRoomFlag = 0"
    Else
        l_str_Sql = l_str_Sql & " AND iInterviewRoomFlag = 1"
    End If
    
    Set l_obj_RsInterviewerRoom = g_obj_Conn.Execute(l_str_Sql)
    Do While Not l_obj_RsInterviewerRoom.EOF
        lstselectedInterviewers.AddItem Trim(l_obj_RsInterviewerRoom.Fields("vInterviewerName").Value)
        l_obj_RsInterviewerRoom.MoveNext
    Loop
    l_obj_RsInterviewerRoom.Close
    Set l_obj_RsInterviewerRoom = Nothing
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdSelectAll_Click()
    'On the click of this button all the Interviewers from the lstAllInterviewers will be transfered to lstSelectedInterviewers
    Dim l_bln_existing As Boolean           ' variable to check whether an interviewer is already existing the in the list box or not
    Dim l_int_Counter As Integer            ' counter variable
    Dim l_int_AllInterviewers As Integer    ' count of interviewers in the list box
    Dim l_str_Sql As String                 ' to store the SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    Dim l_int_IntwrId As Integer            ' to store the interviewer ID
    
    On Error GoTo ErrorHandler
    
    f_bln_SelectAll = True
    If lstAllInterviewers.ListCount >= 1 Then
        For l_int_AllInterviewers = 0 To lstAllInterviewers.ListCount - 1
            l_bln_existing = False
            For l_int_Counter = 0 To lstselectedInterviewers.ListCount - 1
             If Trim(lstselectedInterviewers.List(l_int_Counter)) = Trim(lstAllInterviewers.List(l_int_AllInterviewers)) Then
                l_bln_existing = True
                Exit For
             End If
             Next
            If Not l_bln_existing Then
                lstselectedInterviewers.AddItem lstAllInterviewers.List(l_int_AllInterviewers)
                lstAllInterviewers.ListIndex = l_int_AllInterviewers
                f_str_InterviewerName = lstAllInterviewers.Text
                l_str_Sql = "SELECT iInterviewerProfileId FROM tbSTEInterviewerProfile"
                l_str_Sql = l_str_Sql & " WHERE vInterviewerName='" & f_str_InterviewerName & "'"
                l_obj_Rst.Open l_str_Sql, g_obj_Conn
                If Not l_obj_Rst.EOF Then
                    l_int_IntwrId = l_obj_Rst("iInterviewerProfileId")
                End If
                l_obj_Rst.Close
                Set l_obj_Rst = Nothing
                f_void_UpdateDatabase (l_int_IntwrId)   ' update the database with the latest changes
            End If
        Next
    End If
    
    Call f_void_AllInterviewers     ' refresh the interviewers list
    f_void_CheckButtonStatus        ' enable/disable the direction buttons after the latest change
    f_bln_SelectAll = False
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub
Private Sub cmdSelect_Click()
    'on the click of this button only the Interviewer selected from the lstAllInterviewers will be transfered to
    'lstSelectedInterviewers
    Dim l_bln_existing As Boolean           ' variable to check whether an interviewer is already existing the in the list box or not
    Dim l_int_Counter As Integer            ' counter variable
    Dim l_int_Count As Integer              ' counter variable
    Dim l_int_IntwrId As Integer            ' to store interviewer ID
    Dim l_str_Sql As String                 ' to store the SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    
    On Error GoTo ErrorHandler
    
    f_bln_Select = True
    If lstAllInterviewers.SelCount > 0 Then
        For l_int_Count = lstAllInterviewers.ListCount - 1 To 0 Step -1
            If lstAllInterviewers.Selected(l_int_Count) Then
                For l_int_Counter = 0 To lstselectedInterviewers.ListCount - 1
                    If lstselectedInterviewers.List(l_int_Counter) = lstAllInterviewers.List(l_int_Count) Then
                        l_bln_existing = True
                        Exit For
                    End If
                Next
                
                If Not l_bln_existing Then
                    lstselectedInterviewers.AddItem lstAllInterviewers.List(l_int_Count)
                    lstAllInterviewers.ListIndex = l_int_Count
                    f_str_InterviewerName = lstAllInterviewers.Text
                    l_str_Sql = "SELECT iInterviewerProfileId FROM tbSTEInterviewerProfile" & _
                        " WHERE vInterviewerName='" & f_str_InterviewerName & "'"
                    l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
                    If Not l_obj_Rst.EOF Then
                        l_int_IntwrId = l_obj_Rst("iInterviewerProfileId")
                    End If
                    l_obj_Rst.Close
                    Set l_obj_Rst = Nothing
                    f_void_UpdateDatabase (l_int_IntwrId)
                End If
                lstAllInterviewers.ListIndex = l_int_Count
                f_str_InterviewerName = lstAllInterviewers.Text
            End If
        Next
    End If
    
    Call f_void_AllInterviewers
    f_bln_Select = False
    f_void_CheckButtonStatus
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdDeselect_Click()
    'on the click of this button only the interviewer selected from the lstSelectedInterviewers will be
    'transfered to lstAllInterviewers
    Dim l_int_Count As Integer              ' counter variable
    Dim l_int_IntwrId As Integer            ' to store the interviewer id
    Dim l_str_Sql As String                 ' to store the SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    
    On Error GoTo ErrorHandler
    
    f_bln_DeSelect = True
        If lstselectedInterviewers.SelCount > 0 Then
            For l_int_Count = lstselectedInterviewers.ListCount - 1 To 0 Step -1
                If lstselectedInterviewers.Selected(l_int_Count) Then
                    lstselectedInterviewers.ListIndex = l_int_Count
                    f_str_InterviewerName = lstselectedInterviewers.Text
                    l_str_Sql = "SELECT iInterviewerProfileId FROM tbSTEInterviewerProfile" & _
                        " WHERE vInterviewerName='" & f_str_InterviewerName & "'"
                    l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
                    If Not l_obj_Rst.EOF Then
                        l_int_IntwrId = l_obj_Rst("iInterviewerProfileId")
                    End If
                    l_obj_Rst.Close
                    Set l_obj_Rst = Nothing
                    f_void_UpdateDatabase (l_int_IntwrId)
                    lstselectedInterviewers.RemoveItem l_int_Count
                End If
            Next
        End If
    
    Call f_void_AllInterviewers
    f_bln_DeSelect = True
    f_void_CheckButtonStatus
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdDeselectAll_Click()
    'on the click of this button all the teachers from the lstSelectedTeachers will be removed from
    'the particular department
    Dim l_int_InterviewerCount As Integer   ' count of the interviewers
    Dim l_int_IntwrId As Integer            ' to store the interviewer id
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    Dim l_str_Sql As String                 ' to store the SQL string
    
    On Error GoTo ErrorHandler
    
    f_bln_DeSelectAll = True
    If lstselectedInterviewers.ListCount >= 1 Then
       For l_int_InterviewerCount = lstselectedInterviewers.ListCount - 1 To 0 Step -1
            lstselectedInterviewers.ListIndex = l_int_InterviewerCount
            f_str_InterviewerName = lstselectedInterviewers.Text
            l_str_Sql = "SELECT iInterviewerProfileId FROM tbSTEInterviewerProfile" & _
                " WHERE vInterviewerName='" & f_str_InterviewerName & "'"
            l_obj_Rst.Open l_str_Sql, g_obj_Conn
            If Not l_obj_Rst.EOF Then
                l_int_IntwrId = l_obj_Rst("iInterviewerProfileId")
            End If
            l_obj_Rst.Close
            Set l_obj_Rst = Nothing
            f_void_UpdateDatabase (l_int_IntwrId)
            lstselectedInterviewers.RemoveItem l_int_InterviewerCount
        Next
    End If
    
    Call f_void_AllInterviewers
    f_void_CheckButtonStatus
    f_bln_DeSelectAll = False
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Public Sub f_void_CheckButtonStatus()
    'Procedure to check the status of the buttons
    'i.e enabling and disabling the buttons based on the presense
    'and selection of data in the list boxes

    If lstAllInterviewers.ListCount = 0 Then
        ' left side list box is empty
        cmdSelectAll.Enabled = False
        cmdSelect.Enabled = False
    Else
        ' left side list box is not empty
        cmdSelectAll.Enabled = True
        If lstAllInterviewers.SelCount > 0 Then
            ' something is selected in the left list box
            cmdSelect.Enabled = True
        Else
            ' no item is selected in left list box
            cmdSelect.Enabled = False
        End If
    End If
    
    If lstselectedInterviewers.ListCount = 0 Then
        ' right side list box is empty
        cmdDeselectAll.Enabled = False
        cmdDeselect.Enabled = False
    Else
        ' right side list box is not empty
        cmdDeselectAll.Enabled = True
        If lstselectedInterviewers.SelCount > 0 Then
            ' something is selected in the right list box
            cmdDeselect.Enabled = True
        Else
            ' no item is selected in left right box
            cmdDeselect.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call g_void_CloseChildForm
End Sub

Private Sub lstAllInterviewers_Click()
    'Enables the cmdselect button when any element in the list box is selected else
    'button remains disabled
    f_void_CheckButtonStatus
End Sub

Private Sub lstAllInterviewers_DblClick()
    ' double clcicking on an item in the list box should have the same effect as selecting an item and clicking on the ">" or "<" button
    cmdSelect_Click
    f_void_CheckButtonStatus
End Sub

Private Sub lstselectedInterviewers_Click()
    'Enables the cmddeselect button when any element in the list box is selected else
    'button remains disabled
    f_void_CheckButtonStatus
End Sub

Public Sub f_void_UpdateDatabase(ByVal l_int_IntwrId As Integer)
    'Updating the database based on the status of the flags
    Dim l_str_Insert As String                          ' to hold the insert SQL string
    Dim l_str_Sql1 As String                            ' to hold the SQL string
    Dim l_obj_Rst As New ADODB.Recordset                ' recordset object
    Dim l_str_Sql As String
    Dim l_int_SubjectId As Integer                      ' to store the subject id for later use
    Dim l_str_Delete As String                          ' to store the delete SQL String
    Dim l_obj_RsdelInterviewer As New ADODB.Recordset   ' to store the id's of those interviewers to be deleted
    'New set of variables
    Dim l_str_subjectSelect As String  'Select
    Dim l_str_subjectInsert As String
    Dim l_obj_rstSubjectInsert As New ADODB.Recordset
    Dim l_obj_rstSubject As New ADODB.Recordset
    Dim l_int_SubjectQuestionProfileID As Integer
    Dim l_obj_subjectQuest As New ADODB.Recordset

    On Error GoTo ErrorHandler
    l_str_Sql = "SELECT iSubjectProfileId FROM tbSTESubjectProfile"
    l_str_Sql = l_str_Sql & " WHERE vSubjectName='" & cboSubject.Text & "'"
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    If Not l_obj_Rst.EOF Then
        l_int_SubjectId = l_obj_Rst("iSubjectProfileId")
    End If
    g_obj_Conn.BeginTrans

    If f_bln_SelectAll = True Or f_bln_Select = True Then
        
        Set f_obj_RsInterviewerID = g_obj_Conn.Execute("SELECT iInterviewRoomProfileId FROM tbSTEInterviewRoomProfile")
        
        If Not f_obj_RsInterviewerID.EOF Then
            f_obj_RsInterviewerID.MoveLast
            f_int_InterviewRoomProfileID = f_obj_RsInterviewerID.Fields(0).Value + 1
        Else
            l_str_Sql1 = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='tbSTEInterviewRoomProfile'"
            Set f_obj_RsInterviewerID = g_obj_Conn.Execute(l_str_Sql1)
            If Not f_obj_RsInterviewerID.EOF Then
                f_int_InterviewRoomProfileID = f_obj_RsInterviewerID("iTableCounterIdMapping")
            Else
                f_int_InterviewRoomProfileID = 1
            End If
            Set f_obj_RsInterviewerID = Nothing
        End If
        
        Set f_obj_RsInterviewerID = New ADODB.Recordset
        
        l_str_Insert = "Insert into tbSTEInterViewRoomProfile values(" & _
            f_int_InterviewRoomProfileID & "," & _
            l_int_IntwrId & "," & _
            f_int_RoomProfileID & "," & _
            l_int_SubjectId & "," & _
            cboDay.ListIndex & ",'" & _
            Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
            
        Set f_obj_RsInterviewerID = g_obj_Conn.Execute(l_str_Insert)
        Set f_obj_RsInterviewerID = Nothing
        
        '****************************** changes Mahesh start
       
        l_str_subjectSelect = "Select * from tbSteSubjectQuestionProfile"
        l_obj_rstSubject.Open l_str_subjectSelect, g_obj_Conn, adOpenStatic, adLockReadOnly
        
        If l_obj_rstSubject.RecordCount > 0 Then
            
            If Not l_obj_rstSubject.EOF Then
                l_obj_rstSubject.MoveLast
                l_int_SubjectQuestionProfileID = l_obj_rstSubject.Fields(0).Value + 1
            Else
                l_str_Sql1 = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='tbSteSubjectQuestionProfile'"
                Set l_obj_subjectQuest = g_obj_Conn.Execute(l_str_Sql1)
                If Not l_obj_subjectQuest.EOF Then
                    l_int_SubjectQuestionProfileID = l_obj_subjectQuest("iTableCounterIdMapping")
                Else
                    l_int_SubjectQuestionProfileID = 1
                End If
                Set l_obj_subjectQuest = Nothing
            End If
            Set l_obj_rstSubject = Nothing
            
            l_str_subjectSelect = "Select iInterviewerProfileId from tbSteSubjectQuestionProfile where iSubjectProfileId=" & l_int_SubjectId & " and iInterviewerProfileId =" & l_int_IntwrId
            l_obj_rstSubject.Open l_str_subjectSelect, g_obj_Conn, adOpenStatic, adLockReadOnly
            If l_obj_rstSubject.RecordCount = 0 Then
                l_str_subjectInsert = "Insert into tbSteSubjectQuestionProfile (iSubjectQuestionID,iSubjectProfileId,iInterviewerProfileId,dtCreate,dtUpdate)" _
                & " values(" & l_int_SubjectQuestionProfileID & "," & l_int_SubjectId & "," & l_int_IntwrId & "," & Format(Date, "MM/DD/YYYY") & "," & Format(Date, "MM/DD/YYYY") & ")"
                Set l_obj_rstSubjectInsert = g_obj_Conn.Execute(l_str_subjectInsert)
            End If
        End If

        '****************************** changes Mahesh End
        
    ElseIf f_bln_DeSelectAll = True Or f_bln_DeSelect = True Then
        'Changes Mahesh S
        Dim l_str_selSubject As String
        Dim l_obj_rstInterviewerRoom As New ADODB.Recordset
        
        'Changes Mahesh E
        l_str_Delete = "SELECT iInterviewRoomProfileId FROM tbSTEInterviewRoomProfile" & _
            " WHERE iRoomProfileId = " & f_int_RoomProfileID & _
            " AND iInterviewerProfileId = " & l_int_IntwrId & _
            " AND iSubjectProfileId = " & l_int_SubjectId & _
            " AND iDayFlag = " & cboDay.ListIndex

        Set l_obj_RsdelInterviewer = g_obj_Conn.Execute(l_str_Delete)
        f_int_InterviewRoomProfileID = l_obj_RsdelInterviewer.Fields("iInterviewRoomProfileId").Value
        Set l_obj_RsdelInterviewer = Nothing
        l_str_Delete = "DELETE from tbSTEInterviewRoomProfile where iInterviewRoomprofileId= " & f_int_InterviewRoomProfileID
        g_obj_Conn.Execute l_str_Delete
        
        'Changes Mahesh S
        l_str_selSubject = "Select * from tbSteInterviewRoomProfile where iSubjectProfileId=" & l_int_SubjectId & " and iInterviewerProfileId =" & l_int_IntwrId
        l_obj_rstInterviewerRoom.Open l_str_selSubject, g_obj_Conn, adOpenStatic, adLockReadOnly
        If l_obj_rstInterviewerRoom.RecordCount = 0 Then
            l_str_Delete = "Delete from tbSTESubjectQuestionProfile where iSubjectProfileId=" & l_int_SubjectId & " and iInterviewerProfileId =" & l_int_IntwrId
            g_obj_Conn.Execute l_str_Delete
        End If
        'Changes Mahesh S
    End If
    g_obj_Conn.CommitTrans
    Exit Sub
ErrorHandler:
    g_obj_Conn.RollbackTrans
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub lstselectedInterviewers_DblClick()
    cmdDeselect_Click
    f_void_CheckButtonStatus
End Sub

Private Sub f_void_PopulateSubject()
    ' populate the subject combo box
    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    
    On Error GoTo ErrorHandler
    l_str_Sql = "SELECT vSubjectName FROM tbSTESubjectProfile"
    If m_int_IntRpt = 0 Then
        l_str_Sql = l_str_Sql & " WHERE iExamType=2 or iExamType=4"
    Else
        l_str_Sql = l_str_Sql & " WHERE iExamType=3 or iExamType=5"
    End If
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    If Not l_obj_Rst.EOF Then
        Do While Not l_obj_Rst.EOF
            cboSubject.AddItem l_obj_Rst("vSubjectName")
            l_obj_Rst.MoveNext
        Loop
        cboSubject.ListIndex = 0
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub l_void_PopulateDayCombo()
    ' populate the day combo
    With cboDay
        .Clear
        .AddItem LoadResString(2424)
        .AddItem LoadResString(2425)
        .AddItem LoadResString(2426)
        .ListIndex = 0
    End With
End Sub

