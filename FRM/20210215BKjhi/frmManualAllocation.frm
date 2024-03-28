VERSION 5.00
Begin VB.Form frmManualAllocation 
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmManualAllocation.frx":0000
   ScaleHeight     =   9810
   ScaleWidth      =   12060
   Tag             =   "2431"
   WindowState     =   2  '最大化
   Begin VB.TextBox txtWemenDay 
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
      Left            =   10185
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   8280
      Width           =   1230
   End
   Begin VB.TextBox txtTotalExamineesDay 
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
      Left            =   10185
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   7800
      Width           =   1230
   End
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
      Left            =   10185
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   9120
      Width           =   1230
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
      Left            =   2640
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ListBox lstExaminee 
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
      Height          =   5730
      Left            =   240
      MultiSelect     =   2  '拡張
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   4575
   End
   Begin VB.ComboBox cboSplDay 
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
      Left            =   9150
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1080
      Width           =   2265
   End
   Begin VB.ListBox lstAllotted 
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
      Height          =   5730
      Left            =   6840
      MultiSelect     =   2  '拡張
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   4575
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "<"
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
      Left            =   5220
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "<<"
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
      Left            =   5220
      TabIndex        =   6
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   ">>"
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
      Left            =   5205
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   ">"
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
      Left            =   5205
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblWemenDay 
      Alignment       =   1  '右揃え
      Caption         =   "女性"
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
      TabIndex        =   18
      Top             =   8280
      Width           =   9825
   End
   Begin VB.Label lblErrorDetails 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   8760
      Width           =   11175
   End
   Begin VB.Label lblDayTotal 
      Alignment       =   1  '右揃え
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
      TabIndex        =   15
      Top             =   7800
      Width           =   9825
   End
   Begin VB.Label lblDayRoomTotal 
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
      TabIndex        =   13
      Top             =   9120
      Width           =   9825
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
      Left            =   11040
      TabIndex        =   11
      Top             =   3000
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
      Height          =   435
      Left            =   5160
      TabIndex        =   10
      Top             =   7200
      Width           =   3945
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  '透明
      Caption         =   "1403"
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
      TabIndex        =   9
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      Height          =   435
      Left            =   7005
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmManualAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmManualAllocation
'Author         :   Dileep Cherian
'Created On     :   16/05/02
'Description    :   This form is used for maual allocation of examiees for 2nd phase interviews
'Reference      :   Functional Specs Of Manual Allocation Ver 1.0
'***************************************************************************************************
Option Explicit
Dim f_dt_SplDay As Date                 ' to store the selected spl interview/report day
Dim f_int_SplDayMax As Long          ' to store the max capacity of the selected interview/report day
Dim f_int_SplDayCount As Long        ' counter to check the number of examinees allocated to the selected day
Dim f_int_SplRoomMax As Long         ' max capacity of the selected spl interview/report room
Dim f_int_SplRoomCount As Long       ' count of examinees allocated to the selected spl interview/report room
Dim f_bln_DataChange As Boolean         ' variable to indicate any change operations
Dim f_int_ExamType As Long           ' to identify the exam type
Dim f_str_RoomStatus As String          ' to store the room status before refreshing

Private Function lfCountWemen() As Integer

Dim iLoopCnt As Long
Dim iWemenCnt As Long

    iWemenCnt = 0

    For iLoopCnt = 0 To lstAllotted.ListCount - 1
        If InStr(lstAllotted.List(iLoopCnt), "(*)") = 0 Then
            iWemenCnt = iWemenCnt + 1
        End If
    Next

    lfCountWemen = iWemenCnt

End Function

Private Sub cmdDeselect_Click()
    'on the click of this button only the Interviewer selected from the lstExaminee
    'will be transfered to lstAllotted
    Dim l_int_Count As Long              ' counter
    Dim l_int_Start As Long              ' to extract the juken number from the combined string
    Dim l_int_End As Long                ' to extract the juken number from the combined string
    Dim l_int_JukenNo As Long            ' to store the extracted juken number
    Dim l_bln_Status As Boolean             ' to store the status of the function call
    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    Dim l_int_ExamineeId As Long         ' to store the examinee ID
    
    On Error GoTo ErrorHandler
    
    If lstAllotted.ListCount > 0 Then
        For l_int_Count = 0 To lstAllotted.ListCount - 1
            If l_int_Count > lstAllotted.ListCount - 1 Then Exit For
            If lstAllotted.Selected(l_int_Count) Then
                l_int_JukenNo = Left(lstAllotted.List(l_int_Count), 4)
                
                l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                    " WHERE iJukenNumber=" & l_int_JukenNo & _
                    " AND iNendo=" & g_int_CurrentNendo
                l_obj_Rst.Open l_str_Sql, g_obj_Conn
                If Not l_obj_Rst.EOF Then
                    l_int_ExamineeId = l_obj_Rst("iExamineeProfileId")
                End If
                l_obj_Rst.Close
                Set l_obj_Rst = Nothing
                
                l_bln_Status = f_bln_FreeExaminee(l_int_ExamineeId)
                If l_bln_Status Then
                    If Not f_bln_DataChange Then f_bln_DataChange = True
                    lstExaminee.AddItem lstAllotted.List(l_int_Count)
                    lstAllotted.RemoveItem (l_int_Count)
                    f_int_SplDayCount = f_int_SplDayCount - 1
                    l_int_Count = l_int_Count - 1   ' because an item is removed from the list
                Else
                    MsgBox LoadResString(2416)
                End If
            End If
        Next
    End If
    f_void_CheckButtonStatus
    txtTotal.Text = lstAllotted.ListCount
    ' refresh the room combo after an examinee is moved from one list bot to another
    Call l_void_PopulateRoomCombo(cboSplDay.Text)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdDeselectAll_Click()
    'On the click of this button all the Interviewers from the lstExaminee
    'will be transfered to lstAllotted
    Dim l_int_AllExaminee As Long        ' counter
    Dim l_int_Start As Long              ' to extract the juken number from the combined string
    Dim l_int_End As Long                ' to extract the juken number from the combined string
    Dim l_int_JukenNo As Long            ' to store the juken number
    Dim l_bln_Status As Boolean             ' to track the return value odf the function call
    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    Dim l_int_ExamineeId As Long         ' to store the examinee Id
    
    On Error GoTo ErrorHandler
        
    If lstAllotted.ListCount >= 1 Then
        For l_int_AllExaminee = 0 To lstAllotted.ListCount - 1
            If l_int_AllExaminee > lstAllotted.ListCount - 1 Then Exit For
            
            l_int_JukenNo = Left(lstAllotted.List(l_int_AllExaminee), 4)
                            
            l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                " WHERE iJukenNumber=" & l_int_JukenNo & _
                " AND iNendo=" & g_int_CurrentNendo
            l_obj_Rst.Open l_str_Sql, g_obj_Conn
            If Not l_obj_Rst.EOF Then
                l_int_ExamineeId = l_obj_Rst("iExamineeProfileId")
            End If
            l_obj_Rst.Close
            Set l_obj_Rst = Nothing
            
            l_bln_Status = f_bln_FreeExaminee(l_int_ExamineeId)
            If l_bln_Status Then
                If Not f_bln_DataChange Then f_bln_DataChange = True
                lstExaminee.AddItem lstAllotted.List(l_int_AllExaminee)
                lstAllotted.RemoveItem (l_int_AllExaminee)
                f_int_SplDayCount = f_int_SplDayCount - 1
                l_int_AllExaminee = l_int_AllExaminee - 1   ' because an item is removed from the list
            Else
                MsgBox LoadResString(2416)
            End If
        Next
    End If
    f_void_CheckButtonStatus
    txtTotal.Text = lstAllotted.ListCount
    ' refresh the room combo after an examinee is moved from one list bot to another
    Call l_void_PopulateRoomCombo(cboSplDay.Text)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdSelectAll_Click()
    'On the click of this button all the Interviewers from the lstExaminee
    'will be transfered to lstAllotted
    Dim l_bln_existing As Boolean           ' to see whether the examinee is already existing or not
    Dim l_int_Counter As Long            ' counter
    Dim l_int_AllExaminee As Long        ' counter
    Dim l_int_Start As Long              ' to extract the juken number from the combined string
    Dim l_int_End As Long                ' to extract the juken number from the combined string
    Dim l_int_JukenNo As Long            ' to store the juken number
    Dim l_bln_Flag As Boolean               ' to track the return value of the function call
    Dim l_bln_Status As Boolean             ' to track the return value of the function call
    Dim l_int_RetVal As Long             ' to track the return value of the function call
    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    Dim l_int_ExamineeId As Long         ' to store the examinee Id
    
    On Error GoTo ErrorHandler
        
    If lstExaminee.ListCount >= 1 Then
        For l_int_AllExaminee = 0 To lstExaminee.ListCount - 1
            If l_int_AllExaminee > lstExaminee.ListCount - 1 Then Exit For
            l_bln_existing = False
            For l_int_Counter = 0 To lstAllotted.ListCount - 1
                If Trim(lstAllotted.List(l_int_Counter)) = Trim(lstExaminee.List(l_int_AllExaminee)) Then
                   l_bln_existing = True
                   Exit For
                End If
            Next
            If Not l_bln_existing Then
                l_int_JukenNo = Left(lstExaminee.List(l_int_AllExaminee), 4)
                
                l_bln_Flag = f_void_CheckPreferenceViolation(l_int_JukenNo)
                If Not l_bln_Flag Then
                    l_int_RetVal = MsgBox(LoadResString(2417) & l_int_JukenNo & vbCrLf & LoadResString(2418) _
                                    , vbQuestion + vbYesNo, LoadResString(2423))
                    If l_int_RetVal = vbYes Then
                        If f_int_SplDayCount + 1 <= f_int_SplDayMax Then
                            If f_int_SplRoomCount + 1 <= f_int_SplRoomMax Then
                                l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                                    " WHERE iJukenNumber=" & l_int_JukenNo & _
                                    " AND iNendo=" & g_int_CurrentNendo
                                l_obj_Rst.Open l_str_Sql, g_obj_Conn
                                If Not l_obj_Rst.EOF Then
                                    l_int_ExamineeId = l_obj_Rst("iExamineeProfileId")
                                End If
                                l_obj_Rst.Close
                                Set l_obj_Rst = Nothing
                                
                                l_bln_Status = f_bln_UpdateDatabase(l_int_ExamineeId)
                                If l_bln_Status Then
                                    If Not f_bln_DataChange Then f_bln_DataChange = True
                                    lstAllotted.AddItem lstExaminee.List(l_int_AllExaminee)
                                    lstExaminee.RemoveItem (l_int_AllExaminee)
                                    f_int_SplRoomCount = f_int_SplRoomCount + 1
                                    f_int_SplDayCount = f_int_SplDayCount + 1
                                    l_int_AllExaminee = l_int_AllExaminee - 1   ' because an item is removed from the list'
                                End If
                            Else
                                MsgBox LoadResString(2419), vbCritical
                            End If
                        Else
                            MsgBox LoadResString(2461), vbCritical
                        End If
                    End If
                Else
                    If f_int_SplDayCount + 1 <= f_int_SplDayMax Then
                        If f_int_SplRoomCount + 1 <= f_int_SplRoomMax Then
                            l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                                " WHERE iJukenNumber=" & l_int_JukenNo & _
                                " AND iNendo=" & g_int_CurrentNendo
                            l_obj_Rst.Open l_str_Sql, g_obj_Conn
                            If Not l_obj_Rst.EOF Then
                                l_int_ExamineeId = l_obj_Rst("iExamineeProfileId")
                            End If
                            l_obj_Rst.Close
                            Set l_obj_Rst = Nothing
                            
                            l_bln_Status = f_bln_UpdateDatabase(l_int_ExamineeId)
                            If l_bln_Status Then
                                If Not f_bln_DataChange Then f_bln_DataChange = True
                                lstAllotted.AddItem lstExaminee.List(l_int_AllExaminee)
                                lstExaminee.RemoveItem (l_int_AllExaminee)
                                f_int_SplRoomCount = f_int_SplRoomCount + 1
                                f_int_SplDayCount = f_int_SplDayCount + 1
                                l_int_AllExaminee = l_int_AllExaminee - 1   ' because an item is removed from the list
                            End If
                        Else
                            MsgBox LoadResString(2419), vbCritical
                        End If
                    Else
                        MsgBox LoadResString(2461), vbCritical
                    End If
                End If
            End If
        Next
    End If
    f_void_CheckButtonStatus
    txtTotal.Text = lstAllotted.ListCount
    ' refresh the room combo after an examinee is moved from one list bot to another
    Call l_void_PopulateRoomCombo(cboSplDay.Text)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdSelect_Click()
    'on the click of this button only the Interviewer selected from the lstExaminee
    ' will be transfered to lstAllotted
    Dim l_bln_existing As Boolean           ' to see whether the examinee is already existing or not
    Dim l_int_Counter As Long            ' counter
    Dim l_int_Count As Long              ' counter
    Dim l_bln_Flag As Boolean               ' to see whether the examinee is already existing or not
    Dim l_int_Start As Long              ' to extract the juken number from the combined string
    Dim l_int_End As Long                ' to extract the juken number from the combined string
    Dim l_int_JukenNo As Long            ' to store the juken number
    Dim l_int_RetVal As Long             ' to track the return value of the function call
    Dim l_bln_Status As Boolean             ' to track the return value of the function call
    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    Dim l_int_ExamineeId As Long         ' to store the examinee Id
    
    On Error GoTo ErrorHandler
    
    If lstExaminee.SelCount > 0 Then
        For l_int_Count = 0 To lstExaminee.ListCount - 1
            If l_int_Count > lstExaminee.ListCount - 1 Then Exit For
            If lstExaminee.Selected(l_int_Count) Then
                For l_int_Counter = 0 To lstAllotted.ListCount - 1
                    If lstAllotted.List(l_int_Counter) = lstExaminee.List(l_int_Count) Then
                        l_bln_existing = True
                        Exit For
                    End If
                Next
                
                If Not l_bln_existing Then
                    l_int_JukenNo = Left(lstExaminee.List(l_int_Count), 4)
                    
                    l_bln_Flag = f_void_CheckPreferenceViolation(l_int_JukenNo)
                    If Not l_bln_Flag Then
                        l_int_RetVal = MsgBox(LoadResString(2417) & l_int_JukenNo & vbCrLf & LoadResString(2418) _
                                        , vbQuestion + vbYesNo, LoadResString(2423))
                        If l_int_RetVal = vbYes Then
                            If f_int_SplDayCount + 1 <= f_int_SplDayMax Then
'                                If f_int_SplRoomCount + 1 <= f_int_SplRoomMax Then
                                    l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                                        " WHERE iJukenNumber=" & l_int_JukenNo & _
                                        " AND iNendo=" & g_int_CurrentNendo
                                    l_obj_Rst.Open l_str_Sql, g_obj_Conn
                                    If Not l_obj_Rst.EOF Then
                                        l_int_ExamineeId = l_obj_Rst("iExamineeProfileId")
                                    End If
                                    l_obj_Rst.Close
                                    Set l_obj_Rst = Nothing
                                    
                                    l_bln_Status = f_bln_UpdateDatabase(l_int_ExamineeId)
                                    If l_bln_Status Then
                                        If Not f_bln_DataChange Then f_bln_DataChange = True
                                        lstAllotted.AddItem lstExaminee.List(l_int_Count)
                                        lstExaminee.RemoveItem (l_int_Count)
                                        f_int_SplRoomCount = f_int_SplRoomCount + 1
                                        f_int_SplDayCount = f_int_SplDayCount + 1
                                        l_int_Count = l_int_Count - 1   ' because an item is removed from the list
                                    End If
'                                Else
'                                    MsgBox LoadResString(2419), vbCritical
'                                End If
                            Else
                                MsgBox LoadResString(2461), vbCritical
                            End If
                        End If
                    Else
                        If f_int_SplDayCount + 1 <= f_int_SplDayMax Then
'                            If f_int_SplRoomCount + 1 <= f_int_SplRoomMax Then
                                l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                                    " WHERE iJukenNumber=" & l_int_JukenNo & _
                                    " AND iNendo=" & g_int_CurrentNendo
                                l_obj_Rst.Open l_str_Sql, g_obj_Conn
                                If Not l_obj_Rst.EOF Then
                                    l_int_ExamineeId = l_obj_Rst("iExamineeProfileId")
                                End If
                                l_obj_Rst.Close
                                Set l_obj_Rst = Nothing
                                
                                l_bln_Status = f_bln_UpdateDatabase(l_int_ExamineeId)
                                If l_bln_Status Then
                                    If Not f_bln_DataChange Then f_bln_DataChange = True
                                    lstAllotted.AddItem lstExaminee.List(l_int_Count)
                                    lstExaminee.RemoveItem (l_int_Count)
                                    f_int_SplRoomCount = f_int_SplRoomCount + 1
                                    f_int_SplDayCount = f_int_SplDayCount + 1
                                    l_int_Count = l_int_Count - 1   ' because an item is removed from the list
                                End If
'                            Else
'                                MsgBox LoadResString(2419), vbCritical
'                            End If
                        Else
                            MsgBox LoadResString(2461), vbCritical
                        End If
                    End If
                End If
            End If
        Next
    End If
    f_void_CheckButtonStatus
    txtTotal.Text = lstAllotted.ListCount
    ' refresh the room combo after an examinee is moved from one list bot to another
    Call l_void_PopulateRoomCombo(cboSplDay.Text)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Function f_bln_UpdateDatabase(ByVal iExamineeId As Long) As Boolean
    ' update the database with the current changes
    ' value has to be inserted in tbSTEExamineeRoomProfile
    ' also updation in tbSTEExamineeProfile
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    Dim l_obj_rst1 As New ADODB.Recordset
    Dim l_obj_rst2 As New ADODB.Recordset
    Dim l_int_NewId As Long
    Dim l_int_SubjectId As Long
    Dim l_int_RoomId As Long
    Dim l_int_ExamDate As Date
    Dim l_dt_IntvDate As Date
    Dim l_int_Counter As Long
    Dim l_int_LoopCounter As Long
    Dim l_str_SubjId() As String
    
    On Error GoTo ErrorHandler

    g_obj_Conn.BeginTrans

    l_str_Sql = "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & cboSubject.Text & "'"
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    If Not l_obj_Rst.EOF Then
        l_int_SubjectId = l_obj_Rst.Fields("iSubjectProfileId").Value
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing

    Select Case UCase(cboSplDay.Text)
    Case UCase(LoadResString(2424))
        l_str_Sql = "SELECT dtSecondExamDay1 FROM tbSTESecondExamProfile"
    Case UCase(LoadResString(2425))
        l_str_Sql = "SELECT dtSecondExamDay2 FROM tbSTESecondExamProfile"
    Case UCase(LoadResString(2426))
        l_str_Sql = "SELECT dtSecondExamDay3 FROM tbSTESecondExamProfile"
    End Select

    l_str_Sql = l_str_Sql & " WHERE iSystemProfileId=(SELECT iSystemProfileId" & _
        " FROM tbSTESystemProfile WHERE iActiveFlag=1)"
    l_obj_rst1.Open l_str_Sql, g_obj_Conn
    If Not l_obj_rst1.EOF Then
        If IsNull(l_obj_rst1(0)) Then
            g_obj_Conn.RollbackTrans
            MsgBox LoadResString(2416), vbInformation, "試験日に３日目がありません。"
            f_bln_UpdateDatabase = False
        End If
        l_int_ExamDate = l_obj_rst1(0)
    End If
    l_obj_rst1.Close
    Set l_obj_rst1 = Nothing
    
    l_str_Sql = "SELECT dtSecondExamDay FROM tbSTEExamineeProfile" & _
        " WHERE iExamineeProfileId=" & iExamineeId
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    If Not l_obj_Rst.EOF Then
        If Not IsNull(l_obj_Rst("dtSecondExamDay")) Then
            l_dt_IntvDate = l_obj_Rst("dtSecondExamDay")
        Else
            l_dt_IntvDate = Format(#1/1/1900#, "MM/DD/YYYY")
        End If
    End If
    ' update dtSecondExamDay field is tbSTEExamineeProfile table
'    If IsNull(l_dt_IntvDate) Or l_dt_IntvDate = Format(#1/1/1900#, "MM/DD/YYYY") Or l_dt_IntvDate = l_int_ExamDate Then
        l_str_Sql = "UPDATE tbSTEExamineeProfile SET dtSecondExamDay='" & Format(l_int_ExamDate, "MM/DD/YYYY") & "'," & _
            " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
            " WHERE iExamineeProfileId=" & iExamineeId
        g_obj_Conn.Execute l_str_Sql
'    Else
'        MsgBox LoadResString(2420) & l_dt_IntvDate & _
'        LoadResString(2421), vbCritical, LoadResString(2422)
'        g_obj_Conn.RollbackTrans
'        f_bln_UpdateDatabase = False
'        Exit Function
'    End If
    g_obj_Conn.CommitTrans
    f_bln_UpdateDatabase = True
    Exit Function
ErrorHandler:
    g_obj_Conn.RollbackTrans
    MsgBox LoadResString(2416), vbInformation, LoadResString(1729)
    f_bln_UpdateDatabase = False
End Function


Private Function f_void_CheckPreferenceViolation(ByVal l_int_iJukenNo As Integer) As Boolean
    ' check whteher the movement of selected examinee from unallocated list box & _
    to allocated listbox will cause in violation in the preference day mentioned by & _
    the examinee at the time of registration
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    
    On Error GoTo ErrorHandler
    l_str_Sql = "SELECT iPreferenceDay1Flag,iPreferenceDay2Flag,iPreferenceDay3Flag FROM tbSTEExamineeProfile" & _
        " WHERE iJukenNumber=" & l_int_iJukenNo & _
        " AND iNendo=" & g_int_CurrentNendo
    
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    If Not l_obj_Rst.EOF Then
        Select Case UCase(cboSplDay.Text)
        Case UCase(LoadResString(2424))
            If l_obj_Rst("iPreferenceDay1Flag") = 1 Then
                f_void_CheckPreferenceViolation = True
            Else
                f_void_CheckPreferenceViolation = False
            End If
        Case UCase(LoadResString(2425))
            If l_obj_Rst("iPreferenceDay2Flag") = 1 Then
                f_void_CheckPreferenceViolation = True
            Else
                f_void_CheckPreferenceViolation = False
            End If
        Case UCase(LoadResString(2426))
            If l_obj_Rst("iPreferenceDay3Flag") = 1 Then
                f_void_CheckPreferenceViolation = True
            Else
                f_void_CheckPreferenceViolation = False
            End If
        End Select
    Else
        f_void_CheckPreferenceViolation = False
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
    f_void_CheckPreferenceViolation = False
End Function

Private Sub l_void_PopulateDayCombo()
    ' populate the day combo box
Dim sSQL As String
Dim oRs As ADODB.Recordset
Dim bThirdDay As Boolean

    bThirdDay = False

    sSQL = "SELECT dtSecondExamDay3 FROM tbSTESecondExamProfile "
    sSQL = sSQL & " WHERE iSystemProfileId = "
    sSQL = sSQL & " (SELECT iSystemProfileId FROM tbSTESystemProfile WHERE iActiveFlag=1) "

    Set oRs = g_obj_Conn.Execute(sSQL)

    If Not oRs.EOF Then
        If Not IsNull(oRs.Fields(0)) Then
            bThirdDay = True
        End If
    End If

    oRs.Close
    Set oRs = Nothing

    With cboSplDay
        .Clear
        .AddItem LoadResString(2424)
        .AddItem LoadResString(2425)
        If bThirdDay Then .AddItem LoadResString(2426)
        .ListIndex = 0
    End With
End Sub

Private Sub l_void_PopulateList(ByVal l_dt_dtDay As Date)
    ' populate the list box based on selection made in the day and room combos
    Dim l_obj_Rst As New ADODB.Recordset        ' recordset object
    Dim l_obj_rsExaminee As New ADODB.Recordset       ' recordset object
    Dim l_str_Sql As String                     ' SQL string
    
    On Error GoTo ErrorHandler
    
    lstAllotted.Clear
    f_int_SplRoomCount = 0

    l_str_Sql = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber, substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName ,iPreferenceDay1Flag,iPreferenceDay2Flag,iPreferenceDay3Flag,iSex" & _
        " FROM tbSTEExamineeProfile" & _
        " WHERE dtSecondExamDay='" & Format(l_dt_dtDay, "MM/DD/YYYY") & "'" & _
        " AND iExamineeStatus = " & gclExamineeStatus_1stPass & _
        " AND iNendo=" & g_int_CurrentNendo

    l_obj_rsExaminee.Open l_str_Sql, g_obj_Conn

    Do While Not l_obj_rsExaminee.EOF
        lstAllotted.AddItem l_obj_rsExaminee.Fields("iJukenNumber").Value & _
        " - " & l_obj_rsExaminee.Fields("vExamineeName").Value & _
        " -" & l_obj_rsExaminee.Fields("iPreferenceDay1Flag").Value & _
        "-" & l_obj_rsExaminee.Fields("iPreferenceDay2Flag").Value & _
        "-" & l_obj_rsExaminee.Fields("iPreferenceDay3Flag").Value & _
        "-" & IIf(l_obj_rsExaminee.Fields("iSex") = 0, "(*)", "")
        f_int_SplRoomCount = f_int_SplRoomCount + 1
        l_obj_rsExaminee.MoveNext
    Loop
    l_obj_rsExaminee.Close
    Set l_obj_rsExaminee = Nothing

    Call f_void_PopulateExaminee
    txtTotal.Text = lstAllotted.ListCount
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_PopulateExaminee()
    ' pick up all the unallocated examinees and populate them in the & _
    unllocated (left) List box
    Dim l_int_Count As Long              ' counter
    Dim l_str_Arr() As String               ' to store the examinee id of all examinees
    Dim l_int_Start As Long              ' to extract the juken number
    Dim l_int_End As Long                ' to extract the juken number
    Dim l_int_JukenNo As Long            ' to store the examinee number
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    Dim l_str_Sql As String                 ' SQL string
    
    On Error GoTo ErrorHandler

    l_str_Sql = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber, substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName,iPreferenceDay1Flag,iPreferenceDay2Flag,iPreferenceDay3Flag , iSex" & _
        " FROM tbSTEExamineeProfile" & _
        " WHERE dtSecondExamDay<>'" & Format(f_dt_SplDay, "MM/DD/YYYY") & "'" & _
        " AND iExamineeStatus = " & gclExamineeStatus_1stPass & _
        " AND iNendo=" & g_int_CurrentNendo

    With l_obj_Rst
        .Open l_str_Sql, g_obj_Conn
        lstExaminee.Clear
        
        Do While Not .EOF
            lstExaminee.AddItem l_obj_Rst.Fields("iJukenNumber").Value & _
            " - " & l_obj_Rst.Fields("vExamineeName").Value & _
            " -" & l_obj_Rst.Fields("iPreferenceDay1Flag").Value & _
            "-" & l_obj_Rst.Fields("iPreferenceDay2Flag").Value & _
            "-" & l_obj_Rst.Fields("iPreferenceDay3Flag").Value & _
            "-" & IIf(l_obj_Rst.Fields("iSex") = 0, "(*)", "")
            .MoveNext
        Loop
        
        .Close
        Set l_obj_Rst = Nothing
    End With
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub l_void_PopulateRoomCombo(ByVal l_str_vDay As String)
    ' fill the room combo based on the day selected in the day combo
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    Dim l_str_Sql As String                 ' SQL string
    Dim l_int_NoOfRooms As Long          ' to store the number of rooms
    Dim l_int_Counter As Long            ' counter
    
    On Error GoTo ErrorHandler
    
    ' get the current selected day and room, and their capacities
    l_str_Sql = "SELECT iNumberOfRoomDay1, iNumberOfRoomDay2, iNumberOfRoomDay3," & _
        " dtSecondExamDay1, dtSecondExamDay2, dtSecondExamDay3," & _
        " iNumberOfExamineeDay1, iNumberOfExamineeDay2, iNumberOfExamineeDay3" & _
        " FROM tbSTESecondExamProfile" & _
        " WHERE iSystemProfileId = (" & _
        " SELECT iSystemProfileId FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag = 1)"
    
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    If Not l_obj_Rst.EOF Then
        Select Case UCase(l_str_vDay)
        Case UCase(LoadResString(2424))
            f_dt_SplDay = l_obj_Rst("dtSecondExamDay1")
            f_int_SplDayMax = l_obj_Rst("iNumberOfExamineeDay1")
            l_int_NoOfRooms = l_obj_Rst("iNumberOfRoomDay1")
        Case UCase(LoadResString(2425))
            f_dt_SplDay = l_obj_Rst("dtSecondExamDay2")
            f_int_SplDayMax = l_obj_Rst("iNumberOfExamineeDay2")
            l_int_NoOfRooms = l_obj_Rst("iNumberOfRoomDay2")
        Case UCase(LoadResString(2426))
            f_dt_SplDay = l_obj_Rst("dtSecondExamDay3")
            f_int_SplDayMax = l_obj_Rst("iNumberOfExamineeDay3")
            l_int_NoOfRooms = l_obj_Rst("iNumberOfRoomDay3")
        End Select
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing

    ' to check whether the max capacity of the day is reached or not
    If g_int_ExamType = 1 Or g_int_ExamType = 2 Then
        l_str_Sql = "SELECT e.iExamineeProfileId FROM tbSTEExamineeProfile e WHERE"
    Else
        l_str_Sql = "SELECT r.iExamineeProfileId FROM tbSTEExamineeRoomProfile r inner join tbSTEExamineeProfile e"
        l_str_Sql = l_str_Sql & " on e.iExamineeProfileId = r.iExamineeProfileId"
        l_str_Sql = l_str_Sql & " WHERE r.iSubjectProfileId = (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName='" & cboSubject.Text & "') AND " & _
            "  "
    End If
    Select Case UCase(l_str_vDay)
    Case UCase(LoadResString(2424))
        l_str_Sql = l_str_Sql & "  CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay1,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"
    Case UCase(LoadResString(2425))
        l_str_Sql = l_str_Sql & "  CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay2,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"
    Case UCase(LoadResString(2426))
        l_str_Sql = l_str_Sql & "  CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay3,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"
    End Select
    
    l_str_Sql = l_str_Sql & " AND e.iNendo = " & g_int_CurrentNendo
    
    l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
    If Not l_obj_Rst.EOF Then
        f_int_SplDayCount = l_obj_Rst.RecordCount
    Else
        f_int_SplDayCount = 0
    End If
    txtTotalExamineesDay.Text = f_int_SplDayCount
    txtWemenDay.Text = lfCountWemen
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cboSubject_Click()  ' for special interview/report
    Dim l_str_Sql As String                 ' SQl string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    
    On Error GoTo ErrorHandler
    
    l_str_Sql = "SELECT iExamType FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "'"
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    If Not l_obj_Rst.EOF Then
        f_int_ExamType = l_obj_Rst.Fields("iExamType").Value
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    
    Call l_void_PopulateDayCombo
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cboSplDay_Click()   ' for special interview/report
    Call l_void_PopulateRoomCombo(cboSplDay.Text)
    Call l_void_PopulateList(f_dt_SplDay)
    Call l_void_PopulateRoomCombo(cboSplDay.Text)
    lblSourceCapacity.Caption = CStr(f_int_SplRoomMax)
End Sub

Private Sub l_void_PopulateSubject()
    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    
    On Error GoTo ErrorHandler
    
    l_str_Sql = "SELECT vSubjectName FROM tbSTESubjectProfile" & _
        " WHERE iSubType = 3 "
'        " WHERE iExamType IN(2,4)"
    With l_obj_Rst
        .Open l_str_Sql, g_obj_Conn
        cboSubject.Clear
        
        Do While Not .EOF
            cboSubject.AddItem .Fields("vSubjectName").Value
            .MoveNext
        Loop
    
        If cboSubject.ListCount > 0 Then
            lblErrorDetails.Caption = ""
            cboSubject.ListIndex = 0
            Call l_void_PopulateDayCombo
        Else
            lblErrorDetails.Caption = LoadResString(2499)
        End If
    End With
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Activate()
    fMainForm.mnuTools.Enabled = False
    Dim Index As Long
    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    LoadResStrings Me
    Me.Caption = LoadResString(2431)
    g_void_SetFontProperties Me
    l_void_PopulateSubject
    lblDayTotal.Caption = LoadResString(2487)
    lblDayRoomTotal.Caption = LoadResString(2488)
    Call f_void_CheckButtonStatus
    cmdSelectall.Visible = False
    cmdDeselect.Visible = False
    cmdDeselectall.Visible = False
    lblDayRoomTotal.Visible = False
    txtTotal.Visible = False
    cboSubject.Visible = False
    lblSubject.Visible = False
    lblSourceCapacity.Visible = False
    Label4.Visible = False
    lstExaminee.Font = "ＭＳ ゴシック"
    lstAllotted.Font = "ＭＳ ゴシック"
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Function f_bln_FreeExaminee(ByVal l_int_iExamineeId As Integer) As Boolean
    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_RstExaminee As New ADODB.Recordset    ' recordset object
    Dim l_int_RecCount As Long           ' to store the total no of records
    
    On Error GoTo ErrorHandler
    l_str_Sql = "SELECT iExamineeRoomProfileId FROM tbSTEExamineeRoomProfile" & _
        " WHERE iExamineeProfileId=" & l_int_iExamineeId
    l_obj_RstExaminee.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
    l_int_RecCount = l_obj_RstExaminee.RecordCount
        
    l_str_Sql = "DELETE FROM tbSTEExamineeRoomProfile" & _
        " WHERE iExamineeProfileId=" & l_int_iExamineeId & _
        " AND iSubjectProfileId IN ("
    
    If f_int_ExamType = 2 Or f_int_ExamType = 3 Then
        l_str_Sql = l_str_Sql & "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE iExamType = 2 Or iExamType = 3)"
    Else
        l_str_Sql = l_str_Sql & "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName ='" & cboSubject.Text & "')"
    End If
    
    g_obj_Conn.Execute l_str_Sql
    
    If l_int_RecCount <= 1 Then
        l_str_Sql = "UPDATE tbSTEExamineeProfile SET dtSecondExamDay=null," & _
            " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
            " WHERE iExamineeProfileId=" & l_int_iExamineeId
    
        g_obj_Conn.Execute l_str_Sql
    End If
    
    f_bln_FreeExaminee = True
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
    f_bln_FreeExaminee = False
End Function

Public Sub f_void_CheckButtonStatus()
    'Procedure to check the status of the buttons
    'i.e enabling and disabling the buttons based on the presense
    'and selection of data in the list boxes

    If lstExaminee.ListCount = 0 Then
        cmdSelectall.Enabled = False
        cmdSelect.Enabled = False
    Else
        cmdSelectall.Enabled = True
        If lstExaminee.SelCount > 0 Then
            cmdSelect.Enabled = True
        Else
            cmdSelect.Enabled = False
        End If
    End If
    
    If lstAllotted.ListCount = 0 Then
        cmdDeselect.Enabled = False
        cmdDeselectall.Enabled = False
    Else
        cmdDeselectall.Enabled = True
        If lstAllotted.SelCount > 0 Then
            cmdDeselect.Enabled = True
        Else
            cmdDeselect.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call g_void_CloseChildForm
End Sub

Private Sub lstAllotted_Click()
    Call f_void_CheckButtonStatus
End Sub

Private Sub lstAllotted_DblClick()
'    Call cmdDeselect_Click
End Sub

Private Sub lstExaminee_Click()
    Call f_void_CheckButtonStatus
End Sub

Private Sub lstExaminee_DblClick()
    Call cmdSelect_Click
End Sub
