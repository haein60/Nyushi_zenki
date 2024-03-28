VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmRawScore 
   AutoRedraw      =   -1  'True
   ClientHeight    =   9855
   ClientLeft      =   2400
   ClientTop       =   2445
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Palette         =   "frmRawScore.frx":0000
   Picture         =   "frmRawScore.frx":3AD3
   ScaleHeight     =   9855
   ScaleWidth      =   13470
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdGetRows 
      Caption         =   "1036"
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
      Left            =   4898
      TabIndex        =   9
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "1071"
      Enabled         =   0   'False
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
      Left            =   4898
      TabIndex        =   11
      Top             =   7920
      Width           =   3135
   End
   Begin VB.ComboBox cboInterviewerID 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1080
      TabIndex        =   25
      Text            =   "Combo1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboInterviewer 
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
      TabIndex        =   1
      Top             =   1560
      Width           =   1830
   End
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
      Left            =   6720
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtJukenNoTo 
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
      Height          =   315
      Left            =   10560
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtJukenNoFrom 
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
      Height          =   315
      Left            =   6720
      ScrollBars      =   1  '水平
      TabIndex        =   4
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
      Left            =   2640
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1080
      Width           =   1830
   End
   Begin VB.ComboBox cboRoomId 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6840
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkChooseiScore 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   10560
      TabIndex        =   8
      Top             =   2100
      Width           =   200
   End
   Begin VB.CheckBox chkTotalMarks 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   6720
      TabIndex        =   7
      Top             =   2100
      Width           =   200
   End
   Begin VB.CheckBox chkExamineeName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   2640
      TabIndex        =   6
      Top             =   2100
      Width           =   200
   End
   Begin VB.TextBox txtRandomNo 
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
      Height          =   315
      Left            =   6720
      TabIndex        =   13
      Top             =   1080
      Width           =   1695
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfRawScore 
      Height          =   3885
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   12015
      _cx             =   21193
      _cy             =   6853
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16641260
      ForeColor       =   4194304
      BackColorFixed  =   16047044
      ForeColorFixed  =   8388608
      BackColorSel    =   8388608
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16047044
      BackColorAlternate=   16641260
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   1
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.ComboBox cboRoomName 
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
      Left            =   10560
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblFromToCount 
      Caption         =   "lblFromToCount"
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
      Left            =   9000
      TabIndex        =   26
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lblInterviewers 
      BackStyle       =   0  '透明
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
      Left            =   840
      TabIndex        =   24
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblDay 
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
      Left            =   4680
      TabIndex        =   23
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblRoomName 
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
      Left            =   8520
      TabIndex        =   22
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblJukenNoTo 
      BackStyle       =   0  '透明
      Caption         =   "1955"
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
      Left            =   8520
      TabIndex        =   21
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblJukenNoFrom 
      BackStyle       =   0  '透明
      Caption         =   "1952"
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
      Left            =   4440
      TabIndex        =   20
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  '透明
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
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "1962"
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
      Left            =   4440
      TabIndex        =   17
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "1963"
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
      Left            =   8400
      TabIndex        =   16
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "1805"
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
      TabIndex        =   15
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblRandomNo 
      BackStyle       =   0  '透明
      Caption         =   "1953"
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
      Left            =   4440
      TabIndex        =   12
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblErrorDetails 
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
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   12015
   End
End
Attribute VB_Name = "frmRawScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmRawScore
'Author         :   Dileep Cherian
'Created On     :   14/9/01
'Description    :   This screen is used to provide mechanism for inputting Raw Score for the examinee
'Reference      :   FunctionalSpecs OF Raw Score.doc(Ver1.1)
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History   -
'04/04/2002  -   Dileep Cherian
'User should be able to resize the coulmns, incase part of data is not visible in the normal display
'On pressing ente after editing a column value, the focus should move to the next row (same column)
'Ammendments - NyushiChangesSummary.doc ver 1.0
'10/5/2002 - Mahesh Deshpande
'Check boxes to hide/unhide ExamineeName, Total Marks,ChooseiScore Columns.
'Put in Day combo box and Room number combo box
'Interviewer combo box for reports scree.
'**************************************************************************************************
' to store the currently selected item
Dim m_int_SelectedSubject As Integer    ' to keep track of row change
Dim m_int_CurrentRow As Integer         ' to identify when the edit starts
Dim m_bln_Edit As Boolean               ' to store the subject profile ID to pass to stored procedure
Dim m_int_SubjId As Integer             ' to store number of questions, to pass to the stored procedure
Dim m_int_NoOfQues As Integer           ' to store the scores of diff Questions
Dim m_int_Score(9) As String            ' to store the examinee Id
Dim m_int_ExamineeId() As Integer       ' to store the subject question profile id
Dim m_intSubQuesProfileId(9) As String  ' to store the total score
Dim m_int_TotalScore As Double          ' to store choosei score
Dim m_int_ChooseiScore As Double        ' database related variables
Dim m_obj_Rst As New ADODB.Recordset
Dim m_str_SQl As String
Dim m_int_QuestionLimit As Integer

Private Const prvclNoCol As Long = 0
Private Const prvcsHyotei As String = "評定値" '列見出しがこの値のとき、隣の列が成績概評出力だと判断する
Private Const prvcsSeisekiGaihyo As String = "成績概評"
Private m_SecondExam_Type As Integer '面接か小論文か

Public Sub gsSetSecondType(piSType As Integer)

    If piSType = 0 Then
        m_SecondExam_Type = 0 '面接
    Else
        m_SecondExam_Type = 1 '小論文
    End If

End Sub

Private Sub cboDay_Click()
'    If g_int_ExamType = 3 Or g_int_ExamType = 5 Then
'        f_void_populateInterviewers
'    End If
End Sub

Private Sub cboInterviewer_Click()
    If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
        cboInterviewerID.ListIndex = cboInterviewer.ListIndex
    End If
End Sub

Private Sub cboRoomName_Click()
    cboRoomId.ListIndex = cboRoomName.ListIndex
    If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
        f_void_populateInterviewers
    End If
End Sub

Private Sub cboSubject_Click()
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    On Error GoTo ErrorHandler
    
    l_str_Sql = "SELECT iSubjectProfileId  FROM tbSTESubjectProfile WHERE" & _
        " vSubjectName ='" & cboSubject.Text & "'"
    l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
    If Not l_obj_Rst.EOF Then
        m_int_SelectedSubject = l_obj_Rst("iSubjectProfileId")
    Else
        If g_int_ExamType = 0 Then
            m_int_SelectedSubject = -1
        End If
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Call f_void_ClearGrid
    If g_int_ExamType = 3 Or g_int_ExamType = 5 Then
        f_void_populateInterviewers   'Required for population of interviewers requires for Raw Score for Report
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub chkChooseiScore_Click()
'    With vsfRawScore
'    .Col = .Cols - 1
'    .ColWidth(.Cols - 1) = IIf(chkChooseiScore.Value = 0, 0, 1800)
'    End With
End Sub

Private Sub chkExamineeName_Click()
'    With vsfRawScore
'    .Col = 2
'    .ColWidth(.Col) = IIf(chkExamineeName.Value = 0, 0, 2200)
'     End With
End Sub

Private Sub chkTotalMarks_Click()
'    With vsfRawScore
'    .Col = 3
'    .ColWidth(.Col) = IIf(chkTotalMarks.Value = 0, 0, 1500)
'    End With
End Sub

' if juken no from and to is entered, then get data for that range
' otherwise get all the data
Private Sub cmdGetRows_Click()
    Dim l_int_Counter As Integer
    On Error GoTo ErrorHandler
    
    Select Case g_int_ExamType
    Case 0
        If Len(Trim(txtJukenNoFrom.Text)) = 0 Then
            lblErrorDetails.Caption = "受験番号自を指定してください。"
            lblErrorDetails.Visible = True
            txtJukenNoFrom.SetFocus
            Exit Sub
        End If
        If Len(Trim(txtJukenNoTo.Text)) = 0 Then
            lblErrorDetails.Caption = "受験番号至を指定してください。"
            lblErrorDetails.Visible = True
            txtJukenNoTo.SetFocus
            Exit Sub
        End If
        ' vaidate the from and to values
        If CInt(Trim(txtJukenNoFrom.Text)) > CInt(Trim(txtJukenNoTo.Text)) Then
            lblErrorDetails.Caption = LoadResString(1960)
            lblErrorDetails.Visible = True
            txtJukenNoTo.SetFocus
            Exit Sub
        End If
        lblErrorDetails.Caption = ""
        lblErrorDetails.Visible = False
    
    Case 1, 2
        
        m_str_SQl = "SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType=" & g_int_ExamType & " AND vSubjectName='" & cboSubject.Text & "'"
        Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQl)

        If Not m_obj_Rst.EOF Then
            m_int_SelectedSubject = m_obj_Rst("iSubjectProfileId")
        End If

        ' release the object variables
        m_obj_Rst.Close
        Set m_obj_Rst = Nothing

        If g_int_ExamType = 2 And m_SecondExam_Type = 0 Then
        'インタビュアーのチェック
            m_str_SQl = "SELECT count(*) as cnt "
            m_str_SQl = m_str_SQl & " FROM tbSTESubjectQuestionProfile a , "
            m_str_SQl = m_str_SQl & "      tbSTEInterviewRoomProfile c , "
            m_str_SQl = m_str_SQl & "      tbSTEInterviewerProfile d "
            m_str_SQl = m_str_SQl & " WHERE a.iSubjectProfileId = " & m_int_SelectedSubject
            m_str_SQl = m_str_SQl & " AND  a.iSubjectProfileId = c.iSubjectProfileId"
            m_str_SQl = m_str_SQl & " AND  c.iRoomProfileId = " & cboRoomId.Text & " "
            m_str_SQl = m_str_SQl & " AND  c.iDayFlag = " & Me.cboDay.ListIndex & " "
            m_str_SQl = m_str_SQl & " AND  d.iInterviewerProfileId = c.iInterviewerProfileId"

            Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQl)

            If m_obj_Rst.Fields(0) = 0 Then
                m_obj_Rst.Close
                Set m_obj_Rst = Nothing
                lblErrorDetails.Caption = "採点者が設定されていません。" ' "No interviewers associated with this subject"
                lblErrorDetails.Visible = True
                Exit Sub
            End If

            m_obj_Rst.Close
            Set m_obj_Rst = Nothing

        End If

    End Select
    If (g_int_ExamType = 2 And m_SecondExam_Type = 1) And cboInterviewer.Text = "" Then
        lblErrorDetails.Caption = LoadResString(2484) ' "No interviewers associated with this subject"
        lblErrorDetails.Visible = True
        With vsfRawScore
            .Rows = 2
            .Cols = 15
            .Row = 0
            For l_int_Counter = 0 To .Cols - 1
                .Col = l_int_Counter
                .Text = ""
            Next
            .Row = 1
            For l_int_Counter = 0 To .Cols - 1
                .Col = l_int_Counter
                .Text = ""
            Next
            .Enabled = False
        End With
        cmdUpdate.Enabled = False
        Exit Sub
    Else
        vsfRawScore.Enabled = True
        cmdUpdate.Enabled = True
        
        Call f_void_ClearGrid
        Call f_void_InitializeGrid
    End If

'    With vsfRawScore
'        .Col = .Cols - 1
'        .ColWidth(.Cols - 1) = IIf((chkChooseiScore.Value = 0) Or (Me.cboSubject.Text = "欠席日数"), 0, 1800)
'        .Col = 2
'        .ColWidth(.Col) = IIf((chkExamineeName.Value = 0), 0, 2200)
'        .Col = 3
'        .ColWidth(.Col) = IIf((chkTotalMarks.Value = 0) Or (Me.cboSubject.Text = "欠席日数"), 0, 1500)
'    End With

    If vsfRawScore.Rows > 1 Then
'        lblFromToCount.Caption = vsfRawScore.TextMatrix(1, 1) & "～" & vsfRawScore.TextMatrix(vsfRawScore.Rows - 1, 1) & "  " & Trim(str(vsfRawScore.Rows - 1)) & "件"
        lblFromToCount.Caption = Trim(str(vsfRawScore.Rows - 1)) & "件"
    Else
        lblFromToCount.Caption = ""
    End If

    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdUpdate_Click()
    Dim l_int_Counter As Integer
    Dim l_int_LoopCounter As Integer
    Dim l_bln_Update As Boolean
    Dim l_str_ErrString As String
    
    On Error GoTo ErrorHandler
    l_bln_Update = True
'このあとのRowcolchangeなどでトータルスコアのカラムのデータが更新されている。
'よってm_bln_Editを変更してはならない

    With vsfRawScore
    .Row = 1
    
    For l_int_LoopCounter = 1 To .Rows - 1
        .Row = l_int_LoopCounter
        .Col = 3
        m_int_TotalScore = IIf(Trim(.Text) = "", 0, CDbl(.Text))
        .Col = 4
        For l_int_Counter = 0 To m_int_QuestionLimit - 1
            If Trim(.Text) = "a" Then
                m_int_Score(l_int_Counter) = -1
            ElseIf Trim(.Text) = "b" Then
                m_int_Score(l_int_Counter) = -2
            Else
                m_int_Score(l_int_Counter) = IIf(Trim(.Text) = "", 0, Trim(.Text))
            End If
            If .TextMatrix(0, .Col) = prvcsHyotei Then
                .Col = .Col + 2
            Else
                .Col = .Col + 1
            End If
        Next
        m_int_ChooseiScore = IIf(Trim(.Text) = "", 0, Trim(.Text))
        l_bln_Update = f_bln_UpdateData(l_int_LoopCounter)
        If Not l_bln_Update Then
            l_str_ErrString = l_str_ErrString & CStr(l_int_LoopCounter) & ","
        End If
    Next
    lblErrorDetails.Visible = True
    If Not l_bln_Update Then
        l_str_ErrString = Left(l_str_ErrString, Len(l_str_ErrString) - 1)
        lblErrorDetails = LoadResString(2437) & l_str_ErrString
    Else
        lblErrorDetails.Caption = LoadResString(2404)
    End If
    End With

Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Activate()
    fMainForm.mnuTools.Enabled = False
    Dim index As Integer
    For index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(index).Enabled = False
    Next
End Sub

Private Sub Form_Load()
    ' counter
    Dim l_int_Counter As Integer
    
    On Error GoTo ErrorHandler
    
    LoadResStrings Me
    Me.Caption = LoadResString(1951)
    Call g_void_SetFontProperties(Me)     ' set the font properties
    ' select all subjects that come under the current exam type
    m_str_SQl = "SELECT iSubjectProfileId,vSubjectName " & _
        " FROM tbSTESubjectProfile"
    If g_int_ExamType = 0 Then
        m_str_SQl = m_str_SQl & " WHERE iExamType = 0"
        m_str_SQl = m_str_SQl & " AND iSubType = 0"
    ElseIf g_int_ExamType = 1 Then
        m_str_SQl = m_str_SQl & " WHERE iExamType = 1"
    ElseIf g_int_ExamType = 2 Then
        If m_SecondExam_Type = 0 Then
            m_str_SQl = m_str_SQl & " WHERE iSubType = 3"
        Else
            m_str_SQl = m_str_SQl & " WHERE iSubType = 4"
        End If
    End If
    m_str_SQl = m_str_SQl & " ORDER BY iDispOrder"
    Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQl)
    
    If Not m_obj_Rst.EOF Then
        ' make the first subject as default selected
        m_int_SelectedSubject = m_obj_Rst("iSubjectProfileId")
        lblSubject.Visible = True
        cboSubject.Visible = True
        Do While Not m_obj_Rst.EOF
            cboSubject.AddItem m_obj_Rst("vSubjectName")
            m_obj_Rst.MoveNext
        Loop
        If g_int_ExamType = 0 Then
            lblJukenNoFrom.Visible = True
            txtJukenNoFrom.Visible = True
            lblJukenNoTo.Visible = True
            txtJukenNoTo.Visible = True
            txtRandomNo.Visible = False
            lblRandomNo.Visible = False
            cboDay.Visible = False
            lblDay.Visible = False
            lblRoomName.Visible = False
            cboRoomName.Visible = False
            lblInterviewers.Visible = False
            cboInterviewer.Visible = False
            chkExamineeName.Value = 1
            Label4.Visible = False
            chkTotalMarks.Visible = False
            Label3.Visible = False
            chkChooseiScore.Visible = False
        ElseIf g_int_ExamType = 1 Then
            txtRandomNo.Visible = True
            lblRandomNo.Visible = True
            lblDay.Visible = False
            cboDay.Visible = False
            lblRoomName.Visible = False
            cboRoomName.Visible = False
            lblJukenNoFrom.Visible = False
            txtJukenNoFrom.Visible = False
            lblJukenNoTo.Visible = False
            txtJukenNoTo.Visible = False
            lblInterviewers.Visible = False
            cboInterviewer.Visible = False
            Label2.Caption = "受験番号"
            Label4.Visible = False
            chkTotalMarks.Visible = False
            Label3.Visible = False
            chkChooseiScore.Visible = False
        ElseIf g_int_ExamType = 2 And m_SecondExam_Type = 0 Then
            txtRandomNo.Visible = False
            lblRandomNo.Visible = False
            lblDay.Visible = True
            cboDay.Visible = True
            lblRoomName.Visible = True
            cboRoomName.Visible = True
            lblJukenNoFrom.Visible = False
            txtJukenNoFrom.Visible = False
            lblJukenNoTo.Visible = False
            txtJukenNoTo.Visible = False
            lblInterviewers.Visible = False
            cboInterviewer.Visible = False
            Label3.Visible = False
            chkChooseiScore.Visible = False
            ' add the subjects to combo box
            Call l_void_AddRooms  'Populate Room Combo
            Call l_void_PopulateDayCombo 'Populate Day combo
        ElseIf g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
            txtRandomNo.Visible = False
            lblRandomNo.Visible = False
'乱数で一意なため、日付は削除
            lblDay.Visible = False
            cboDay.Visible = False
'乱数で一意なため、日付は削除end
            lblRoomName.Visible = True
            cboRoomName.Visible = True
            lblJukenNoFrom.Visible = False
            txtJukenNoFrom.Visible = False
            lblJukenNoTo.Visible = False
            txtJukenNoTo.Visible = False
            lblInterviewers.Visible = True
            cboInterviewer.Visible = True
            ' add the subjects to combo box
            Call l_void_AddRoomsRand  'Populate Room Combo
'乱数で一意なため、日付は削除
'            Call l_void_PopulateDayCombo 'Populate Day combo
            Call f_void_populateInterviewers       'Populate interviewers combo
            Label2.Caption = "受験番号"
'            Label2.Visible = False
'            chkExamineeName.Visible = False
'            Label4.Visible = False
'            chkTotalMarks.Visible = False
            Label3.Visible = False
            chkChooseiScore.Visible = False
        End If
    End If
    ' release the object variables
    m_obj_Rst.Close
    Set m_obj_Rst = Nothing

    cboSubject.ListIndex = 0
    Call f_void_InitGrid
    
    ' initialize array values to zero
    For l_int_Counter = 0 To 9
        m_int_Score(l_int_Counter) = 0
    Next
    lblFromToCount.Caption = ""
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_PopulateGrid()
    Dim l_int_Counter As Integer
    Dim l_int_SrNo As Integer
    Dim l_int_ScoreId As Integer
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    Dim l_dbl_ChooseiScore As Double
    Dim l_int_StrLen As Integer
    Dim l_int_Cnt As Integer
    Dim l_obj_Populate  As New ADODB.Recordset
    Dim l_int_CurYear As Integer
    Dim l_str_SqlDay As String
    Dim l_obj_rstDay As New ADODB.Recordset
    Dim f_dt_SourceDay As Date
    Dim f_int_SourceDayMax As Integer
    Dim l_int_NoOfRooms As Integer
    Dim lCol As Long
    Dim lsvCol As Long
    
    On Error GoTo ErrorHandler
    l_int_CurYear = g_int_CurrentNendo  'global variable in form load
    
    '**** New code added for day,room and subject (Mahesh)
    
    l_str_SqlDay = "SELECT iNumberOfRoomDay1, iNumberOfRoomDay2, iNumberOfRoomDay3," & _
        " dtSecondExamDay1, dtSecondExamDay2, dtSecondExamDay3," & _
        " iNumberOfExamineeDay1, iNumberOfExamineeDay2, iNumberOfExamineeDay3" & _
        " FROM tbSTESecondExamProfile" & _
        " WHERE iSystemProfileId = (" & _
        " SELECT iSystemProfileId FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag = 1)"
    
    l_obj_rstDay.Open l_str_SqlDay, g_obj_Conn
    If Not l_obj_rstDay.EOF Then
        Select Case UCase(cboDay.Text)
        Case UCase(LoadResString(2424))     ' day 1
            f_dt_SourceDay = l_obj_rstDay("dtSecondExamDay1")
            f_int_SourceDayMax = l_obj_rstDay("iNumberOfExamineeDay1")
            l_int_NoOfRooms = l_obj_rstDay("iNumberOfRoomDay1")
        Case UCase(LoadResString(2425))     ' day 2
            f_dt_SourceDay = l_obj_rstDay("dtSecondExamDay2")
            f_int_SourceDayMax = l_obj_rstDay("iNumberOfExamineeDay2")
            l_int_NoOfRooms = l_obj_rstDay("iNumberOfRoomDay2")
        Case UCase(LoadResString(2426))     ' day 3
            f_dt_SourceDay = l_obj_rstDay("dtSecondExamDay3")
            f_int_SourceDayMax = l_obj_rstDay("iNumberOfExamineeDay3")
            l_int_NoOfRooms = l_obj_rstDay("iNumberOfRoomDay3")
        End Select
    End If
    l_obj_rstDay.Close
    Set l_obj_rstDay = Nothing
    '**** (New Code Ends)
    
    If Len(Trim(txtRandomNo.Text)) <> 0 Then
        If g_int_ExamType = 1 Then
            m_str_SQl = "SELECT e.iExamineeProfileId, e.iJukenNumber,e.vExamineeName " & _
                " FROM tbSTEExamineeProfile e,tbSTERoomProfile r " & _
                " WHERE r.iRoomProfileId = e.iRoomProfileId " & _
                " AND e.iNendo = " & l_int_CurYear & _
                " AND r.iRandom =" & txtRandomNo.Text & _
                " AND e.iExamineeStatus = 0" & _
                " AND e.iAbsentFlag = 0"
                
                
        ' comdesign
        '
        '
        '
            Select Case Trim(cboSubject.Text)
            Case "数学"
                m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0"
            Case "英語"
                m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND iLanguageSubjProfileId=" & m_int_SubjId
            Case "独語"
                m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND iLanguageSubjProfileId=" & m_int_SubjId
            Case "仏語"
                m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND iLanguageSubjProfileId=" & m_int_SubjId
            Case "物理"
                m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
            Case "化学"
                m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
            Case "生物"
                m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
            End Select
              
                
                
                
        Else
            m_str_SQl = "SELECT iExamineeProfileId, iJukenNumber, vExamineeName from tbSTEExamineeProfile" & _
                " WHERE iexamineeprofileid in(SELECT iExamineeProfileId FROM tbSTEExamineeRoomProfile" & _
                " WHERE iRoomProfileId=(SELECT iRoomProfileId FROM tbSTERoomProfile " & _
                " WHERE iRandom =" & txtRandomNo.Text & "))" & _
                " AND iNendo = " & l_int_CurYear & _
                " AND iExamineeStatus = 1" & _
                " AND iAbsentFlag = 0"
        End If
    Else
        m_str_SQl = "Select iExamineeProfileId, iJukenNumber, vExamineeName" & _
            " from tbSTEExamineeProfile " & _
            " WHERE iNendo=" & l_int_CurYear
        
        If g_int_ExamType = 1 Then
            m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0"
                Select Case Trim(cboSubject.Text)
                Case "数学"
                    m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0"
                Case "英語"
                    m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND iLanguageSubjProfileId=" & m_int_SubjId
                Case "独語"
                    m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND iLanguageSubjProfileId=" & m_int_SubjId
                Case "仏語"
                    m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND iLanguageSubjProfileId=" & m_int_SubjId
                Case "物理"
                    m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
                Case "化学"
                    m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
                Case "生物"
                    m_str_SQl = m_str_SQl & " AND iExamineeStatus = 0 AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
                End Select
            
        ElseIf g_int_ExamType = 2 Then
            'Changes to sql start (Mahesh)
            If m_SecondExam_Type = 0 Then
                m_str_SQl = m_str_SQl & " AND iExamineeStatus = 1" & _
                    " AND iExamineeProfileId IN" & _
                    " (SELECT iExamineeProfileId " & _
                    " From tbSteExamineeRoomProfile" & _
                    " WHERE iSubjectProfileid = " & m_int_SubjId & " AND iRoomProfileid = " & cboRoomId.Text & ") AND" & _
                    " dtSecondExamDay = '" & Format(f_dt_SourceDay, "MM/DD/YYYY") & "'"
            Else
                m_str_SQl = m_str_SQl & " AND iExamineeStatus = 1"
                m_str_SQl = m_str_SQl & " AND iShoronbunRandomNo = " & cboRoomId.Text & " "
                m_str_SQl = m_str_SQl & " AND exists ( SELECT 1 FROM tbSTEInterviewRoomProfile as ir "
                m_str_SQl = m_str_SQl & "       WHERE ir.iRandomNo = " & cboRoomId.Text & " "
                m_str_SQl = m_str_SQl & "       AND ir.iInterviewerProfileId = " & Me.cboInterviewerID.Text & " ) "
            End If
            'Changes end
        End If
        
        m_str_SQl = m_str_SQl & " AND iAbsentFlag = 0"
        
        If Trim(txtJukenNoFrom.Text) <> "" And Trim(txtJukenNoTo.Text) <> "" Then
            m_str_SQl = m_str_SQl & " AND iJukenNumber between " & txtJukenNoFrom.Text & " AND " & txtJukenNoTo.Text
        End If
    End If
    
    l_obj_Populate.Open m_str_SQl, g_obj_Conn, adOpenStatic, adLockReadOnly

    If Not l_obj_Populate.EOF Then
        lblErrorDetails.Caption = ""
        cmdUpdate.Enabled = True
        ' enable the grid
        vsfRawScore.Enabled = True

        With vsfRawScore
            Do While Not l_obj_Populate.EOF              'loop thru recordset to populate grid
               
                l_int_SrNo = l_int_SrNo + 1
                ' store examinee id for later use
                ReDim Preserve m_int_ExamineeId(l_int_SrNo)
                
                m_int_ExamineeId(l_int_SrNo) = l_obj_Populate("iExamineeProfileId")
                .Row = l_obj_Populate.AbsolutePosition
                .CellPictureAlignment = flexAlignCenterCenter
                
                .Col = 0
                .Text = l_int_SrNo
                .Col = .Col + 1
'                If g_int_ExamType <> 0 And g_int_ExamType <> 1 Then
'                    l_int_StrLen = Len(l_obj_Populate("iJukenNumber"))
'                    For l_int_Cnt = 0 To l_int_StrLen - 1
'                        .Text = .Text & "*"
'                    Next
'                Else
                    .Text = l_obj_Populate("iJukenNumber")
'                End If

                .Col = .Col + 1
                If g_int_ExamType <> 0 Then
                    l_int_StrLen = Len(l_obj_Populate("vExamineeName"))
                    For l_int_Cnt = 0 To l_int_StrLen - 1
                        .Text = .Text & "*"
                    Next
                Else
                    .Text = l_obj_Populate("vExamineeName")
                End If
                               
                .Col = .Col + 1
               
                l_str_Sql = "SELECT s.iScoreprofileId, s.fRawScore, s.fChoseiScore" & _
                    " FROM tbSTEScoreProfile s, tbSTEExamineeProfile e" & _
                    " WHERE s.iSubjectProfileId=" & m_int_SubjId & _
                    " AND e.iExamineeProfileId=" & m_int_ExamineeId(l_int_SrNo) & _
                    " AND e.iExamineeProfileId = s.iExamineeProfileId"
                Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
               
                If Not l_obj_Rst.EOF Then
                    If Trim(l_obj_Rst("fRawScore")) = "-1" Then
                        .Text = "a"
                    ElseIf Trim(l_obj_Rst("fRawScore")) = "-2" Then
                        .Text = "b"
                    ElseIf Trim(l_obj_Rst("fRawScore")) <> "" Then
                        If .TextMatrix(0, .Col) = "欠席日数" Then
                            .Text = l_obj_Rst("fRawScore")
                        Else
                            .Text = Format(l_obj_Rst("fRawScore"), "##0.00") '小数点以下は常に２位まで
                        End If
                        If .Text = "0" Then
                            .Text = ""
                        End If
                    Else
                        .Text = ""
                    End If
                    l_int_ScoreId = l_obj_Rst("iScoreProfileId")
                    If Trim(l_obj_Rst("fChoseiScore")) <> "" Then
                        l_dbl_ChooseiScore = l_obj_Rst("fChoseiScore")
                    Else
                        l_dbl_ChooseiScore = 0
                    End If
                Else
                    .Text = ""
                End If
                
                ' release the object variable
                Set l_obj_Rst = Nothing

                l_str_Sql = "SELECT d.fDetailScore FROM tbSTEScoreDetail d, tbSTEScoreProfile s "
                l_str_Sql = l_str_Sql & " WHERE d.iScoreProfileId = " & l_int_ScoreId
                l_str_Sql = l_str_Sql & " AND s.iExamineeProfileId = " & m_int_ExamineeId(l_int_SrNo)
                l_str_Sql = l_str_Sql & " AND s.iSubjectProfileId = " & m_int_SubjId
                l_str_Sql = l_str_Sql & " AND s.iScoreProfileId = d.iScoreProfileId"
                If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
                    l_str_Sql = l_str_Sql & " AND d.iSubjectQuestionId = " & cboInterviewerID.Text
                End If
                Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
                
                ' set the question limit
                
                If m_int_NoOfQues <= 10 Then
                    m_int_QuestionLimit = m_int_NoOfQues
                    lblErrorDetails.Caption = ""
                    lblErrorDetails.Visible = False
                Else
                    m_int_QuestionLimit = 10
                    lblErrorDetails.Caption = LoadResString(2500) & m_int_NoOfQues & " " & LoadResString(2501)
                    lblErrorDetails.Visible = True
                End If
                
                For l_int_Counter = 1 To m_int_QuestionLimit
                   ' this will initialize chosei score also to zero
                   If Not l_obj_Rst.EOF Then
                        .Col = .Col + 1
                        If .TextMatrix(0, .Col) <> "成績概評" Then
                            .CellBackColor = &HC0C0FF
                        End If
                        If Trim(l_obj_Rst("fDetailScore")) = "-1" Then
                            .Text = "a"
                        ElseIf Trim(l_obj_Rst("fDetailScore")) = "-2" Then
                            .Text = "b"
                        ElseIf Trim(l_obj_Rst("fDetailScore")) <> "" Then
                            If l_obj_Rst("fDetailScore") = 0 Then
                                .Text = ""
                            Else
                                If .TextMatrix(0, .Col) = "欠席日数" Then
                                    .Text = l_obj_Rst("fDetailScore")
                                Else
                                    .Text = Format(l_obj_Rst("fDetailScore"), "##0.00") '小数点以下は常に２位まで
                                End If
                            End If
                        Else
                            .Text = ""
                        End If
                        If .TextMatrix(0, .Col) = prvcsHyotei Then
                            .Col = .Col + 1
                            .Text = gfSeisekiGaihyo(l_obj_Rst("fDetailScore"))
                        End If
                        l_obj_Rst.MoveNext
                   Else
                        .Col = .Col + 1
                        If .TextMatrix(0, .Col) <> "成績概評" Then
                            .CellBackColor = &HC0C0FF
                        End If
                        .Text = ""
                    End If
                Next
               
                ' release the object variable
                Set l_obj_Rst = Nothing
               
                .Col = .Col + 1
                If .TextMatrix(0, .Col) <> "成績概評" Then
                    .CellBackColor = &HC0C0FF
                End If
                If l_dbl_ChooseiScore = 0 Then
                   .Text = ""
                Else
                   .Text = l_dbl_ChooseiScore
                End If
                .Rows = .Rows + 1                    'add a new row to the grid
                l_obj_Populate.MoveNext
            Loop
           
            .Rows = .Rows - 1                       'remove the last row because it's blank
            .Row = 1
        End With
    Else
        cmdUpdate.Enabled = False
        lblErrorDetails.Caption = LoadResString(1964)
        lblErrorDetails.Visible = True
        Call f_void_ClearGrid
        vsfRawScore.Enabled = False
    End If
    
    ' release the object variable
    l_obj_Populate.Close
    Set l_obj_Populate = Nothing
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub


Private Sub optRandom_Click()
    cboSubject.Enabled = False
End Sub

Private Sub optSubject_Click()
    txtRandomNo.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call g_void_CloseChildForm
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

Private Sub txtRandomNo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsfRawScore_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfRawScore
        If .Col > 3 Then
            If Trim(.TextMatrix(Row, Col)) <> "" Then
                If IsNumeric(.TextMatrix(Row, Col)) Then
                    If .TextMatrix(0, .Col) <> "欠席日数" Then
                        .TextMatrix(Row, Col) = Format(Round(.TextMatrix(Row, Col), 2), "##0.00") '小数点以下は常に２位まで
                    End If
                    If .TextMatrix(0, Col) = prvcsHyotei Then
                        .TextMatrix(Row, Col + 1) = gfSeisekiGaihyo(.TextMatrix(Row, Col))
                    End If
                    If .TextMatrix(Row, Col) = 0 Then
                        .TextMatrix(Row, Col) = ""
                    End If
                End If
                ' change in comdesign, arka 19apr 2002 end
'NextCol:
'                If .Col < .Cols - 1 Then
'                    If .TextMatrix(0, Col) = prvcsHyotei Then
'                        If .Col < .Cols - 2 Then
'                            .Col = .Col + 2
'                        Else
'                            .Col = .Col + 1
'                        End If
'                    Else
'                        .Col = .Col + 1
'                    End If
'                    If .ColWidth(.Col) = 0 Then GoTo NextCol
'                Else
'                    If .Row < .Rows - 1 Then
'                        .Row = .Row + 1
'                    ElseIf .Row = .Rows - 1 Then
'                        .Row = .Rows - 2
'                    End If
'                    .Col = 3
'                    GoTo NextCol
'                End If
            Else
                If .TextMatrix(0, Col) = prvcsHyotei Then
                    .TextMatrix(Row, Col + 1) = ""
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfRawScore_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim l_bln_Sp As Boolean
    Dim l_int_Counter As Integer
    Dim l_bln_Update As Boolean
    Dim l_int_EditedRow As Integer

    On Error GoTo ErrorHandler
    
    If vsfRawScore.Row <> m_int_CurrentRow Then
        If m_bln_Edit Then
            m_bln_Edit = False
            With vsfRawScore
                .Row = OldRow
                l_int_EditedRow = .Row
                .Col = 4
                For l_int_Counter = 0 To m_int_QuestionLimit - 1
                    If Len(Trim(.Text)) <> 0 Then
                        If .TextMatrix(0, .Col) = "欠席日数" Then
                            If Trim(.Text) = "a" Then
                                m_int_Score(l_int_Counter) = -1
                            ElseIf Trim(.Text) = "b" Then
                                m_int_Score(l_int_Counter) = -2
                            Else
                                m_int_Score(l_int_Counter) = .Text
                            End If
                        Else
                            m_int_Score(l_int_Counter) = .Text
                        End If
                    Else
                        m_int_Score(l_int_Counter) = 0
                        .Text = ""
                    End If
                    If .TextMatrix(0, .Col) = prvcsHyotei Then
                        .Col = .Col + 2
                    Else
                        .Col = .Col + 1
                    End If
                Next
                If Len(Trim(.Text)) = 0 Then
                    .Text = 0
                End If

                .Col = OldCol
                If .Col > 3 Then
'                    If .TextMatrix(0, .Col) <> "欠席日数" And .TextMatrix(0, .Col) <> prvcsSeisekiGaihyo Then
                        l_bln_Sp = f_bln_CallSP()
                        If Not l_bln_Sp Then
                            lblErrorDetails.Caption = LoadResString(1956)
                            Exit Sub
                        Else
                            lblErrorDetails.Caption = ""
                            m_bln_Edit = False
                        End If
'                    End If
                End If
                .Row = NewRow
                .Col = NewCol
            End With
        End If
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub vsfRawScore_Click()
    vsfRawScore.EditCell
End Sub

Private Sub vsfRawScore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

    If Col < 4 Then
        KeyAscii = 0
    ElseIf vsfRawScore.TextMatrix(0, vsfRawScore.Col) = prvcsSeisekiGaihyo Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
'm_bln_Edit = False
        If g_int_ExamType = 2 And m_SecondExam_Type = 0 Then
        '縦移動
            If vsfRawScore.Row < vsfRawScore.Rows - 1 Then
                vsfRawScore.Row = vsfRawScore.Row + 1
            Else
NextCol1:
                If vsfRawScore.Col < vsfRawScore.Cols - 1 Then
                    vsfRawScore.Col = vsfRawScore.Col + 1
                    vsfRawScore.Row = 1
                    If vsfRawScore.ColWidth(vsfRawScore.Col) = 0 Or vsfRawScore.CellBackColor <> &HC0C0FF Then GoTo NextCol1
                Else
                    If cmdUpdate.Enabled Then
                        vsfRawScore.Row = 1
                        vsfRawScore.Col = 4
                        cmdUpdate.SetFocus
                    End If
                End If
            End If
        Else
        '横移動
NextCol:
            If vsfRawScore.Col < vsfRawScore.Cols - 1 Then
                If vsfRawScore.TextMatrix(0, vsfRawScore.Col) = prvcsHyotei Then
                    If vsfRawScore.Col < vsfRawScore.Cols - 2 Then
                        vsfRawScore.Col = vsfRawScore.Col + 2
                    Else
                        vsfRawScore.Col = vsfRawScore.Col + 1
                    End If
                Else
                    vsfRawScore.Col = vsfRawScore.Col + 1
                    If vsfRawScore.ColWidth(vsfRawScore.Col) = 0 Or vsfRawScore.CellBackColor <> &HC0C0FF Then GoTo NextCol
                End If
            Else
                If vsfRawScore.Row < vsfRawScore.Rows - 1 Then
                    vsfRawScore.Row = vsfRawScore.Row + 1
                    vsfRawScore.Col = 4
                    If vsfRawScore.ColWidth(vsfRawScore.Col) = 0 Or vsfRawScore.CellBackColor <> &HC0C0FF Then GoTo NextCol
                Else
                    If cmdUpdate.Enabled Then
                        vsfRawScore.Row = 1
                        vsfRawScore.Col = 4
                        cmdUpdate.SetFocus
                    End If
                End If
            End If
        End If
'm_bln_Edit = True
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab And KeyAscii <> 46 Then
        If Not ((KeyAscii = Asc("A") Or KeyAscii = Asc("B") Or KeyAscii = Asc("a") Or KeyAscii = Asc("b")) And vsfRawScore.TextMatrix(0, vsfRawScore.Col) = "欠席日数") Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(StrConv(Chr(KeyAscii), vbLowerCase))
        End If
    'This is how to restrict certain characters
    ElseIf KeyAscii = 46 And InStr(1, vsfRawScore.EditText, ".") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsfRawScore_LostFocus()
    If vsfRawScore.Rows = 2 Then
        m_bln_Edit = True
        m_int_CurrentRow = 0
        Call vsfRawScore_AfterRowColChange(vsfRawScore.Row, vsfRawScore.Col, 0, 0)
    End If
End Sub

Private Sub vsfRawScore_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    m_bln_Edit = True
    m_int_CurrentRow = vsfRawScore.Row
End Sub

Private Function f_bln_CallSP() As Boolean
    Dim l_obj_Cmd As New ADODB.Command
    Dim l_int_counter1 As Integer
    Dim l_int_Counter2 As Integer
    Dim l_int_OldCol As Integer
    Dim l_int_DataCnt As Integer
    
    On Error GoTo ErrorHandler
    Set l_obj_Cmd.ActiveConnection = g_obj_Conn
    l_obj_Cmd.CommandText = "UspCTMCalScore"
    l_obj_Cmd.CommandType = adCmdStoredProc

    l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("iExamineeProfileId", adInteger, adParamInput, 4, IIf(g_int_ExamType = 2 And m_SecondExam_Type = 1, m_int_ExamineeId(vsfRawScore.Row), -1))
    l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("SubjectProfileId", adInteger, adParamInput, 4, m_int_SubjId)

    l_int_DataCnt = 0
    For l_int_counter1 = 0 To m_int_QuestionLimit - 1
        If l_int_counter1 > 9 Then Exit For     ' exit after adding 10 question parameters
        If m_int_Score(l_int_counter1) > 0 Then l_int_DataCnt = l_int_DataCnt + 1
    Next

    l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("NumberOfParams", adInteger, adParamInput, 4, IIf(g_int_ExamType = 2 And m_SecondExam_Type = 1, cboInterviewerID.Text, l_int_DataCnt))
    ' actual question
    For l_int_counter1 = 0 To m_int_QuestionLimit - 1
        If l_int_counter1 > 9 Then Exit For     ' exit after adding 10 question parameters
        If m_int_Score(l_int_counter1) > 0 Then
            l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("Score" & str(l_int_counter1), adDouble, adParamInput, 4, m_int_Score(l_int_counter1))
        End If
    Next
    ' remaining questions out of total possible of 10
    For l_int_Counter2 = l_int_DataCnt To 9
        l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("Score" & str(l_int_Counter2), adInteger, adParamInput, 4, 0)
    Next
   
    l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("RETURN_VALUE", adDouble, adParamOutput, 4)
    
    l_obj_Cmd.Execute
        
    If Err.Number <> 0 Then
        f_bln_CallSP = False
    Else
        m_int_TotalScore = IIf(IsNull(l_obj_Cmd.Parameters("RETURN_VALUE").Value), 0, l_obj_Cmd.Parameters("RETURN_VALUE").Value)
        m_int_TotalScore = Round(m_int_TotalScore, 2)
        With vsfRawScore
            ' add the choosei score also
            l_int_OldCol = .Col
            .Col = .Cols - 1
            If Len(Trim(.Text)) = 0 Then
                m_int_ChooseiScore = 0
            Else
                m_int_ChooseiScore = Trim(.Text)
            End If
            
            .Col = 3
            If m_int_TotalScore = 0 Then
                .Text = ""
            Else
                .Text = Format(m_int_TotalScore, "##0.00")
            End If
            .Col = l_int_OldCol
        End With
        f_bln_CallSP = True
    End If
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

' to clear the grid in case there is no records found
Private Sub f_void_ClearGrid()
    Dim l_int_Counter As Integer
    With vsfRawScore
        .Rows = 2
        .Row = 0
        For l_int_Counter = 0 To .Cols - 1
            .Col = l_int_Counter
            .Text = ""
        Next
        .Row = 1
        For l_int_Counter = 0 To .Cols - 1
            .Col = l_int_Counter
            .Text = ""
        Next
        .Refresh
    End With
    
End Sub

' update data into scoreprofile table and scoredetail table
Private Function f_bln_UpdateData(ByVal rownum As Integer) As Boolean
Dim l_str_Sql As String
Dim l_obj_Rst As New ADODB.Recordset
Dim l_int_NewScoreProfileId As Integer
Dim l_int_ScoreProfileId As Integer
Dim l_int_NewScoreDetailId As Integer
Dim l_int_RawScore As Double
Dim l_int_Counter As Integer
Dim l_bln_existing As Boolean
Dim l_bln_existings() As Boolean
Dim l_int_ScoreDetailId(9) As Integer
Dim l_obj_rst1 As New ADODB.Recordset
Dim l_obj_rst2 As New ADODB.Recordset
Dim l_str_Sql1 As String
Dim l_str_Sql_sub As String
Dim l_int_Counter2 As Integer

On Error GoTo ErrorHandler

    l_int_RawScore = m_int_TotalScore   ' assign the total score calculated from SP
    ' begin the transaction
    g_obj_Conn.BeginTrans

    ' insert or update into scoreprofile table
    l_str_Sql = "SELECT iScoreProfileId FROM tbSTEScoreProfile WHERE iSubjectProfileId=" & m_int_SubjId & _
        " AND iExamineeProfileId=" & m_int_ExamineeId(rownum)
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    If Not l_obj_Rst.EOF Then
        l_bln_existing = True
        l_int_ScoreProfileId = l_obj_Rst("iScoreProfileId")
    Else
        ' ***** find the new scoreprofileid to be inserted ******
        
        Dim l_str_sql2 As String
        
        l_str_sql2 = "SELECT iScoreProfileId FROM tbSTEScoreProfile"
        l_obj_rst2.Open l_str_sql2, g_obj_Conn, adOpenStatic, adLockReadOnly
        If Not l_obj_rst2.EOF Then
            l_obj_rst2.MoveLast
            l_int_NewScoreProfileId = l_obj_rst2("iScoreProfileId") + 1
        Else
            l_str_Sql1 = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='tbSTEScoreProfile'"
            l_obj_rst1.Open l_str_Sql1, g_obj_Conn, adOpenStatic, adLockReadOnly
            If Not l_obj_rst1.EOF Then
                l_int_NewScoreProfileId = l_obj_rst1("iTableCounterIdMapping")
            Else
                l_int_NewScoreProfileId = 1
            End If
            Set l_obj_rst1 = Nothing
        End If
        ' release the object variable
        Set l_obj_rst2 = Nothing
        '***********************************************************
        
        l_int_ScoreProfileId = l_int_NewScoreProfileId
        l_bln_existing = False
    End If
    ' release the object variable
    Set l_obj_Rst = Nothing

    If l_bln_existing Then
        l_str_Sql = "UPDATE tbSTEScoreProfile SET  fRawScore=" & l_int_RawScore & _
            " , fChoseiScore=" & m_int_ChooseiScore & _
            ", dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
            " WHERE iScoreProfileId=" & l_int_ScoreProfileId
    Else
        l_str_Sql = "INSERT INTO tbSTEScoreProfile VALUES("
        l_str_Sql = l_str_Sql & l_int_NewScoreProfileId & ","
        l_str_Sql = l_str_Sql & m_int_SubjId & ","
        l_str_Sql = l_str_Sql & m_int_ExamineeId(rownum) & ","
        l_str_Sql = l_str_Sql & l_int_RawScore & ","
        l_str_Sql = l_str_Sql & m_int_ChooseiScore & ","
        l_str_Sql = l_str_Sql & "0,'"
        l_str_Sql = l_str_Sql & Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
    End If
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    ' release the object variable
    Set l_obj_Rst = Nothing

    ' insert or update into tbSTEScoreDetail table
    ReDim l_bln_existings(m_int_QuestionLimit - 1)
    l_int_NewScoreDetailId = -1
    l_str_Sql = "SELECT iScoreDetailId FROM tbSTEScoreDetail WHERE iScoreProfileId=" & l_int_ScoreProfileId

    For l_int_Counter2 = 0 To m_int_QuestionLimit - 1

        If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
            l_str_Sql_sub = " AND iSubjectQuestionId=" & cboInterviewerID.Text
        Else
            l_str_Sql_sub = " AND iSubjectQuestionId=" & m_intSubQuesProfileId(l_int_Counter2)
        End If

        Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql & l_str_Sql_sub)
        If Not l_obj_Rst.EOF Then
            l_int_Counter = 0
            Do While Not l_obj_Rst.EOF
                l_int_ScoreDetailId(l_int_Counter2) = l_obj_Rst("iScoreDetailId")
                l_obj_Rst.MoveNext
            Loop
            l_bln_existings(l_int_Counter2) = True
        Else
            l_bln_existings(l_int_Counter2) = False
            If l_int_NewScoreDetailId = -1 Then
            '***************************************************
                ' find the new scoredetailid to be inserted
                l_str_Sql1 = "SELECT iScoreDetailId FROM tbSTEScoreDetail"
                l_obj_rst2.Open l_str_Sql1, g_obj_Conn, adOpenStatic, adLockReadOnly
                If Not l_obj_rst2.EOF Then
                    l_obj_rst2.MoveLast
                    l_int_ScoreDetailId(l_int_Counter2) = l_obj_rst2("iScoreDetailId") + 1
                Else
                    l_str_Sql1 = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='tbSTEScoreDetail'"
                    l_obj_rst1.Open l_str_Sql1, g_obj_Conn, adOpenStatic, adLockReadOnly
                    If Not l_obj_rst1.EOF Then
                        l_int_ScoreDetailId(l_int_Counter2) = l_obj_rst1("iTableCounterIdMapping")
                    Else
                        l_int_ScoreDetailId(l_int_Counter2) = 1
                    End If
                    Set l_obj_rst1 = Nothing
                End If
                ' release the object variable
                Set l_obj_rst2 = Nothing
                l_int_NewScoreDetailId = l_int_ScoreDetailId(l_int_Counter2) + 1
                '*******************************************************
            Else
                l_int_ScoreDetailId(l_int_Counter2) = l_int_NewScoreDetailId
                l_int_NewScoreDetailId = l_int_ScoreDetailId(l_int_Counter2) + 1
            End If
        End If
    Next

    ' release the object variable
    Set l_obj_Rst = Nothing
    
    For l_int_Counter = 0 To m_int_QuestionLimit - 1
        If l_bln_existings(l_int_Counter) Then
            l_str_Sql = "UPDATE tbSTEScoreDetail SET "
            l_str_Sql = l_str_Sql & " fDetailScore=" & m_int_Score(l_int_Counter)
            l_str_Sql = l_str_Sql & ", dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'"
            If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
                l_str_Sql = l_str_Sql & " WHERE iSubjectQuestionId=" & cboInterviewerID.Text
            Else
                l_str_Sql = l_str_Sql & " WHERE iSubjectQuestionId=" & m_intSubQuesProfileId(l_int_Counter)
            End If
            l_str_Sql = l_str_Sql & " AND iScoreDetailId=" & l_int_ScoreDetailId(l_int_Counter)
        Else
            l_str_Sql = "INSERT INTO tbSTEScoreDetail VALUES("
            l_str_Sql = l_str_Sql & l_int_ScoreDetailId(l_int_Counter) & ","
            l_str_Sql = l_str_Sql & l_int_ScoreProfileId & ","
            If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
                l_str_Sql = l_str_Sql & cboInterviewerID.Text & ","
            Else
                l_str_Sql = l_str_Sql & m_intSubQuesProfileId(l_int_Counter) & ","
            End If
            l_str_Sql = l_str_Sql & m_int_Score(l_int_Counter) & ",'"
            l_str_Sql = l_str_Sql & Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
        End If

        Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
        If Not l_bln_existing Then
            l_int_NewScoreDetailId = l_int_NewScoreDetailId + 1
        End If
    Next
    ' release the object variable
    Set l_obj_Rst = Nothing
        
    ' if no error, then commit the transaction
    g_obj_Conn.CommitTrans
    f_bln_UpdateData = True
    
    Exit Function
ErrorHandler:
    g_obj_Conn.RollbackTrans
    MsgBox Err.Description, vbInformation, LoadResString(1729)
    f_bln_UpdateData = False
        
End Function

Private Sub f_void_InitGrid()
    
    With vsfRawScore
        .Visible = False
        .BackColor = &HFFFFFF
        .BackColorBkg = &HFFFFFF
        .BackColorFixed = &H8000000F
        .BackColorSel = &H800000
        .FixedCols = 0
        .TextStyleFixed = flexTextFlat
        .Font.Bold = False
        .ForeColorFixed = &H80000008
        .ForeColor = &H800000
        '.CellTextStyle = "0"
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .GridColor = &H808080
        .Visible = True
    End With
End Sub

Private Sub f_void_InitializeGrid()
    Dim l_int_Counter As Integer
    Dim l_int_Count As Integer
    Dim l_int_SrNo As Integer
    Dim l_int_ScoreId As Integer
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    Dim l_dbl_ChooseiScore As Double

Dim sSubList As String
Dim sSqlSubList As String
Dim lLoopCnt As Long

    On Error GoTo ErrorHandler
    Select Case g_int_ExamType
    Case 0, 1
        m_str_SQl = "SELECT iSubjectQuestionId, iSubjectProfileId, iQuestionNo, vQuestionName" & _
            " FROM tbSTESubjectQuestionProfile" & _
            " Where iSubjectProfileId = " & m_int_SelectedSubject
    Case 2
        If m_SecondExam_Type = 0 Then
            m_str_SQl = "SELECT c.iInterviewerProfileId as iSubjectQuestionId, a.iSubjectProfileId,"
            m_str_SQl = m_str_SQl & " d.vInterviewerName as vQuestionName "
            m_str_SQl = m_str_SQl & " FROM tbSTESubjectQuestionProfile a , "
            m_str_SQl = m_str_SQl & "      tbSTEInterviewRoomProfile c , "
            m_str_SQl = m_str_SQl & "      tbSTEInterviewerProfile d "
            m_str_SQl = m_str_SQl & " WHERE a.iSubjectProfileId = " & m_int_SelectedSubject
            m_str_SQl = m_str_SQl & " AND  a.iSubjectProfileId = c.iSubjectProfileId"
            m_str_SQl = m_str_SQl & " AND  c.iRoomProfileId = " & cboRoomId.Text & " "
            m_str_SQl = m_str_SQl & " AND  c.iDayFlag = " & Me.cboDay.ListIndex & " "
            m_str_SQl = m_str_SQl & " AND  d.iInterviewerProfileId = c.iInterviewerProfileId"
            m_str_SQl = m_str_SQl & " ORDER BY c.iInterviewerProfileId"
        Else
            m_str_SQl = "SELECT a.iSubjectQuestionId, a.iSubjectProfileId,"
            m_str_SQl = m_str_SQl & " a.vQuestionName "
            m_str_SQl = m_str_SQl & " FROM tbSTESubjectQuestionProfile a,"
            m_str_SQl = m_str_SQl & "      tbSTEInterviewRoomProfile c "
            m_str_SQl = m_str_SQl & " WHERE a.iSubjectProfileId = " & m_int_SelectedSubject
            m_str_SQl = m_str_SQl & " AND  a.iSubjectProfileId = c.iSubjectProfileId"
            m_str_SQl = m_str_SQl & " AND  c.iRandomNo = " & cboRoomId.Text & " "
            m_str_SQl = m_str_SQl & " AND  c.iInterviewerProfileId = " & IIf(cboInterviewerID.ListIndex = -1, "-1", cboInterviewerID.Text)
            m_str_SQl = m_str_SQl & " ORDER BY a.iSubjectQuestionId"
        End If
    End Select
    
    m_obj_Rst.Open m_str_SQl, g_obj_Conn, adOpenStatic, adLockReadOnly
    
    m_int_NoOfQues = m_obj_Rst.RecordCount
    If m_int_NoOfQues > 10 Then
        m_int_QuestionLimit = 10
    Else
        m_int_QuestionLimit = m_int_NoOfQues
    End If

    If Not m_obj_Rst.EOF Then
        lblErrorDetails.Caption = ""
        m_obj_Rst.MoveFirst
        ' store the subject ID for later use
        m_int_SubjId = m_obj_Rst("iSubjectProfileId")

        With vsfRawScore
            .Rows = 2
            .Cols = m_int_QuestionLimit + 5       'get the number of grid cols
            .FixedRows = 1
            ' header row
' Col 0:行番号 prvclNoCol = 0
' Col 1:受験番号 prvclJyukenNoCol = 1
' Col 2:受験者氏名 prvclJyukenNameCol = 2
' Col 3:受験者氏名 prvclJyukenNameCol = 3
' Col 4:総合得点 prvclSogoCol = 4
' Col 5:以降、サブジェクト これをサブジェクトの基点とする prvclSubCol = 5
            .Row = 0
            .Col = prvclNoCol
            .ColWidth(.Col) = 700
            .CellAlignment = flexAlignRightBottom

            .Text = LoadResString(1756)
            
            If .Col < .Cols - 1 Then .Col = .Col + 1
            .Text = LoadResString(1961)
            .CellAlignment = flexAlignRightBottom
            .ColWidth(.Col) = 1800
            
            If .Col < .Cols - 1 Then .Col = .Col + 1
            .Text = LoadResString(1805)
            .CellAlignment = flexAlignLeftBottom
            
            .ColWidth(.Col) = 0 'Hide initially
            If .Col < .Cols - 1 Then .Col = .Col + 1
            .Text = LoadResString(1962)
            .CellAlignment = flexAlignRightBottom
            .ColWidth(.Col) = 0
            
            If .Col < .Cols - 1 Then .Col = .Col + 1
            
            l_int_Counter = 0
            
            For l_int_Count = 0 To m_int_QuestionLimit - 1   'populate header row with names of fields

                If g_int_ExamType = 1 Or g_int_ExamType = 0 Then
                    .Text = Trim(m_obj_Rst("vQuestionName"))
                ElseIf g_int_ExamType = 2 Then
'                    .Text = Trim(m_obj_Rst("vInterviewerName"))
                    .Text = Trim(m_obj_Rst("vQuestionName"))
                End If
                .CellAlignment = flexAlignRightBottom
                .ColWidth(.Col) = 1500
                m_intSubQuesProfileId(l_int_Counter) = Trim(m_obj_Rst("iSubjectQuestionId"))
                l_int_Counter = l_int_Counter + 1
                
                ' hide all question columns, except the selected question in the combo box
'                If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
'                    If cboInterviewerID.ListIndex = l_int_Count Then
'                        .ColHidden(.Col) = False
'                    Else
'                        .ColHidden(.Col) = True
'                    End If
'                End If

                If g_int_ExamType = 0 And .Text = prvcsHyotei Then
                    '評定値ならばとなりを評定概評にする
                    .Cols = .Cols + 1
                    .Col = .Col + 1
                    .Text = prvcsSeisekiGaihyo
                    .CellAlignment = flexAlignRightBottom
                    .ColWidth(.Col) = 1500
                End If

                If .Col < .Cols - 1 Then
                    .Col = .Col + 1
                    m_obj_Rst.MoveNext
                End If

            Next
            ' last
            .Text = LoadResString(1963)
            .ColWidth(.Col) = 0
            .CellAlignment = flexAlignRightBottom
            .Refresh
        End With
       Call f_void_PopulateGrid  'Copied from GetRows
    Else
        lblErrorDetails.Caption = LoadResString(1126)
        ' release the object variables
        m_obj_Rst.Close
        Set m_obj_Rst = Nothing
        vsfRawScore.Enabled = False
        Exit Sub
    End If  ' for EOF
    
    ' release the object variables
    m_obj_Rst.Close
    Set m_obj_Rst = Nothing
    With vsfRawScore
        .Col = .Cols - 1
        .ColWidth(.Cols - 1) = IIf(chkChooseiScore.Value = 0, 0, 1800)
        If g_int_ExamType = 1 Or (g_int_ExamType = 2 And m_SecondExam_Type = 1) Then
            .Col = 1
            .ColWidth(.Col) = IIf(chkExamineeName.Value = 0, 0, 2200)
        Else
            .Col = 2
            .ColWidth(.Col) = IIf(chkExamineeName.Value = 0, 0, 2200)
        End If
        .Col = 3
        .ColWidth(.Col) = IIf(chkTotalMarks.Value = 0, 0, 1500)

'        lsvCol = -1
        For l_int_Count = 0 To .Cols - 1
            .Col = l_int_Count
            If .ColWidth(l_int_Count) > 0 Then
                If .CellBackColor = &HC0C0FF Then
'                    lsvCol = l_int_Count
                    Exit For
                End If
            End If
        Next
        If .Enabled Then .SetFocus
'        If lsvCol = -1 Then
'            .Col = 1
'        End If
    End With
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub l_void_AddRooms()

    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_Sql = "SELECT iRoomProfileId, vRoomName FROM tbSTERoomProfile" _
              & " WHERE iMaxCapacity > 0"
    l_str_Sql = l_str_Sql & " AND iInterviewRoomFlag = 0"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    If Not l_obj_Rst.EOF Then
        Do While Not l_obj_Rst.EOF
            cboRoomName.AddItem l_obj_Rst("vRoomName")
            cboRoomId.AddItem l_obj_Rst("iRoomProfileId")
            l_obj_Rst.MoveNext
        Loop
        cboRoomName.ListIndex = 0
        cboRoomId.ListIndex = 0
    Else
        lblErrorDetails.Caption = LoadResString(2010)
        Unload Me
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub l_void_AddRoomsRand()

    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_Sql = "SELECT distinct iRandomNo, iRandomNo FROM tbSTEInterviewRoomProfile" _
              & " WHERE iNendo = ( select top 1 iNendo from tbSTEsystemProfile where iActiveFlag = 1 ) " _
              & " AND iRandomNo is not null "
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    If Not l_obj_Rst.EOF Then
        Do While Not l_obj_Rst.EOF
            cboRoomName.AddItem l_obj_Rst("iRandomNo")
            cboRoomId.AddItem l_obj_Rst("iRandomNo")
            l_obj_Rst.MoveNext
        Loop
        cboRoomName.ListIndex = 0
        cboRoomId.ListIndex = 0
    Else
        lblErrorDetails.Caption = LoadResString(2010)
        Unload Me
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub l_void_PopulateDayCombo()

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

    With cboDay
        .Clear
        .AddItem LoadResString(2424)
        .AddItem LoadResString(2425)
        If bThirdDay Then .AddItem LoadResString(2426)
        .ListIndex = 0
    End With

End Sub

Private Sub f_void_populateInterviewers()
    'Populate Interviewers
    Dim l_str_Sql As String
    Dim l_int_Counter As Integer
    Dim l_obj_Rst As New ADODB.Recordset
    cboInterviewer.Clear
    cboInterviewerID.Clear
'    l_str_Sql = "SELECT a.iSubjectQuestionId, a.iSubjectProfileId," & _
        " a.iInterviewerProfileId, b.vInterviewerName FROM tbSTESubjectQuestionProfile a," & _
        " tbSTEInterviewerProfile b WHERE a.iInterviewerProfileId IN " & _
        " (SELECT iInterviewerProfileId From tbSTEInterviewRoomProfile" & _
        " WHERE iRoomProfileId = " & cboRoomId.Text & " AND iDayFlag = " & cboDay.ListIndex & " AND iSubjectProfileId = " & m_int_SelectedSubject & ")" & _
        " AND a.iSubjectProfileId = " & m_int_SelectedSubject & " AND  a.iInterviewerProfileId = b.iInterviewerProfileId" & _
        " ORDER BY a.iInterviewerProfileId"
    l_str_Sql = "SELECT "
    l_str_Sql = l_str_Sql & "  iv.iInterviewerProfileId "
    l_str_Sql = l_str_Sql & " ,iv.vInterviewerName "
    l_str_Sql = l_str_Sql & " FROM tbSTEInterviewerProfile as iv "
    l_str_Sql = l_str_Sql & " WHERE exists ( select 1 from tbSTEInterviewRoomProfile as ir "
    l_str_Sql = l_str_Sql & "                where iRandomNo = " & cboRoomId.Text
    l_str_Sql = l_str_Sql & "                and iNendo = ( select top 1 iNendo from tbSTEsystemProfile where iActiveFlag = 1 ) "
    l_str_Sql = l_str_Sql & "                and iv.iInterviewerProfileId = ir.iInterviewerProfileId )"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    Do While Not l_obj_Rst.EOF
        l_int_Counter = l_int_Counter + 1
        If l_int_Counter > 10 Then Exit Do
        cboInterviewer.AddItem l_obj_Rst("VinterviewerName")
        cboInterviewerID.AddItem l_obj_Rst("iInterviewerProfileID")
        l_obj_Rst.MoveNext
    Loop
    
    If cboInterviewer.ListCount > 0 Then
        cboInterviewer.ListIndex = 0
    End If
End Sub

Private Function gfSeisekiGaihyo(pdScore As Double) As String

    If 0 < pdScore And pdScore < 1.9 Then
        gfSeisekiGaihyo = "E"
    ElseIf 1.9 <= pdScore And pdScore < 2.7 Then
        gfSeisekiGaihyo = "D"
    ElseIf 2.7 <= pdScore And pdScore < 3.5 Then
        gfSeisekiGaihyo = "C"
    ElseIf 3.5 <= pdScore And pdScore < 4.3 Then
        gfSeisekiGaihyo = "B"
    ElseIf 4.3 <= pdScore And pdScore <= 5 Then
        gfSeisekiGaihyo = "A"
    Else
        gfSeisekiGaihyo = ""
    End If

End Function
