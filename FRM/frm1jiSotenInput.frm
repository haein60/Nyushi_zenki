VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frm1jiSotenInput 
   AutoRedraw      =   -1  'True
   ClientHeight    =   10035
   ClientLeft      =   2400
   ClientTop       =   2445
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Palette         =   "frm1jiSotenInput.frx":0000
   Picture         =   "frm1jiSotenInput.frx":3AD3
   ScaleHeight     =   10035
   ScaleWidth      =   11985
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdHaitenWari 
      Caption         =   "配点割合"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8985
      TabIndex        =   25
      Top             =   8580
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton cmdGetRows 
      Caption         =   "レコード表示"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4875
      TabIndex        =   8
      Top             =   1830
      Width           =   2895
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更  新"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4875
      TabIndex        =   10
      Top             =   8640
      Width           =   3135
   End
   Begin VB.ComboBox cboInterviewerID 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4455
      TabIndex        =   23
      Text            =   "cboInterviewerID"
      Top             =   510
      Visible         =   0   'False
      Width           =   1725
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
      Left            =   2610
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1830
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
      Height          =   360
      Left            =   10620
      TabIndex        =   4
      Top             =   810
      Visible         =   0   'False
      Width           =   1125
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
      Height          =   360
      Left            =   9195
      ScrollBars      =   1  '水平
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboSubject 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   2640
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1050
      Width           =   1830
   End
   Begin VB.ComboBox cboRoomId 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13110
      TabIndex        =   17
      Text            =   "cboRoomId"
      Top             =   615
      Visible         =   0   'False
      Width           =   1320
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
      Left            =   8925
      TabIndex        =   7
      Top             =   1515
      Visible         =   0   'False
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
      Left            =   5085
      TabIndex        =   6
      Top             =   1515
      Visible         =   0   'False
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
      TabIndex        =   5
      Top             =   1515
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.TextBox txtRandomNo 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6555
      MaxLength       =   2
      TabIndex        =   12
      Text            =   "1"
      Top             =   1065
      Width           =   1035
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfRawScore 
      Height          =   6045
      Left            =   1245
      TabIndex        =   9
      Top             =   2430
      Width           =   9675
      _cx             =   17066
      _cy             =   10663
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
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
      HighLight       =   2
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
      FormatString    =   $"frm1jiSotenInput.frx":75A6
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
      ShowComboButton =   -1  'True
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
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   7230
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   510
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblFromToCount 
      BackStyle       =   0  '透明
      Caption         =   "lblFromToCount"
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
      Height          =   330
      Left            =   8070
      TabIndex        =   24
      Top             =   1905
      Width           =   6375
   End
   Begin VB.Label lblInterviewers 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "採点者"
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
      Height          =   315
      Left            =   1695
      TabIndex        =   22
      Top             =   540
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblRoomName 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "乱数1"
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
      Height          =   360
      Left            =   6285
      TabIndex        =   21
      Top             =   510
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblJukenNoTo 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "受験番号 至"
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
      Height          =   345
      Left            =   10530
      TabIndex        =   20
      Top             =   540
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblJukenNoFrom 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "受験番号 自"
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
      Height          =   330
      Left            =   9015
      TabIndex        =   19
      Top             =   525
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "科   目"
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
      Height          =   285
      Left            =   840
      TabIndex        =   18
      Top             =   1110
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  '右揃え
      Caption         =   "平均点"
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
      Height          =   300
      Left            =   3120
      TabIndex        =   16
      Top             =   1470
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      Caption         =   "調整点"
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
      Height          =   270
      Left            =   6765
      TabIndex        =   15
      Top             =   1455
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Caption         =   "受験生氏名"
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
      Height          =   315
      Left            =   810
      TabIndex        =   14
      Top             =   1470
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lblRandomNo 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "乱数"
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
      Height          =   375
      Left            =   5205
      TabIndex        =   11
      Top             =   1110
      Width           =   1275
   End
   Begin VB.Label lblErrorDetails 
      BackStyle       =   0  '透明
      Caption         =   "lblErrorDetails"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1215
      TabIndex        =   13
      Top             =   9225
      Width           =   10515
   End
End
Attribute VB_Name = "frm1jiSotenInput"
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
'2004/12/06 D.Momoda
'評定平均と欠席日数は別Subjectである必要があり、frmRowScoreでは対応できないため別画面にした。
'**************************************************************************************************
' to store the currently selected item
Dim m_int_SelectedSubject    As Long                ' to keep track of row change
Dim m_int_CurrentRow         As Long                ' to identify when the edit starts
Dim m_bln_Edit               As Boolean             ' to store the subject profile ID to pass to stored procedure
Dim m_int_SubjId             As Long                ' to store number of questions, to pass to the stored procedure
Dim m_int_NoOfQues           As Long                ' to store the scores of diff Questions
Dim m_int_Score(9)           As String              ' to store the examinee Id
Dim m_int_ExamineeId()       As Long   'Integer     ' to store the subject question profile id
Dim m_intSubjectProfileId(9) As String              ' to store the total score
Dim m_intSubQuesProfileId(9) As String              ' to store the total score
Dim m_int_TotalScore         As Double              ' to store choosei score
Dim m_int_ChooseiScore       As Double              ' database related variables

Dim oRs                      As New ADODB.Recordset
Dim sSQL                     As String

Dim m_int_QuestionLimit      As Long

Private Const prvclNoCol     As Long = 0
Private m_SecondExam_Type    As Long '面接か小論文かflag

Private m_bDirty             As Boolean


Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Dim adoRs As New ADODB.Recordset ' レコードセット
    Dim sSQL  As String
    Dim icnt  As Long



    g_int_ExamType = 1

    m_bDirty = False

    lblErrorDetails.Caption = ""

     '**************************************************************************
     '* 試験科目 取得                                                          *
     '**************************************************************************
     sSQL = ""
     sSQL = sSQL & "SELECT" & vbCrLf
     sSQL = sSQL & "    iSubjectProfileId" & vbCrLf
     sSQL = sSQL & "   ,vSubjectName" & vbCrLf
     sSQL = sSQL & "FROM" & vbCrLf
     sSQL = sSQL & "    tbSTESubjectProfile" & vbCrLf
     sSQL = sSQL & " WHERE" & vbCrLf
     sSQL = sSQL & "     iExamType = 1"

    ''''2019.05.07 jhi 科目ID、科目名を取得するSQL文作成
     sSQL = sSQL & " ORDER BY iDispOrder"

'-------------------------------------------------------------------------------
'2021.12.17 add jhi
'-------------------------------------------------------------------------------
'SELECT
'--    iSubjectProfileId
'--   ,vSubjectName
'   *
'From
'    tbSTESubjectProfile
'Where
'    1=1
'    --and iSubType = 4
'Order By
'    iDispOrder
'-------------------------------------------------------------------------------

    Set adoRs = g_obj_Conn.Execute(sSQL)
    
    If Not adoRs.EOF Then

        'make the first subject as default selected
        m_int_SelectedSubject = adoRs("iSubjectProfileId") '30-小論文

        Do While Not adoRs.EOF
            cboSubject.AddItem adoRs("vSubjectName")                              '小論文
            cboSubject.ItemData(cboSubject.NewIndex) = adoRs("iSubjectProfileId") '30
            adoRs.MoveNext
        Loop


    End If

    ' release the object variables
    adoRs.Close
    Set adoRs = Nothing

    cboSubject.ListIndex = 0 '数学

    'Grid Setting
    Call f_void_InitGrid


    ' initialize array values to zero
    For icnt = 0 To 9
        m_int_Score(icnt) = 0
    Next

    lblFromToCount.Caption = ""

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim Index As Integer

    fMainForm.mnuTools.Enabled = False  ' disable tools menu

    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"
End Sub


Public Sub gsSetSecondType(piSType As Long)

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

''''[乱数combo]の内容をクリックした場合の処理
Private Sub cboInterviewer_Click()
    If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
        cboInterviewerID.ListIndex = cboInterviewer.ListIndex
    End If
End Sub


Private Sub cboSubject_Click()

    On Error GoTo ErrorHandler

    Dim oRs       As New ADODB.Recordset
    Dim l_str_Sql As String
    

'    l_str_Sql = "SELECT iSubjectProfileId  FROM tbSTESubjectProfile WHERE" & _
'        " vSubjectName ='" & cboSubject.Text & "'"
'    oRs.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
'    If Not oRs.EOF Then
'        m_int_SelectedSubject = oRs("iSubjectProfileId")
'    Else
'        If g_int_ExamType = 0 Then
'            m_int_SelectedSubject = -1
'        End If
'    End If
'    oRs.Close
'    Set oRs = Nothing


    If cboSubject.ListIndex >= 0 Then
        m_int_SelectedSubject = cboSubject.ItemData(cboSubject.ListIndex)
    Else
        m_int_SelectedSubject = -1
    End If

    lblFromToCount.Caption = ""
    Call f_void_ClearGrid

    If g_int_ExamType = 3 Or g_int_ExamType = 5 Then
        f_void_populateInterviewers   'Required for population of interviewers requires for Raw Score for Report
    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"    ''''LoadResString(1729)

End Sub

Private Sub cboSubject_GotFocus()

''''2021.12.22 del jhi
''''If g_int_ExamType = 0 And txtJukenNoFrom.Enabled Then
''''    txtJukenNoFrom.SetFocus
''''End If

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

'*******************************************************************************
'* 【レコード表示】 ボタン処理                                                 *
'*******************************************************************************
Private Sub cmdGetRows_Click()

    On Error GoTo ErrorHandler
    Dim l_int_Counter As Long



    If Trim(txtRandomNo.Text) = "" Then
        MsgBox "乱数番号を入力してレコードを取得してください。"
        Exit Sub
    End If

    cmdGetRows.Enabled = False
    lblErrorDetails.Caption = ""

'    If m_bDirty Then ''''false
'        If vbCancel = MsgBox("入力データが保存されていません。" & vbCrLf & "保存せずに別データを表示してもよろしいですか？", vbOKCancel) Then
'            Exit Sub
'        End If
'    End If

    m_bDirty = False

        
    ''''---------------------------------------
    '''' 2019.05.07 jhi調査
    '''' 小論文科目ID=30を取得する
    ''''SELECT
    ''''    iSubjectProfileID
    ''''FROM
    ''''    tbSTESubjectProfile
    ''''WHERE
    ''''        iExamType    = 2
    ''''    AND vSubjectName = '小論文'; --> combo科目名
    ''''---------------------------------------


    ''''comboで選択した科目のコードを取得する
    sSQL = "SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType=" & g_int_ExamType & " AND vSubjectName='" & cboSubject.Text & "'"
    Set oRs = g_obj_Conn.Execute(sSQL)

    ''''2019.05.07 add(comment) jhi 選択されたらid=30(小論文)、id=20(面接Ⅰ)などcomboで選択した科目のIDをセットする
    If Not oRs.EOF Then
        m_int_SelectedSubject = oRs("iSubjectProfileId") 'コードを変数にセット
    End If

    ' release the object variables
     oRs.Close
     Set oRs = Nothing
 

    '---------------------------------------------------------------------------
    '1次試験、素点入力処理
    '---------------------------------------------------------------------------
    vsfRawScore.Enabled = True
    cmdUpdate.Enabled = True      ''''【更新】    ボタン
    cmdHaitenWari.Enabled = True  ''''【配点割合】ボタン
        
    Call f_void_ClearGrid
    Call f_void_InitializeGrid

    If vsfRawScore.Rows > 1 Then
        lblFromToCount.Caption = Trim(str(vsfRawScore.Rows - 1)) & "件"
        
        ''''2019.05.08 add jhi 「レコードがありませんでした。」が表示するので、前のレコード件数が残る現象を直した。
        If lblErrorDetails.Caption = "レコードがありませんでした。" Then ''''LoadResString(1964) Then
            lblFromToCount.Caption = ""
        Else
            lblFromToCount.Caption = Trim(str(vsfRawScore.Rows - 1)) & "件" & " レコードを取得しました。"
        End If

    Else
        lblFromToCount.Caption = ""
    End If


    cmdGetRows.Enabled = True


    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)

End Sub

'*******************************************************************************
'* 入力素点を tbSTEScoreProfile Tableに反映する                                *
'* 2022.01.24 update jhi                                                       *
'*******************************************************************************
Private Sub cmdUpdate_Click()

    Dim l_int_Counter     As Long
    Dim l_int_LoopCounter As Long
    Dim l_bln_Update      As Boolean
    Dim l_str_ErrString   As String
    
    Dim SQL               As String
    Dim RS                As New ADODB.Recordset
    Dim intLockFlag       As Integer

    Dim i As Long
    Dim k As Long
    Dim s As Long

    On Error GoTo ErrorHandler



START_00:

    lblErrorDetails.Caption = ""
    Screen.MousePointer = vbHourglass
    DoEvents
''''Sleep 10000


    l_bln_Update = True

'このあとのRowcolchangeなどでトータルスコアのカラムのデータが更新されている。
'よってm_bln_Editを変更してはならない


'2017/12/11,S-------
'同時更新　待ち

    cmdUpdate.Enabled = False

    'Update時 Table lock状態をチェックする
    intLockFlag = 0
    SQL = "SELECT ISNULL(iLocks,0) iLocks FROM tbSTELocks WITH(NOLOCK) WHERE vTarget = 'tbSTEScoreProfile' "
    Set RS = g_obj_Conn.Execute(SQL)

     If Not RS.EOF Then
        intLockFlag = RS.Fields(0).Value
     End If

    RS.Close
    Set RS = Nothing


    Sleep (1)

    If intLockFlag = 1 Then
''''    For i = 1 To 10
        For i = 1 To 100
            For k = 1 To 1000
                s = k
            Next k
        Next i

''''    Call cmdUpdate_Click    ''''2022.01.24 add jhi

        Sleep 2000
        DoEvents
        GoTo START_00
        
    End If
'2017/12/11,E-------


    With vsfRawScore
    .Row = 1
    
    For l_int_LoopCounter = 1 To .Rows - 1
        .Row = l_int_LoopCounter
        .Col = 3

        If Trim(.Text) = "" Then
            m_int_TotalScore = 0
        Else
            If Trim(.Text) = "a" Then
                m_int_TotalScore = -1
            ElseIf Trim(.Text) = "b" Then
                m_int_TotalScore = -2
            Else
                m_int_TotalScore = CDbl(.Text)
            End If
        End If

        .Col = 4
        For l_int_Counter = 0 To m_int_QuestionLimit - 1
            If Trim(.Text) = "a" Then
                m_int_Score(l_int_Counter) = -1
            ElseIf Trim(.Text) = "b" Then
                m_int_Score(l_int_Counter) = -2
            Else
                If Trim(.Text) = "" Then
                    m_int_Score(l_int_Counter) = 0
                Else
                    m_int_Score(l_int_Counter) = Trim(.Text)
                End If
            End If

            If .TextMatrix(0, .Col) = gcsHyotei Then
                .Col = .Col + 2
            Else
                .Col = .Col + 1
            End If
        Next

        If Trim(.Text) = "" Then
            m_int_ChooseiScore = 0
        Else
            m_int_ChooseiScore = .Text
        End If
        'Add,xzg,2016/12/19,S-----

        '固定０に設定
        m_int_ChooseiScore = 0
        'Add,xzg,2016/12/19,E-----


        '-----------------------------------------------------------------------
        ' 入力の素点をupdate or insert する
        '-----------------------------------------------------------------------
        l_bln_Update = f_bln_UpdateData(l_int_LoopCounter)

        If Not l_bln_Update Then
            l_str_ErrString = l_str_ErrString & CStr(l_int_LoopCounter) & ","
        End If

    Next


    If Not l_bln_Update Then
        l_str_ErrString = Left(l_str_ErrString, Len(l_str_ErrString) - 1)
''''    lblErrorDetails = LoadResString(2437) & l_str_ErrString ''''2022.0124 del jhi
        lblErrorDetails.Caption = "以下の列で入力データでエラーが発生しました。(" & l_str_ErrString & ")"
    Else
        m_bDirty = False
        lblErrorDetails.Caption = "入力素点を正常にDBに反映しました。" '''LoadResString(2404)
    End If

    End With


    cmdUpdate.Enabled = True
    Screen.MousePointer = vbDefault


    Exit Sub


ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub


Private Sub f_void_PopulateGrid()

    Dim l_int_Counter As Long
    Dim l_int_SrNo As Long

'update,xzg,2007/02/14,S-----------
    'Dim l_int_ScoreId As Long
    Dim l_int_ScoreId As Long
'update,xzg,2007/02/14,E-----------

    Dim l_str_Sql As String
    Dim oRs As New ADODB.Recordset
    Dim l_dbl_ChooseiScore As Double
    Dim l_int_StrLen As Long
    Dim l_int_Cnt As Long

    Dim l_obj_Populate  As New ADODB.Recordset

    Dim l_int_CurYear As Long
    Dim l_str_SqlDay As String

    Dim l_obj_rstDay As New ADODB.Recordset
    Dim f_dt_SourceDay As Date
    Dim f_int_SourceDayMax As Long
    Dim l_int_NoOfRooms As Long
    Dim lCol As Long
    Dim lsvCol As Long
    Dim l_int_counter2 As Long

    Dim bChecked As Boolean

    Dim errLine As String

    On Error GoTo ErrorHandler
    

errLine = "1"
    l_int_CurYear = g_int_CurrentNendo  'global variable in form load
    

errLine = "3"

   If Len(Trim(txtRandomNo.Text)) <> 0 Then

   End If

 
    If g_int_ExamType = 1 Then

'入試実施時の不具合No11対応  2004/01/24
        sSQL = "SELECT e.iExamineeProfileId, dbo.usfMakeDispJukenNumber(e.iJukenNumber) as iJukenNumber ,e.vExamineeName " & _
            " FROM tbSTEExamineeProfile e,tbSTERoomProfile r " & _
            " WHERE r.iRoomProfileId = e.iRoomProfileId " & _
            " AND e.iNendo = " & l_int_CurYear & _
            " AND r.iRandom =" & txtRandomNo.Text & _
            " AND e.iExamineeStatus = " & gclExamineeStatus_Default

        sSQL = sSQL & " AND not exists ( select 1 from tbSTEScoreProfile as s , tbSTESubjectProfile as su where s.iExamineeProfileID = e.iExamineeProfileID "
        sSQL = sSQL & " AND su.iSubjectProfileID = s.iSubjectProfileID "
        sSQL = sSQL & " AND su.vSubjectName = '" & cboSubject.Text & "' "
        sSQL = sSQL & " AND s.iAbsentFlag = 1 ) "

        ' comdesign
        '
        '
        '
            Select Case Trim(cboSubject.Text)
            Case "数学"
                ''''sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default ''''2022.01.24 del jhi 意味なし
            Case "英語"
                sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=" & m_int_SubjId
            Case "物理"
                sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
            Case "化学"
                sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
            Case "生物"
                sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
            End Select

        End If

        
'2019.05.07 jhi comment文入れた: not exist:親表の中、駆動表にないレコードのみ返す
''''        sSQL = sSQL & " AND not exists ( select 1 from tbSTEScoreProfile as s , tbSTESubjectProfile as su where s.iExamineeProfileID = tbSTEExamineeProfile.iExamineeProfileID "

'2019.05.07 jhi for 確認(not除いて existsにしてみた):exist:親表の中、駆動表にあるレコードのみ返す
''''    sSQL = sSQL & " AND     exists ( select 1 from tbSTEScoreProfile as s , tbSTESubjectProfile as su where s.iExamineeProfileID = tbSTEExamineeProfile.iExamineeProfileID "
        
''''2022.01.24 del jhi 上と同じだったので
''''        sSQL = sSQL & " AND su.iSubjectProfileID = s.iSubjectProfileID "
''''        sSQL = sSQL & " AND su.vSubjectName = '" & cboSubject.Text & "' "
''''        sSQL = sSQL & " AND s.iAbsentFlag = 1 ) "
''''
''''        If Trim(txtJukenNoFrom.Text) <> "" And Trim(txtJukenNoTo.Text) <> "" Then
''''            sSQL = sSQL & " AND iJukenNumber between " & txtJukenNoFrom.Text & " AND " & txtJukenNoTo.Text
''''        End If
 

    sSQL = sSQL & " order by iJukenNumber "


'2021.12.10 add jhi
'Select
'    iExamineeProfileId
'   ,dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber
'   ,vExamineeName
'From
'    tbSTEExamineeProfile
'Where
'       iNendo = 2020
'    AND iExamineeStatus = 1
'    AND iShoronbunRandomNo = 1
'    AND exists ( SELECT 1 FROM tbSTEInterviewRoomProfile as ir  WHERE ir.iRandomNo = 1  AND ir.iInterviewerProfileId = 125 )
'    AND not exists ( select 1 from tbSTEScoreProfile as s , tbSTESubjectProfile as su where s.iExamineeProfileID = tbSTEExamineeProfile.iExamineeProfileID  AND su.iSubjectProfileID = s.iSubjectProfileID  AND su.vSubjectName = '小論文'  AND s.iAbsentFlag = 1 )
'order by iJukenNumber


    l_obj_Populate.Open sSQL, g_obj_Conn, adOpenStatic, adLockReadOnly

errLine = "4"
    
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
errLine = "41"
                m_int_ExamineeId(l_int_SrNo) = l_obj_Populate("iExamineeProfileId")
                .Row = l_obj_Populate.AbsolutePosition
                .CellPictureAlignment = flexAlignCenterCenter
errLine = "42"
                .Col = 0
                .Text = l_int_SrNo
                .Col = .Col + 1

                .Text = l_obj_Populate("iJukenNumber")

errLine = "43"
                .Col = .Col + 1
errLine = "44"
                If g_int_ExamType <> 0 Then
                    l_int_StrLen = Len(l_obj_Populate("vExamineeName"))
errLine = "45"
                    For l_int_Cnt = 0 To l_int_StrLen - 1
                        .Text = .Text & "*"
                    Next
errLine = "46"
                Else
                    .Text = l_obj_Populate("vExamineeName")
                End If
errLine = "47"
                .Col = .Col + 1

errLine = "48"
                l_str_Sql = "SELECT s.iScoreprofileId, s.fRawScore, s.fChoseiScore"
                l_str_Sql = l_str_Sql & " FROM tbSTEScoreProfile s, tbSTEExamineeProfile e"
                If g_int_ExamType = 0 Then
                    l_str_Sql = l_str_Sql & " WHERE exists ( select 1 from tbSTESubjectProfile as sp where sp.iSubType = 0 and sp.iSubjectProfileID = s.iSubjectProfileId ) "
                Else
                    l_str_Sql = l_str_Sql & " WHERE s.iSubjectProfileId=" & m_int_SubjId
                End If
                l_str_Sql = l_str_Sql & " AND e.iExamineeProfileId=" & m_int_ExamineeId(l_int_SrNo)
                l_str_Sql = l_str_Sql & " AND e.iExamineeProfileId = s.iExamineeProfileId"

'---------------------------------------------------
'2021.12.10 add jhi
'SELECT
'    s.iScoreprofileId
'   ,s.fRawScore
'   ,s.fChoseiScore
'From
'    tbSTEScoreProfile s
'   ,tbSTEExamineeProfile e
'Where
'        s.iSubjectProfileID = 30
'    AND e.iExamineeProfileId=41693
'    AND e.iExamineeProfileId = s.iExamineeProfileId

'SELECT
'   -- s.iScoreprofileId
'   --,s.fRawScore
'   --,s.fChoseiScore
'   s.*
'From
'    tbSTEScoreProfile s
'   ,tbSTEExamineeProfile e
'Where
'        s.iSubjectProfileID = 30
'--    AND e.iExamineeProfileId=41693
'    AND e.iExamineeProfileId = s.iExamineeProfileId
'    and substring(convert(varchar,s.dtUpdate,23),1,4) = '2020'
'order by iScoreProfileId
    

'---------------------------------------------------


                Set oRs = g_obj_Conn.Execute(l_str_Sql)
errLine = "5"
                If Not oRs.EOF Then
                    If Trim(oRs("fRawScore")) = "-1" Then
                        .Text = "a"
                    ElseIf Trim(oRs("fRawScore")) = "-2" Then
                        .Text = "b"
                    ElseIf Trim(oRs("fRawScore")) <> "" Then
                        If .TextMatrix(0, .Col) = gcsKessekiNissuu Then
                            .Text = oRs("fRawScore")
                        Else
                            .Text = Format(oRs("fRawScore"), "##0.0")  '小数点以下は常に２位まで
                        End If
                        If .Text = "0" Then
                            'del.2014/12
                            '.Text = ""
                        End If
                    Else
                        .Text = ""
                    End If
                    l_int_ScoreId = oRs("iScoreProfileId")
                    If Trim(oRs("fChoseiScore")) <> "" Then
                        l_dbl_ChooseiScore = oRs("fChoseiScore")
                    Else
                        l_dbl_ChooseiScore = 0
                    End If
                Else
                    .Text = ""
                End If
                
                ' release the object variable
                Set oRs = Nothing

                l_str_Sql = "SELECT d.fDetailScore , s.iSubjectProfileId , d.iSubjectQuestionId FROM tbSTEScoreDetail d, tbSTEScoreProfile s "
'                l_str_Sql = l_str_Sql & " WHERE d.iScoreProfileId = " & l_int_ScoreId
                l_str_Sql = l_str_Sql & " WHERE s.iExamineeProfileId = " & m_int_ExamineeId(l_int_SrNo)
                If g_int_ExamType = 0 Then
                    l_str_Sql = l_str_Sql & " AND exists ( select 1 from tbSTESubjectProfile as sp where sp.iSubType = 0 and sp.iSubjectProfileID = s.iSubjectProfileId ) "
                Else
                    l_str_Sql = l_str_Sql & " AND s.iSubjectProfileId=" & m_int_SubjId
                End If
                l_str_Sql = l_str_Sql & " AND s.iScoreProfileId = d.iScoreProfileId"
                If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
                    l_str_Sql = l_str_Sql & " AND d.iSubjectQuestionId = " & cboInterviewerID.Text
                End If
                l_str_Sql = l_str_Sql & " ORDER BY s.iSubjectProfileId , d.iSubjectQuestionId "

                Set oRs = g_obj_Conn.Execute(l_str_Sql)

'---------------------------------------------------
'2021.12.10 add jhi
'SELECT
'    d.fDetailScore
'   ,s.iSubjectProfileId
'   ,d.iSubjectQuestionId
'From
'    tbSTEScoreDetail d
'   ,tbSTEScoreProfile s
'Where
'        s.iExamineeProfileId = 41693
'    AND s.iSubjectProfileId=30
'    AND s.iScoreProfileId = d.iScoreProfileId
'    AND d.iSubjectQuestionId = 125
'Order By
'    s.iSubjectProfileID
'   ,d.iSubjectQuestionId
'---------------------------------------------------


errLine = "6"
                ' set the question limit
                
                If m_int_NoOfQues <= 10 Then
                    m_int_QuestionLimit = m_int_NoOfQues
                    lblErrorDetails.Caption = ""
''''                lblErrorDetails.Visible = False    ''''2022.01.24 del jhi
                Else
                    m_int_QuestionLimit = 10
''''                lblErrorDetails.Caption = LoadResString(2500) & m_int_NoOfQues & " " & LoadResString(2501)
                    lblErrorDetails.Caption = "表示可能な質問数は１０個までです。" & m_int_NoOfQues & " " & LoadResString(2501)
                    lblErrorDetails.Visible = True
                End If

errLine = "61"
                For l_int_Counter = 1 To m_int_QuestionLimit
                   ' this will initialize chosei score also to zero
                   If Not oRs.EOF Then
                        For l_int_counter2 = 0 To m_int_QuestionLimit - 1
                            If m_intSubQuesProfileId(l_int_counter2) = oRs("iSubjectQuestionId") Then
                                If m_intSubjectProfileId(l_int_counter2) = oRs("iSubjectProfileId") Then
                                    If g_int_ExamType = 0 And .TextMatrix(0, l_int_counter2 + 4) <> gcsHyotei Then
                                        .Col = l_int_counter2 + 5
                                    Else
                                        .Col = l_int_counter2 + 4
                                    End If
                                    bChecked = True
                                    Exit For
                                End If
                            End If
                        Next
errLine = "62"
                        If .TextMatrix(0, .Col) <> gcsSeisekiGaihyo Then
                            .CellBackColor = &HC0C0FF
                        End If

                        If bChecked And Trim(oRs("fDetailScore")) <> "" Then
                            If oRs("fDetailScore") = 0 Then
                                'update,2014/12
                                '.Text = ""
                                .Text = "0.0"
                            ElseIf Trim(oRs("fDetailScore")) = "-1" Then
                                If .TextMatrix(0, .Col) = gcsHyotei Then
                                    .Text = ""
                                Else
                                    .Text = "a"
                                End If
                            ElseIf Trim(oRs("fDetailScore")) = "-2" Then
                                If .TextMatrix(0, .Col) = gcsHyotei Then
                                    .Text = ""
                                Else
                                    .Text = "b"
                                End If
                            Else
                                If .TextMatrix(0, .Col) = gcsKessekiNissuu Then
                                    .Text = oRs("fDetailScore")
                                Else
                                    .Text = Format(oRs("fDetailScore"), "##0.0") '小数点以下は常に２位まで
                                End If
                            End If
                        Else
                            .Text = ""
                        End If
errLine = "63"
                        If .TextMatrix(0, .Col) = gcsHyotei Then
                            .Col = .Col + 1
                            .Text = gfSeisekiGaihyo(oRs("fDetailScore"))
                        End If
errLine = "64"
                        oRs.MoveNext
                   Else
                        .Col = .Col + 1
                        If .TextMatrix(0, .Col) <> gcsSeisekiGaihyo Then
                            .CellBackColor = &HC0C0FF
                        End If
                        .Text = ""
                    End If
                Next

errLine = "7"
                ' release the object variable
                Set oRs = Nothing
               
                .Col = .Col + 1
                If .TextMatrix(0, .Col) <> gcsSeisekiGaihyo Then
                    .CellBackColor = &HC0C0FF
                End If
                If l_dbl_ChooseiScore = 0 Then
                    'del,2014/12
                   '.Text = ""
                Else
                   .Text = l_dbl_ChooseiScore
                End If
errLine = "71"
                .Rows = .Rows + 1                    'add a new row to the grid
                l_obj_Populate.MoveNext
errLine = "72"
            Loop

errLine = "8"
'           .TextMatrix(0, 1000) = "test"
           
            .Rows = .Rows - 1                       'remove the last row because it's blank
            .Row = 1
        End With
    Else
        cmdUpdate.Enabled = False
        lblErrorDetails.Caption = "レコードがありませんでした。" ''''LoadResString(1964)
        lblErrorDetails.Visible = True

        Call f_void_ClearGrid
        vsfRawScore.Enabled = False

    End If

errLine = "9"

    ' release the object variable
    l_obj_Populate.Close
    Set l_obj_Populate = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"
    MsgBox "Err Line " & errLine
    
End Sub

Private Sub optRandom_Click()
    cboSubject.Enabled = False
End Sub

Private Sub optSubject_Click()
    txtRandomNo.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If m_bDirty Then
        If vbCancel = MsgBox("入力後、保存されていません。" & vbCrLf & "保存せず終了してもよろしいですか？", vbOKCancel) Then
            Cancel = 1
        Else
            Call g_void_CloseChildForm
        End If
    Else
        Call g_void_CloseChildForm
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

Private Sub txtRandomNo_KeyPress(KeyAscii As Integer)

    lblFromToCount.Caption = ""  ''''2022.01.20 add jhi recode 取得メッセージclear
    lblErrorDetails.Caption = "" ''''2022.01.24 add jhi msg clear
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub vsfRawScore_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    With vsfRawScore
        If .Col > 3 Then
            If Trim(.TextMatrix(Row, Col)) <> "" Then
                If IsNumeric(.TextMatrix(Row, Col)) Then
                    If .TextMatrix(0, .Col) <> gcsKessekiNissuu Then
                        .TextMatrix(Row, Col) = Format(Round(.TextMatrix(Row, Col), 2), "##0.0") '小数点以下は常に２位まで
                    End If
                    If .TextMatrix(0, Col) = gcsHyotei Then
                        .TextMatrix(Row, Col + 1) = gfSeisekiGaihyo(.TextMatrix(Row, Col))
                    End If
                    If .TextMatrix(Row, Col) = 0 Then
                        'del,2014/12
                       ' .TextMatrix(Row, Col) = ""
                    End If
                End If
                ' change in comdesign, arka 19apr 2002 end
'NextCol:
'                If .Col < .Cols - 1 Then
'                    If .TextMatrix(0, Col) = gcsHyotei Then
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
                If .TextMatrix(0, Col) = gcsHyotei Then
                    .TextMatrix(Row, Col + 1) = ""
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfRawScore_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)

    Dim l_bln_Sp As Boolean
    Dim l_int_Counter As Long
    Dim l_bln_Update As Boolean
    Dim l_int_EditedRow As Long

    On Error GoTo ErrorHandler
    
    If vsfRawScore.Row <> m_int_CurrentRow Then
        If m_bln_Edit Then
            m_bln_Edit = False
            With vsfRawScore
                If OldRow = 0 Then Exit Sub
                .Row = OldRow
                l_int_EditedRow = .Row
                .Col = 4
                For l_int_Counter = 0 To m_int_QuestionLimit - 1
                    If Len(Trim(.Text)) <> 0 Then
                        If .TextMatrix(0, .Col) = gcsKessekiNissuu Then
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
                    If .TextMatrix(0, .Col) = gcsHyotei Then
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
'                    If .TextMatrix(0, .Col) <> gcsKessekiNissuu And .TextMatrix(0, .Col) <> gcsSeisekiGaihyo Then
                    If g_int_ExamType <> 0 Then
                        l_bln_Sp = f_bln_CallSP()
                        If Not l_bln_Sp Then
''''                        lblErrorDetails.Caption = LoadResString(1956) ''''2022.01.24 del jhi
                            lblErrorDetails.Caption = "Sp実行中にエラーが発生しました。" ''''LoadResString(1956)
                            Exit Sub
                        Else
                            lblErrorDetails.Caption = ""
                            m_bln_Edit = False
                        End If
                    End If
                End If
                .Row = NewRow
                .Col = NewCol
            End With
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

Private Sub vsfRawScore_Click()
    vsfRawScore.EditCell
End Sub

Private Sub vsfRawScore_KeyPress(KeyAscii As Integer)
    Call vsfRawScore_KeyPressEdit(vsfRawScore.Row, vsfRawScore.Col, KeyAscii)
End Sub

Private Sub vsfRawScore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

    If Col < 4 Then
        KeyAscii = 0
    ElseIf vsfRawScore.TextMatrix(0, vsfRawScore.Col) = gcsSeisekiGaihyo Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
'm_bln_Edit = False
        If g_int_ExamType = 2 And m_SecondExam_Type = 0 Then
        '縦移動
            If vsfRawScore.Row < vsfRawScore.Rows - 1 Then
                vsfRawScore.Row = vsfRawScore.Row + 1
            Else
NextCol1:
                If vsfRawScore.Col < vsfRawScore.cols - 1 Then
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
            If vsfRawScore.Col < vsfRawScore.cols - 1 Then
                If vsfRawScore.TextMatrix(0, vsfRawScore.Col) = gcsHyotei Then
                    If vsfRawScore.Col < vsfRawScore.cols - 2 Then
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
                        m_bln_Edit = True
                        Call vsfRawScore_AfterRowColChange(Row, Col, 1, 4)
'                        vsfRawScore.Row = 1
'                        vsfRawScore.Col = 4
                        cmdUpdate.SetFocus
                    End If
                End If
            End If
        End If
'm_bln_Edit = True
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab And Not (vsfRawScore.TextMatrix(0, vsfRawScore.Col) <> gcsKessekiNissuu And KeyAscii = 46) Then
        If Not ((KeyAscii = Asc("A") Or KeyAscii = Asc("B") Or KeyAscii = Asc("a") Or KeyAscii = Asc("b")) And vsfRawScore.TextMatrix(0, vsfRawScore.Col) = gcsKessekiNissuu) Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(StrConv(Chr(KeyAscii), vbLowerCase))
            m_bDirty = True
        End If
    'This is how to restrict certain characters
    ElseIf KeyAscii = 46 And InStr(1, vsfRawScore.EditText, ".") > 0 Then
        KeyAscii = 0
    Else
        m_bDirty = True
    End If

End Sub

Private Sub vsfRawScore_LostFocus()
    If vsfRawScore.Rows = 2 Then
        m_bln_Edit = True
        m_int_CurrentRow = 0
        Call vsfRawScore_AfterRowColChange(vsfRawScore.Row, vsfRawScore.Col, 0, 0)
    End If
'    If vsfRawScore.Row = vsfRawScore.Rows - 1 Then
'        Call vsfRawScore_AfterRowColChange(vsfRawScore.Row, vsfRawScore.Col, 0, 0)
'    End If
End Sub

Private Sub vsfRawScore_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    m_bln_Edit = True
    m_int_CurrentRow = vsfRawScore.Row
End Sub

Private Function f_bln_CallSP() As Boolean

    Dim l_obj_Cmd As New ADODB.Command
    Dim l_int_counter1 As Long
    Dim l_int_counter2 As Long
    Dim l_int_OldCol As Long
    Dim l_int_DataCnt As Long
    Dim l_int_DataCnt1 As Long
        
    On Error GoTo ErrorHandler

    Set l_obj_Cmd.ActiveConnection = g_obj_Conn

    l_obj_Cmd.CommandText = "UspCTMCalScore"
    l_obj_Cmd.CommandType = adCmdStoredProc

    l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("iExamineeProfileId", adInteger, adParamInput, 4, IIf(g_int_ExamType = 2 And m_SecondExam_Type = 1, m_int_ExamineeId(vsfRawScore.Row), -1))
    l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("SubjectProfileId", adInteger, adParamInput, 4, m_int_SubjId)

    l_int_DataCnt = 0
    For l_int_counter1 = 0 To m_int_QuestionLimit - 1
        If l_int_counter1 > 9 Then Exit For     ' exit after adding 10 question parameters
        
        'udate,xzg,2017/12/11,S--------
        '2017 面接０でも平均点計算できます。
         If m_int_Score(l_int_counter1) > 0 Then l_int_DataCnt = l_int_DataCnt + 1
         
        If g_int_ExamType = 2 And m_SecondExam_Type = 0 Then
            If m_int_Score(l_int_counter1) > -1 Then l_int_DataCnt1 = l_int_DataCnt1 + 1
        Else
            If m_int_Score(l_int_counter1) > 0 Then l_int_DataCnt1 = l_int_DataCnt1 + 1
        End If
        'udate,xzg,2017/12/11,E--------
    Next

    l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("NumberOfParams", adInteger, adParamInput, 4, IIf(g_int_ExamType = 2 And m_SecondExam_Type = 1, cboInterviewerID.Text, l_int_DataCnt1))
    ' actual question
    For l_int_counter1 = 0 To m_int_QuestionLimit - 1
        If l_int_counter1 > 9 Then Exit For     ' exit after adding 10 question parameters
        
        If m_int_Score(l_int_counter1) > 0 Then
            l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("Score" & str(l_int_counter1), adDouble, adParamInput, 4, m_int_Score(l_int_counter1))
        End If
    Next
    ' remaining questions out of total possible of 10
    For l_int_counter2 = l_int_DataCnt To 9
        l_obj_Cmd.Parameters.Append l_obj_Cmd.CreateParameter("Score" & str(l_int_counter2), adInteger, adParamInput, 4, 0)
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
            .Col = .cols - 1
            If Len(Trim(.Text)) = 0 Then
                m_int_ChooseiScore = 0
            Else
                m_int_ChooseiScore = Trim(.Text)
            End If
            
            .Col = 3
            If m_int_TotalScore = 0 Then
                .Text = ""
            Else
                .Text = Format(m_int_TotalScore, "##0.0")
            End If
            .Col = l_int_OldCol
        End With
        f_bln_CallSP = True
    End If

    Exit Function

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)

End Function

' to clear the grid in case there is no records found
Private Sub f_void_ClearGrid()

    Dim i As Long

    With vsfRawScore
        .Rows = 2
        .Row = 0

        For i = 0 To .cols - 1
            .Col = i
            .Text = ""
        Next i

        .Row = 1
        For i = 0 To .cols - 1
            .Col = i
            .Text = ""
        Next i

        .Refresh

    End With
    
End Sub


' update data into scoreprofile table and scoredetail table
Private Function f_bln_UpdateData(ByVal rownum As Long) As Boolean

Dim l_str_Sql As String
Dim oRs As New ADODB.Recordset
Dim l_int_NewScoreProfileId As Long
Dim l_int_ScoreProfileId(1) As Long
Dim l_int_NewScoreDetailId As Long
Dim l_int_RawScore As Double
Dim l_int_Counter As Long
Dim l_bln_existing(1) As Boolean
Dim l_bln_existings() As Boolean
Dim l_int_ScoreDetailId(9) As Long
Dim oRs1 As New ADODB.Recordset
Dim oRs2 As New ADODB.Recordset
Dim l_str_Sql1 As String
Dim l_str_Sql_sub As String
Dim l_int_counter2 As Long
Dim bRtn As Boolean

On Error GoTo ErrorHandler

    l_int_RawScore = m_int_TotalScore   ' assign the total score calculated from SP


    ' begin the transaction
    g_obj_Conn.BeginTrans


    l_str_Sql = "Update tbSTELocks set iLocks = 1 where vTarget = 'tbSTEScoreProfile' "
    Call g_obj_Conn.Execute(l_str_Sql)

    ' insert or update into scoreprofile table
    l_str_Sql = "SELECT iScoreProfileId FROM tbSTEScoreProfile "

    If g_int_ExamType = 0 Then
        l_str_Sql = l_str_Sql & " WHERE exists ( select 1 from tbSTESubjectProfile as sp where iSubType = 0 and sp.iSubjectProfileId = tbSTEScoreProfile.iSubjectProfileId ) "
    Else
        l_str_Sql = l_str_Sql & " WHERE iSubjectProfileId=" & m_int_SubjId
    End If

    l_str_Sql = l_str_Sql & " AND iExamineeProfileId=" & m_int_ExamineeId(rownum)
    l_str_Sql = l_str_Sql & " ORDER BY iSubjectProfileId "

    Set oRs = g_obj_Conn.Execute(l_str_Sql)

    If Not oRs.EOF Then
        l_bln_existing(0) = True
        l_int_ScoreProfileId(0) = oRs("iScoreProfileId")
        oRs.MoveNext
        If oRs.EOF Then
            If g_int_ExamType = 0 Then
                l_bln_existing(1) = False
                bRtn = getNewId("tbSTEScoreProfile", "iScoreProfileId", l_int_NewScoreProfileId)
                l_int_ScoreProfileId(1) = l_int_NewScoreProfileId
            Else
                l_bln_existing(1) = False
                l_int_ScoreProfileId(1) = -1
            End If
        Else
            l_bln_existing(1) = True
            l_int_ScoreProfileId(1) = oRs("iScoreProfileId")
        End If
    Else
        bRtn = getNewId("tbSTEScoreProfile", "iScoreProfileId", l_int_NewScoreProfileId)
        l_int_ScoreProfileId(0) = l_int_NewScoreProfileId
        l_bln_existing(0) = False
        If g_int_ExamType = 0 Then
            l_int_ScoreProfileId(1) = l_int_NewScoreProfileId + 1
            l_bln_existing(1) = False
        Else
            l_int_ScoreProfileId(1) = -1
            l_bln_existing(1) = False
        End If
    End If

    ' release the object variable
    Set oRs = Nothing

'欠席日数がa,bのとき、評定値が入らないため（＝０で登録される）トリガで-1、-2のデータが入ってきたら評定値も同じ値で更新するようにした。
'↑やめた

    If g_int_ExamType = 0 Then

        For l_int_counter2 = 0 To 1

            If l_int_counter2 = 0 And m_int_Score(l_int_counter2) = 0 And m_int_Score(1) < 0 Then
                '評定値登録時、評定値が0で欠席日数がa,bのとき
                m_int_Score(0) = m_int_Score(1)
            End If

            If l_bln_existing(l_int_counter2) Then
                l_str_Sql = "UPDATE tbSTEScoreProfile SET  fRawScore=" & m_int_Score(l_int_counter2)
                l_str_Sql = l_str_Sql & ", dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'"
                l_str_Sql = l_str_Sql & " WHERE iScoreProfileId=" & l_int_ScoreProfileId(l_int_counter2)
            Else
                l_str_Sql = "INSERT INTO tbSTEScoreProfile VALUES("
                l_str_Sql = l_str_Sql & l_int_ScoreProfileId(l_int_counter2) & ","
                l_str_Sql = l_str_Sql & m_intSubjectProfileId(l_int_counter2) & ","
                l_str_Sql = l_str_Sql & m_int_ExamineeId(rownum) & ","
                l_str_Sql = l_str_Sql & m_int_Score(l_int_counter2) & ","
                l_str_Sql = l_str_Sql & m_int_ChooseiScore & ","
                l_str_Sql = l_str_Sql & "0,'"
                l_str_Sql = l_str_Sql & Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
            End If

            Set oRs = g_obj_Conn.Execute(l_str_Sql)
            ' release the object variable
            Set oRs = Nothing
        Next

    Else
        If l_bln_existing(0) Then
            l_str_Sql = "UPDATE tbSTEScoreProfile SET  fRawScore=" & l_int_RawScore
            l_str_Sql = l_str_Sql & ", dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'"
            l_str_Sql = l_str_Sql & " WHERE iScoreProfileId=" & l_int_ScoreProfileId(0)
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

        Set oRs = g_obj_Conn.Execute(l_str_Sql)
        Set oRs = Nothing

    End If

    ' insert or update into tbSTEScoreDetail table
    ReDim l_bln_existings(m_int_QuestionLimit - 1)
    l_int_NewScoreDetailId = -1

    For l_int_counter2 = 0 To m_int_QuestionLimit - 1

        l_str_Sql = "SELECT iScoreDetailId FROM tbSTEScoreDetail "
        If g_int_ExamType = 0 Then
            l_str_Sql = l_str_Sql & "WHERE iScoreProfileId=" & l_int_ScoreProfileId(l_int_counter2)
        Else
            l_str_Sql = l_str_Sql & "WHERE iScoreProfileId=" & l_int_ScoreProfileId(0)
        End If

        If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
            l_str_Sql = l_str_Sql & " AND iSubjectQuestionId=" & cboInterviewerID.Text
        Else
            l_str_Sql = l_str_Sql & " AND iSubjectQuestionId=" & m_intSubQuesProfileId(l_int_counter2)
        End If

        Set oRs = g_obj_Conn.Execute(l_str_Sql)
        If Not oRs.EOF Then
            l_int_Counter = 0
            Do While Not oRs.EOF
                l_int_ScoreDetailId(l_int_counter2) = oRs("iScoreDetailId")
                oRs.MoveNext
            Loop
            l_bln_existings(l_int_counter2) = True
        Else
            l_bln_existings(l_int_counter2) = False
            If l_int_NewScoreDetailId = -1 Then
            '***************************************************
                ' find the new scoredetailid to be inserted
                bRtn = getNewId("tbSTEScoreDetail", "iScoreDetailId", l_int_NewScoreDetailId)
                l_int_ScoreDetailId(l_int_counter2) = l_int_NewScoreDetailId
            Else
                l_int_ScoreDetailId(l_int_counter2) = l_int_NewScoreDetailId + 1
                l_int_NewScoreDetailId = l_int_ScoreDetailId(l_int_counter2)
            End If
        End If
    Next

    ' release the object variable
    Set oRs = Nothing


'欠席日数がa,bのとき、評定値が入らないため（＝０で登録される）トリガで-1、-2のデータが入ってきたら評定値も同じ値で更新するようにした。
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
            l_str_Sql = "INSERT INTO tbSTEScoreDetail ( iScoreDetailId , iScoreProfileId , iSubjectQuestionId , fDetailScore , dtCreate , dtUpdate ) VALUES("
            l_str_Sql = l_str_Sql & l_int_ScoreDetailId(l_int_Counter) & ","
            l_str_Sql = l_str_Sql & l_int_ScoreProfileId(IIf(g_int_ExamType = 0, l_int_Counter, 0)) & ","
            If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
                l_str_Sql = l_str_Sql & cboInterviewerID.Text & ","
            Else
                l_str_Sql = l_str_Sql & m_intSubQuesProfileId(l_int_Counter) & ","
            End If
            l_str_Sql = l_str_Sql & m_int_Score(l_int_Counter) & ",'"
            l_str_Sql = l_str_Sql & Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
        End If

        Set oRs = g_obj_Conn.Execute(l_str_Sql)

        If Not l_bln_existings(l_int_Counter) Then
            l_int_NewScoreDetailId = l_int_NewScoreDetailId + 1
        End If
    Next

    ' release the object variable
    Set oRs = Nothing

    l_str_Sql = "Update tbSTELocks set iLocks = 0 where vTarget = 'tbSTEScoreProfile' "
    Call g_obj_Conn.Execute(l_str_Sql)

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
    

    ''''VSFlexGrid初期値設置
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

        .Rows = 21
''''    .cols = 3 ''''受験番号を表示しない
        .cols = 2
        .FixedRows = 1


    End With

End Sub

Private Sub f_void_InitializeGrid()

    On Error GoTo ErrorHandler

    Dim oRs                As New ADODB.Recordset
    Dim l_int_Counter      As Long
    Dim l_int_Count        As Long
    Dim l_int_SrNo         As Long
    Dim l_int_ScoreId      As Long
    Dim l_str_Sql          As String
    Dim l_dbl_ChooseiScore As Double

    Dim sSubList           As String
    Dim sSqlSubList        As String
    Dim lLoopCnt           As Long

    Dim sSQL               As String


    Select Case g_int_ExamType    ''''1
    Case 0, 1
        sSQL = "SELECT iSubjectQuestionId, iSubjectProfileId, iQuestionNo, vQuestionName"
        sSQL = sSQL & " FROM tbSTESubjectQuestionProfile as q"

        If g_int_ExamType = 0 Then
            sSQL = sSQL & " WHERE exists ( select 1 from tbSTESubjectProfile as sp where sp.iSubType = 0 and sp.iSubjectProfileID = q.iSubjectProfileId ) "
        Else
            sSQL = sSQL & " Where iSubjectProfileId = " & m_int_SelectedSubject
        End If

        sSQL = sSQL & " Order by iSubjectProfileId , iSubjectQuestionId "


    End Select
    
    oRs.Open sSQL, g_obj_Conn, adOpenStatic, adLockReadOnly
    
    m_int_NoOfQues = oRs.RecordCount

    If m_int_NoOfQues > 10 Then
        m_int_QuestionLimit = 10
    Else
        m_int_QuestionLimit = m_int_NoOfQues
    End If

    If Not oRs.EOF Then
        lblErrorDetails.Caption = ""
        oRs.MoveFirst
        ' store the subject ID for later use
        m_int_SubjId = oRs("iSubjectProfileId")

        With vsfRawScore
            .Rows = 2
            .cols = m_int_QuestionLimit + 5       'get the number of grid cols
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
            .ColWidth(.Col) = 860
            .CellAlignment = flexAlignRightBottom

            .Text = "行番号"      ''''LoadResString(1756) 2021.12.02 update jhi
            
            If .Col < .cols - 1 Then .Col = .Col + 1

            .Text = "受験番号"    ''''LoadResString(1961) 2021.12.02 update jhi
            .CellAlignment = flexAlignRightBottom
''''            .ColWidth(.Col) = 1200 ''''1800
            .ColWidth(.Col) = 1200 ''''1800
           .ColWidth(.Col) = 0 ''''1800

            If .Col < .cols - 1 Then .Col = .Col + 1
            .Text = "受験生氏名"  ''''LoadResString(1805)
            .CellAlignment = flexAlignLeftBottom
            
            .ColWidth(.Col) = 0 'Hide initially
            If .Col < .cols - 1 Then .Col = .Col + 1
            .Text = LoadResString(1962)
            .CellAlignment = flexAlignRightBottom
            .ColWidth(.Col) = 0
            
            If .Col < .cols - 1 Then .Col = .Col + 1
            
            l_int_Counter = 0
            
            For l_int_Count = 0 To m_int_QuestionLimit - 1   'populate header row with names of fields

                If g_int_ExamType = 1 Or g_int_ExamType = 0 Then
                    .Text = Trim(oRs("vQuestionName"))
                ElseIf g_int_ExamType = 2 Then
'                    .Text = Trim(oRs("vInterviewerName"))
                    .Text = Trim(oRs("vQuestionName"))
                End If

'                .CellAlignment = flexAlignRightBottom
                .CellAlignment = flexAlignCenterBottom
                .ColWidth(.Col) = 1500
                m_intSubjectProfileId(l_int_Counter) = Trim(oRs("iSubjectProfileId"))
                m_intSubQuesProfileId(l_int_Counter) = Trim(oRs("iSubjectQuestionId"))
                l_int_Counter = l_int_Counter + 1
                
                ' hide all question columns, except the selected question in the combo box
'                If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
'                    If cboInterviewerID.ListIndex = l_int_Count Then
'                        .ColHidden(.Col) = False
'                    Else
'                        .ColHidden(.Col) = True
'                    End If
'                End If

                If g_int_ExamType = 0 And .Text = gcsHyotei Then
                    '評定値ならばとなりを評定概評にする
                    .cols = .cols + 1
                    .Col = .Col + 1
                    .Text = gcsSeisekiGaihyo
'                    .CellAlignment = flexAlignRightBottom
                    .CellAlignment = flexAlignCenterBottom
                    .ColWidth(.Col) = 1500
                End If

                If .Col < .cols - 1 Then
                    .Col = .Col + 1
                    oRs.MoveNext
                End If

            Next

            ' last
            .Text = "調整点"    ''''LoadResString(1963)
            .ColWidth(.Col) = 0
'            .CellAlignment = flexAlignRightBottom
            .CellAlignment = flexAlignCenterBottom
            .Refresh
        End With
        
       Call f_void_PopulateGrid  'Copied from GetRows
    
    Else
        lblErrorDetails.Caption = "質問はありません。ほかの科目を選択してください" ''''LoadResString(1126)
        ' release the object variables
        oRs.Close
        Set oRs = Nothing
        vsfRawScore.Enabled = False
        Exit Sub
    End If  ' for EOF
    

    ' release the object variables
    oRs.Close
    Set oRs = Nothing

    With vsfRawScore
        .Col = .cols - 1
        .ColWidth(.cols - 1) = IIf(chkChooseiScore.Value = 0, 0, 1800)

        If g_int_ExamType = 1 Or (g_int_ExamType = 2 And m_SecondExam_Type = 1) Then
            .Col = 1
            .ColWidth(.Col) = IIf(chkExamineeName.Value = 0, 0, 1500) ''''2200) 受験番号
        Else
            .Col = 2
            .ColWidth(.Col) = IIf(chkExamineeName.Value = 0, 0, 2200)
        End If

        .Col = 3
        .ColWidth(.Col) = IIf(chkTotalMarks.Value = 0, 0, 1500)

'        lsvCol = -1
        For l_int_Count = 0 To .cols - 1
            .Col = l_int_Count
            If .ColWidth(l_int_Count) > 0 Then
                If .CellBackColor = &HC0C0FF Then
'                   lsvCol = l_int_Count
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
    Dim oRs As New ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_Sql = "SELECT iRoomProfileId, vRoomName FROM tbSTERoomProfile" _
              & " WHERE iMaxCapacity > 0"
    l_str_Sql = l_str_Sql & " AND iInterviewRoomFlag = 0"
    Set oRs = g_obj_Conn.Execute(l_str_Sql)
    If Not oRs.EOF Then
        Do While Not oRs.EOF
            cboRoomName.AddItem oRs("vRoomName")
            cboRoomId.AddItem oRs("iRoomProfileId")
            oRs.MoveNext
        Loop
        cboRoomName.ListIndex = 0
        cboRoomId.ListIndex = 0
    Else
        lblErrorDetails.Caption = LoadResString(2010)
        Unload Me
    End If
    oRs.Close
    Set oRs = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

'*******************************************************************************
'* cboRoomName - 乱数を取得してcboRoomNameに設定する                           *
'*******************************************************************************
Private Sub l_void_AddRoomsRand()

    On Error GoTo ErrorHandler

    Dim adoRs  As New ADODB.Recordset
    Dim sSQL   As String

    
    
    sSQL = "SELECT distinct iRandomNo, iRandomNo FROM tbSTEInterviewRoomProfile" _
              & " WHERE iNendo = ( select top 1 iNendo from tbSTEsystemProfile where iActiveFlag = 1 ) " _
              & " AND iRandomNo is not null "

    Set adoRs = g_obj_Conn.Execute(sSQL)

    If Not adoRs.EOF Then

        Do While Not adoRs.EOF
            cboRoomName.AddItem adoRs("iRandomNo")
            cboRoomId.AddItem adoRs("iRandomNo")    ''''隠しcombo
            adoRs.MoveNext
        Loop

''''2021.12.10 del jhi--> 止めた
        cboRoomName.ListIndex = 0 ''''cboRoomName_Click()が自動動作する
        cboRoomId.ListIndex = 0

    Else
        lblErrorDetails.Caption = "割当可能な会場が見つかりませんでした。先に会場を定義してください。" ''''LoadResString(2010)
''''    Unload Me ''''2021.12.10 add jhi
    End If

    adoRs.Close
    Set adoRs = Nothing

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'*******************************************************************************
'* cboRoomName(乱数)を指定すると cboInterviewer(採点者)データを自動に変化させる*
'* ため この関数を動かせる                                                     *
'*******************************************************************************
Private Sub cboRoomName_Click()

    cboRoomId.ListIndex = cboRoomName.ListIndex

    If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
        f_void_populateInterviewers
    End If

End Sub


Private Sub l_void_PopulateDayCombo()

    Dim sSQL As String
    Dim oRs As ADODB.Recordset
    Dim bThirdDay As Boolean

    bThirdDay = False


''''2022.01.24 del jhi
#If 0 Then

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

#End If



End Sub

Private Sub f_void_populateInterviewers()

    'Populate Interviewers
    Dim l_str_Sql As String
    Dim l_int_Counter As Long
    Dim oRs As New ADODB.Recordset

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

    Set oRs = g_obj_Conn.Execute(l_str_Sql)

    Do While Not oRs.EOF
        l_int_Counter = l_int_Counter + 1
        If l_int_Counter > 10 Then Exit Do
        cboInterviewer.AddItem oRs("VinterviewerName")
        cboInterviewerID.AddItem oRs("iInterviewerProfileID")
        oRs.MoveNext
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

'*******************************************************************************
'* 【配点割合】ボタン 未使用                                                   *
'*******************************************************************************
'add,2007/11/09,S----------
Private Sub cmdHaitenWari_Click()

    dlgHaitenWari.Show 1

End Sub
'add,2007/11/09,E----------

