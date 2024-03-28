VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmRawScore 
   AutoRedraw      =   -1  'True
   ClientHeight    =   9855
   ClientLeft      =   2400
   ClientTop       =   2445
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Palette         =   "frmRawScoreHyotei.frx":0000
   Picture         =   "frmRawScoreHyotei.frx":3AD3
   ScaleHeight     =   9855
   ScaleWidth      =   12765
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdHaitenWari 
      Caption         =   "配　点　割　合"
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
      Left            =   6690
      TabIndex        =   27
      Top             =   8775
      Visible         =   0   'False
      Width           =   3135
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
      TabIndex        =   9
      Top             =   2640
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
      Left            =   2970
      TabIndex        =   11
      Top             =   8760
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
      Left            =   4560
      TabIndex        =   25
      Text            =   "cboInterviewerID"
      Top             =   1620
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
      Left            =   2640
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1560
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
      Height          =   400
      Left            =   10005
      TabIndex        =   5
      Top             =   1035
      Width           =   1500
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
      Height          =   400
      Left            =   6540
      ScrollBars      =   1  '水平
      TabIndex        =   4
      Top             =   1050
      Width           =   1500
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
      Left            =   11745
      TabIndex        =   18
      Text            =   "cboRoomId"
      Top             =   1065
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
      Left            =   10560
      TabIndex        =   8
      Top             =   2235
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
      Top             =   2235
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
      Top             =   2235
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
      Height          =   360
      Left            =   6555
      TabIndex        =   13
      Top             =   1065
      Width           =   1395
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfRawScore 
      Height          =   5100
      Left            =   1275
      TabIndex        =   10
      Top             =   3270
      Width           =   9675
      _cx             =   17066
      _cy             =   8996
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
      FormatString    =   $"frmRawScoreHyotei.frx":75A6
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
      Left            =   10005
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1050
      Width           =   1695
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
      Left            =   6555
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   1050
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
      Height          =   330
      Left            =   8070
      TabIndex        =   26
      Top             =   2715
      Width           =   4290
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
      Left            =   825
      TabIndex        =   24
      Top             =   1620
      Width           =   1695
   End
   Begin VB.Label lblDay 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "面接日"
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
      Left            =   5160
      TabIndex        =   23
      Top             =   1095
      Width           =   1275
   End
   Begin VB.Label lblRoomName 
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
      Height          =   360
      Left            =   8955
      TabIndex        =   22
      Top             =   1125
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
      Left            =   8385
      TabIndex        =   21
      Top             =   1110
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
      Left            =   5040
      TabIndex        =   20
      Top             =   1095
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
      TabIndex        =   19
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
      Height          =   375
      Left            =   4755
      TabIndex        =   17
      Top             =   2190
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
      Height          =   375
      Left            =   8400
      TabIndex        =   16
      Top             =   2175
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
      Height          =   375
      Left            =   810
      TabIndex        =   15
      Top             =   2190
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
      TabIndex        =   12
      Top             =   1110
      Width           =   1275
   End
   Begin VB.Label lblErrorDetails 
      Caption         =   "lblErrorDetails"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1305
      TabIndex        =   14
      Top             =   9300
      Visible         =   0   'False
      Width           =   10515
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

Dim m_obj_Rst                As New ADODB.Recordset
Dim m_str_SQL                As String

Dim m_int_QuestionLimit      As Long

Private Const prvclNoCol     As Long = 0
Private m_SecondExam_Type    As Long '面接か小論文かflag

Private m_bDirty             As Boolean

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Dim adoRs As New ADODB.Recordset ' レコードセット
    Dim sSQL  As String
    Dim icnt  As Long

    m_bDirty = False

''''LoadResStrings Me
    Me.Caption = LoadResString(1951)  '1951:素点を入力してください。
    Call g_void_SetFontProperties(Me) 'set the font properties

    'Grid Setting
    Call f_void_InitGrid



    ' select all subjects that come under the current exam type

    '--------------------------------------------------------------------------
    ' 願書受付フェーズの【評定】処理 <---メニューから削除された 2021.12.22
    '--------------------------------------------------------------------------
    If g_int_ExamType = 0 Then

        '表示に意味は無いので１つにしぼる。
        sSQL = "SELECT TOP 1 iSubjectProfileId,vSubjectName "
        sSQL = sSQL & " FROM tbSTESubjectProfile"
        sSQL = sSQL & " WHERE iExamType = 0"
        sSQL = sSQL & " AND iSubType = 0"
        sSQL = sSQL & " AND iSubType = 0"

    '--------------------------------------------------------------------------
    ' 一次試験フェーズの【素点入力】処理
    '--------------------------------------------------------------------------
    ElseIf g_int_ExamType = 1 Then
        sSQL = "SELECT iSubjectProfileId,vSubjectName "
        sSQL = sSQL & " FROM tbSTESubjectProfile"
        sSQL = sSQL & " WHERE iExamType = 1"

    '---------------------------------------------------------------------------
    '2次試験 処理の【面接】、【小論文】処理
    '---------------------------------------------------------------------------
    ElseIf g_int_ExamType = 2 Then

        ''''科目を取得する --> 小論文
        sSQL = "SELECT iSubjectProfileId,vSubjectName "
        sSQL = sSQL & " FROM tbSTESubjectProfile"

        If m_SecondExam_Type = 0 Then
            sSQL = sSQL & " WHERE iSubType = 3" ''''面接
        Else
            sSQL = sSQL & " WHERE iSubType = 4" ''''小論文
        End If
    End If
    
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

        '科目コントロール部品を見えるように
        lblSubject.Visible = True
        cboSubject.Visible = True

        Do While Not adoRs.EOF
            cboSubject.AddItem adoRs("vSubjectName")                              '小論文
            cboSubject.ItemData(cboSubject.NewIndex) = adoRs("iSubjectProfileId") '30
            adoRs.MoveNext
        Loop

        '-----------------------------------------------------------------------
        '願書受付フェーズ : 評定 ---> menuから削除
        '-----------------------------------------------------------------------
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

        '-----------------------------------------------------------------------
        '1次試験 : 素点入力
        '-----------------------------------------------------------------------
        ElseIf g_int_ExamType = 1 Then

            lblRandomNo.Visible = True '乱数
            txtRandomNo.Visible = True '乱数Text

            cmdUpdate.Left = 4700 ''''【更新】ボタンを表の真ん中に移動 2021.12.22 add jhi

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

            Label4.Visible = False
            chkTotalMarks.Visible = False
            Label3.Visible = False
            chkChooseiScore.Visible = False

            '受験生番号checkbox 2021.12.28 add jhi
            Label2.Caption = "受験番号"
            Label2.Visible = False
            chkExamineeName.Visible = False


        '-----------------------------------------------------------------------
        '2次試験 : 素点入力(面接)
        '-----------------------------------------------------------------------
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

            ' add the subjects to combo box
            Call l_void_AddRooms         'Populate Room Combo
            Call l_void_PopulateDayCombo 'Populate Day combo

            chkTotalMarks.Value = 1          '平均点
            chkChooseiScore.Visible = False  '調整点

        '-----------------------------------------------------------------------
        '2次試験 : 素点入力(小論文)
        '-----------------------------------------------------------------------
        ElseIf g_int_ExamType = 2 And m_SecondExam_Type = 1 Then

            Label2.Caption = "受験番号" ''''chkExamineeName 左Label

            lblRandomNo.Visible = False
            txtRandomNo.Visible = False

'乱数で一意なため、日付は削除
            lblDay.Visible = False
            cboDay.Visible = False
'乱数で一意なため、日付は削除end

            lblJukenNoFrom.Visible = False
            txtJukenNoFrom.Visible = False

            lblJukenNoTo.Visible = False
            txtJukenNoTo.Visible = False

            lblRoomName.Visible = True
            cboRoomName.Visible = True
            lblInterviewers.Visible = True
            cboInterviewer.Visible = True

            '乱数comboboxに値をセットする
            Call l_void_AddRoomsRand  'Populate Room Combo

            '採点者comboboxに値をセットする
''''        Call f_void_populateInterviewers  ''''2021.12.10 del jhi 自動実行されるので

'乱数で一意なため、日付は削除
'           Call l_void_PopulateDayCombo 'Populate Day combo


'           Label2.Visible = False
'           chkExamineeName.Visible = False
'           Label4.Visible = False
'           chkTotalMarks.Visible = False

            Label3.Visible = False

            ''''2021.12.10 add jhi
            chkExamineeName.Value = 1       '受験者氏名 checkboxをon
            chkTotalMarks.Value = 1         '平均点     checkboxをon
            chkChooseiScore.Visible = False '調整点

        End If
    End If

    ' release the object variables
    adoRs.Close
    Set adoRs = Nothing

    cboSubject.ListIndex = 0
    Call f_void_InitGrid

    'cboRoomName.ListIndex = 0
    'cboRoomID.ListIndex = 0



    ' initialize array values to zero
    For icnt = 0 To 9
        m_int_Score(icnt) = 0
    Next

    lblFromToCount.Caption = ""
    'add,2007/11/09,S--------------

    '入試課題要望　面接と小論文の配点割合
    If g_int_ExamType = 2 Then
        Me.cmdHaitenWari.Visible = True
    Else
        Me.cmdHaitenWari.Visible = False
    End If
    'add,2007/11/09,E--------------

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)

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
    MsgBox Err.Description, vbInformation, LoadResString(1729)
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


    '1次試験のみ有効にする
    If g_int_ExamType = 1 Then
        If Trim(txtRandomNo.Text) = "" Then
            MsgBox "乱数番号を入力してレコードを取得してください。"
            Exit Sub
        End If
    End If

    cmdGetRows.Enabled = False
    lblErrorDetails.Caption = ""

'    If m_bDirty Then ''''false
'        If vbCancel = MsgBox("入力データが保存されていません。" & vbCrLf & "保存せずに別データを表示してもよろしいですか？", vbOKCancel) Then
'            Exit Sub
'        End If
'    End If

    m_bDirty = False

    Select Case g_int_ExamType ''''2
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
        m_str_SQL = "SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType=" & g_int_ExamType & " AND vSubjectName='" & cboSubject.Text & "'"
        Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQL)

        ''''2019.05.07 add(comment) jhi 選択されたらid=30(小論文)、id=20(面接Ⅰ)などcomboで選択した科目のIDをセットする
        If Not m_obj_Rst.EOF Then
            m_int_SelectedSubject = m_obj_Rst("iSubjectProfileId") 'コードを変数にセット
        End If

        ' release the object variables
        m_obj_Rst.Close
        Set m_obj_Rst = Nothing

        '***************************************************************************
        '* 2次試験 : 素点入力(面接)                                                *
        '***************************************************************************
        If g_int_ExamType = 2 And m_SecondExam_Type = 0 Then
           'インタビュアーのチェック
            m_str_SQL = "SELECT count(*) as cnt "
            m_str_SQL = m_str_SQL & " FROM tbSTESubjectQuestionProfile a , "
            m_str_SQL = m_str_SQL & "      tbSTEInterviewRoomProfile c , "
            m_str_SQL = m_str_SQL & "      tbSTEInterviewerProfile d "
            m_str_SQL = m_str_SQL & " WHERE a.iSubjectProfileId = " & m_int_SelectedSubject
            m_str_SQL = m_str_SQL & " AND  a.iSubjectProfileId = c.iSubjectProfileId"
            m_str_SQL = m_str_SQL & " AND  c.iRoomProfileId = " & cboRoomId.Text & " "
            m_str_SQL = m_str_SQL & " AND  c.iDayFlag = " & Me.cboDay.ListIndex & " "
            m_str_SQL = m_str_SQL & " AND  d.iInterviewerProfileId = c.iInterviewerProfileId"

            Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQL)

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
 
    '***************************************************************************
    '* 2次試験 : 素点入力(小論文)                                              *
    '***************************************************************************
    If (g_int_ExamType = 2 And m_SecondExam_Type = 1) And cboInterviewer.Text = "" Then

        lblErrorDetails.Caption = "この面接グループに該当する委員がいません。" ''''LoadResString(2484)
        lblErrorDetails.Visible = True

        With vsfRawScore
            .Rows = 2
            .cols = 15
            .Row = 0
            For l_int_Counter = 0 To .cols - 1
                .Col = l_int_Counter
                .Text = ""
            Next
            .Row = 1
            For l_int_Counter = 0 To .cols - 1
                .Col = l_int_Counter
                .Text = ""
            Next
            .Enabled = False
        End With
        cmdUpdate.Enabled = False
        
        ''''2019.05.07 add jhi 選択件数が残ったのでクリア処理を追加した。
        lblFromToCount.Caption = ""
        
        Exit Sub


    '---------------------------------------------------------------------------
    '1次試験、素点入力処理
    '---------------------------------------------------------------------------
    Else ''''2021.12.17(金) ここに来る
        vsfRawScore.Enabled = True
        cmdUpdate.Enabled = True      ''''【更新】    ボタン
        cmdHaitenWari.Enabled = True  ''''【配点割合】ボタン
        
        Call f_void_ClearGrid
        Call f_void_InitializeGrid

    End If


'    With vsfRawScore
'        .Col = .Cols - 1
'        .ColWidth(.Cols - 1) = IIf((chkChooseiScore.Value = 0) Or (Me.cboSubject.Text = gcsKessekiNissuu), 0, 1800)
'        .Col = 2
'        .ColWidth(.Col) = IIf((chkExamineeName.Value = 0), 0, 2200)
'        .Col = 3
'        .ColWidth(.Col) = IIf((chkTotalMarks.Value = 0) Or (Me.cboSubject.Text = gcsKessekiNissuu), 0, 1500)
'    End With


    If vsfRawScore.Rows > 1 Then
'       lblFromToCount.Caption = vsfRawScore.TextMatrix(1, 1) & "～" & vsfRawScore.TextMatrix(vsfRawScore.Rows - 1, 1) & "  " & Trim(str(vsfRawScore.Rows - 1)) & "件"
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
'* 【配点割合】ボタン                                                          *
'*******************************************************************************
'add,2007/11/09,S----------
Private Sub cmdHaitenWari_Click()

    dlgHaitenWari.Show 1

End Sub
'add,2007/11/09,E----------

Private Sub cmdUpdate_Click()

    Dim l_int_Counter As Long
    Dim l_int_LoopCounter As Long
    Dim l_bln_Update As Boolean
    Dim l_str_ErrString As String
    
    Dim SQL As String
    Dim RS As New ADODB.Recordset
    Dim intLockFlag As Integer
    
    On Error GoTo ErrorHandler

    l_bln_Update = True
'このあとのRowcolchangeなどでトータルスコアのカラムのデータが更新されている。
'よってm_bln_Editを変更してはならない

'2017/12/11,S-------
'同時更新　待ち
    intLockFlag = 0
    SQL = "SELECT ISNULL(iLocks,0) iLocks FROM tbSTELocks WITH(NOLOCK) WHERE vTarget = 'tbSTEScoreProfile' "
    Set RS = g_obj_Conn.Execute(SQL)

    If Not RS.EOF Then
        intLockFlag = RS.Fields(0).Value
    End If

    RS.Close
    Set RS = Nothing

    Dim i As Long
    Dim k As Long
    Dim s As Long

    If intLockFlag = 1 Then
        For i = 1 To 10
            For k = 1 To 1000
                s = k
            Next k
        Next i
        cmdUpdate_Click
        
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
        m_bDirty = False
        lblErrorDetails.Caption = LoadResString(2404)
    End If
    End With

Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
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
    
    '**** New code added for day,room and subject (Mahesh)
    
    l_str_SqlDay = "SELECT iNumberOfRoomDay1, iNumberOfRoomDay2, iNumberOfRoomDay3," & _
        " dtSecondExamDay1, dtSecondExamDay2, dtSecondExamDay3," & _
        " iNumberOfExamineeDay1, iNumberOfExamineeDay2, iNumberOfExamineeDay3" & _
        " FROM tbSTESecondExamProfile" & _
        " WHERE iSystemProfileId = (" & _
        " SELECT iSystemProfileId FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag = 1)"

'2021.12.10 add jhi
'SELECT
'    iNumberOfRoomDay1
'   ,iNumberOfRoomDay2
'   ,iNumberOfRoomDay3
'   ,dtSecondExamDay1
'   ,dtSecondExamDay2
'   ,dtSecondExamDay3
'   ,iNumberOfExamineeDay1
'   ,iNumberOfExamineeDay2
'   ,iNumberOfExamineeDay3
'From
'    tbSTESecondExamProfile
'Where
'    iSystemProfileId = ( SELECT iSystemProfileId FROM tbSTESystemProfile WHERE iActiveFlag = 1)


errLine = "2"

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

errLine = "3"

    If Len(Trim(txtRandomNo.Text)) <> 0 Then
        If g_int_ExamType = 1 Then

'入試実施時の不具合No11対応  2004/01/24
            m_str_SQL = "SELECT e.iExamineeProfileId, dbo.usfMakeDispJukenNumber(e.iJukenNumber) as iJukenNumber ,e.vExamineeName " & _
                " FROM tbSTEExamineeProfile e,tbSTERoomProfile r " & _
                " WHERE r.iRoomProfileId = e.iRoomProfileId " & _
                " AND e.iNendo = " & l_int_CurYear & _
                " AND r.iRandom =" & txtRandomNo.Text & _
                " AND e.iExamineeStatus = " & gclExamineeStatus_Default
            m_str_SQL = m_str_SQL & " AND not exists ( select 1 from tbSTEScoreProfile as s , tbSTESubjectProfile as su where s.iExamineeProfileID = e.iExamineeProfileID "
            m_str_SQL = m_str_SQL & " AND su.iSubjectProfileID = s.iSubjectProfileID "
            m_str_SQL = m_str_SQL & " AND su.vSubjectName = '" & cboSubject.Text & "' "
            m_str_SQL = m_str_SQL & " AND s.iAbsentFlag = 1 ) "

        ' comdesign
        '
        '
        '
            Select Case Trim(cboSubject.Text)
            Case "数学"
                m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default
            Case "英語"
                m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=" & m_int_SubjId
            Case "独語"
                m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=" & m_int_SubjId
            Case "仏語"
                m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=" & m_int_SubjId
            Case "物理"
                m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
            Case "化学"
                m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
            Case "生物"
                m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
            End Select

        Else
            m_str_SQL = "SELECT iExamineeProfileId, dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber, vExamineeName from tbSTEExamineeProfile" & _
                " WHERE iexamineeprofileid in(SELECT iExamineeProfileId FROM tbSTEExamineeRoomProfile" & _
                " WHERE iRoomProfileId=(SELECT iRoomProfileId FROM tbSTERoomProfile " & _
                " WHERE iRandom =" & txtRandomNo.Text & "))" & _
                " AND iNendo = " & l_int_CurYear & _
                " AND iExamineeStatus = " & gclExamineeStatus_1stPass & _
                " AND iAbsentFlag = 0"
        End If
    Else
        m_str_SQL = "Select iExamineeProfileId, dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber, vExamineeName" & _
            " from tbSTEExamineeProfile " & _
            " WHERE iNendo=" & l_int_CurYear
        
        If g_int_ExamType = 1 Then
            m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default
                Select Case Trim(cboSubject.Text)
                Case "数学"
                    m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default
                Case "英語"
                    m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=" & m_int_SubjId
                Case "独語"
                    m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=" & m_int_SubjId
                Case "仏語"
                    m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=" & m_int_SubjId
                Case "物理"
                    m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
                Case "化学"
                    m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
                Case "生物"
                    m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & m_int_SubjId & " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
                End Select
            
        ElseIf g_int_ExamType = 2 Then
            'Changes to sql start (Mahesh)
            If m_SecondExam_Type = 0 Then
                m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass & _
                    " AND iExamineeProfileId IN" & _
                    " (SELECT iExamineeProfileId " & _
                    " From tbSteExamineeRoomProfile" & _
                    " WHERE iSubjectProfileid = " & m_int_SubjId & " AND iRoomProfileid = " & cboRoomId.Text & ") AND" & _
                    " dtSecondExamDay = '" & Format(f_dt_SourceDay, "MM/DD/YYYY") & "'"
            Else
                m_str_SQL = m_str_SQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                m_str_SQL = m_str_SQL & " AND iShoronbunRandomNo = " & cboRoomId.Text & " "
                m_str_SQL = m_str_SQL & " AND exists ( SELECT 1 FROM tbSTEInterviewRoomProfile as ir "
                m_str_SQL = m_str_SQL & "       WHERE ir.iRandomNo = " & cboRoomId.Text & " "
                m_str_SQL = m_str_SQL & "       AND ir.iInterviewerProfileId = " & Me.cboInterviewerID.Text & " ) "
            End If
            'Changes end
        End If

'入試実施時の不具合No11対応  2004/01/24
'        m_str_SQl = m_str_SQl & " AND iAbsentFlag = 0"
        
''''2019.05.07 jhi comment文入れた: not exist:親表の中、駆動表にないレコードのみ返す
        m_str_SQL = m_str_SQL & " AND not exists ( select 1 from tbSTEScoreProfile as s , tbSTESubjectProfile as su where s.iExamineeProfileID = tbSTEExamineeProfile.iExamineeProfileID "

''''2019.05.07 jhi for 確認(not除いて existsにしてみた):exist:親表の中、駆動表にあるレコードのみ返す
''''    m_str_SQL = m_str_SQL & " AND     exists ( select 1 from tbSTEScoreProfile as s , tbSTESubjectProfile as su where s.iExamineeProfileID = tbSTEExamineeProfile.iExamineeProfileID "
        
        m_str_SQL = m_str_SQL & " AND su.iSubjectProfileID = s.iSubjectProfileID "
        m_str_SQL = m_str_SQL & " AND su.vSubjectName = '" & cboSubject.Text & "' "
        m_str_SQL = m_str_SQL & " AND s.iAbsentFlag = 1 ) "

        If Trim(txtJukenNoFrom.Text) <> "" And Trim(txtJukenNoTo.Text) <> "" Then
            m_str_SQL = m_str_SQL & " AND iJukenNumber between " & txtJukenNoFrom.Text & " AND " & txtJukenNoTo.Text
        End If
    End If

    m_str_SQL = m_str_SQL & " order by iJukenNumber "

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


    l_obj_Populate.Open m_str_SQL, g_obj_Conn, adOpenStatic, adLockReadOnly

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

'                If g_int_ExamType <> 0 And g_int_ExamType <> 1 Then
'                    l_int_StrLen = Len(l_obj_Populate("iJukenNumber"))
'                    For l_int_Cnt = 0 To l_int_StrLen - 1
'                        .Text = .Text & "*"
'                    Next
'                Else
                    .Text = l_obj_Populate("iJukenNumber")
'                End If

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
                    lblErrorDetails.Visible = False
                Else
                    m_int_QuestionLimit = 10
                    lblErrorDetails.Caption = LoadResString(2500) & m_int_NoOfQues & " " & LoadResString(2501)
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
        lblErrorDetails.Caption = LoadResString(1964)
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
    MsgBox Err.Description, vbInformation, LoadResString(1729)
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

    lblFromToCount.Caption = "" ''''2022.01.20 add jhi recode 取得メッセージclear

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
                            lblErrorDetails.Caption = LoadResString(1956)
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
    MsgBox Err.Description, vbInformation, LoadResString(1729)

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
        ' release the object variable
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
'        If g_int_ExamType = 2 And m_SecondExam_Type = 1 Then
'            l_str_Sql_sub = " AND iSubjectQuestionId=" & cboInterviewerID.Text
'        Else
'            l_str_Sql_sub = " AND iSubjectQuestionId=" & m_intSubQuesProfileId(l_int_Counter2)
'        End If
'
'        Set oRs = g_obj_Conn.Execute(l_str_Sql & l_str_Sql_sub)
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
'                l_str_Sql1 = "SELECT (iScoreDetailId) FROM tbSTEScoreDetail"
'                oRs2.Open l_str_Sql1, g_obj_Conn, adOpenStatic, adLockReadOnly
'                If Not oRs2.EOF Then
'                    oRs2.MoveLast
'                    l_int_ScoreDetailId(l_int_counter2) = oRs2("iScoreDetailId") + 1
'                Else
'                    l_str_Sql1 = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='tbSTEScoreDetail'"
'                    oRs1.Open l_str_Sql1, g_obj_Conn, adOpenStatic, adLockReadOnly
'                    If Not oRs1.EOF Then
'                        l_int_ScoreDetailId(l_int_counter2) = oRs1("iTableCounterIdMapping")
'                    Else
'                        l_int_ScoreDetailId(l_int_counter2) = 1
'                    End If
'                    Set oRs1 = Nothing
'                End If
'                ' release the object variable
'                Set oRs2 = Nothing
'                l_int_NewScoreDetailId = l_int_ScoreDetailId(l_int_counter2) + 1
'                '*******************************************************
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


    Select Case g_int_ExamType    ''''2
    Case 0, 1
        sSQL = "SELECT iSubjectQuestionId, iSubjectProfileId, iQuestionNo, vQuestionName"
        sSQL = sSQL & " FROM tbSTESubjectQuestionProfile as q"
'        sSQL = sSQL & " Where iSubjectProfileId = " & m_int_SelectedSubject

        If g_int_ExamType = 0 Then
            sSQL = sSQL & " WHERE exists ( select 1 from tbSTESubjectProfile as sp where sp.iSubType = 0 and sp.iSubjectProfileID = q.iSubjectProfileId ) "
        Else
            sSQL = sSQL & " Where iSubjectProfileId = " & m_int_SelectedSubject
        End If

        sSQL = sSQL & " Order by iSubjectProfileId , iSubjectQuestionId "

    Case 2
        If m_SecondExam_Type = 0 Then
            sSQL = "SELECT c.iInterviewerProfileId as iSubjectQuestionId, a.iSubjectProfileId,"
            sSQL = sSQL & " d.vInterviewerName as vQuestionName "
            sSQL = sSQL & " FROM tbSTESubjectQuestionProfile a , "
            sSQL = sSQL & "      tbSTEInterviewRoomProfile c , "
            sSQL = sSQL & "      tbSTEInterviewerProfile d "
            sSQL = sSQL & " WHERE a.iSubjectProfileId = " & m_int_SelectedSubject
            sSQL = sSQL & " AND  a.iSubjectProfileId = c.iSubjectProfileId"
            sSQL = sSQL & " AND  c.iRoomProfileId = " & cboRoomId.Text & " "
            sSQL = sSQL & " AND  c.iNendo = " & g_int_CurrentNendo & " "
            sSQL = sSQL & " AND  c.iDayFlag = " & Me.cboDay.ListIndex & " "
            sSQL = sSQL & " AND  d.iInterviewerProfileId = c.iInterviewerProfileId"
            sSQL = sSQL & " ORDER BY c.iInterviewerProfileId"
        Else ''''m_SecondExam_Type=1
            sSQL = "SELECT c.iInterviewerProfileId as iSubjectQuestionId, a.iSubjectProfileId,"
            sSQL = sSQL & " d.vInterviewerName as vQuestionName "
            sSQL = sSQL & " FROM tbSTESubjectQuestionProfile a,"
            sSQL = sSQL & "      tbSTEInterviewRoomProfile c , "
            sSQL = sSQL & "      tbSTEInterviewerProfile d "
            sSQL = sSQL & " WHERE a.iSubjectProfileId = " & m_int_SelectedSubject
            sSQL = sSQL & " AND  a.iSubjectProfileId = c.iSubjectProfileId"
            sSQL = sSQL & " AND  c.iRandomNo = " & cboRoomId.Text & " "
            sSQL = sSQL & " AND  c.iNendo = " & g_int_CurrentNendo & " "
            sSQL = sSQL & " AND  c.iInterviewerProfileId = " & IIf(cboInterviewerID.ListIndex = -1, "-1", cboInterviewerID.Text)
            sSQL = sSQL & " AND  d.iInterviewerProfileId = c.iInterviewerProfileId"
            sSQL = sSQL & " ORDER BY a.iSubjectQuestionId"

'SELECT
'    c.iInterviewerProfileId as iSubjectQuestionId
'   ,a.iSubjectProfileId
'   ,d.vInterviewerName as vQuestionName
'From
'    tbSTESubjectQuestionProfile a
'   ,tbSTEInterviewRoomProfile   c
'   ,tbSTEInterviewerProfile     d
'Where
'        a.iSubjectProfileID = 30
'    AND a.iSubjectProfileId = c.iSubjectProfileId
'    AND  c.iRandomNo = 1
'    AND  c.iNendo = 2020
'    AND  c.iInterviewerProfileId = 125
'    AND  d.iInterviewerProfileId = c.iInterviewerProfileId
'Order By
'   a.iSubjectQuestionId


        End If

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
