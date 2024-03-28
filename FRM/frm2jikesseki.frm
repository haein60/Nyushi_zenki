VERSION 5.00
Begin VB.Form frm2jikesseki 
   Caption         =   "frm2jikesseki : 欠席者入力(2次試験)"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm2jikesseki.frx":0000
   ScaleHeight     =   10065
   ScaleWidth      =   12765
   WindowState     =   2  '最大化
   Begin VB.TextBox txtJuTotal 
      Alignment       =   1  '右揃え
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7470
      Width           =   900
   End
   Begin VB.CommandButton cmdJukenList 
      Caption         =   "2次 受験生リストCSV出力"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      TabIndex        =   24
      Top             =   7860
      Width           =   2800
   End
   Begin VB.CommandButton cmdKessekiList 
      Caption         =   "2次 欠席者リストCSV出力"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7080
      TabIndex        =   23
      Top             =   7860
      Width           =   2800
   End
   Begin VB.TextBox txtSourceName 
      Height          =   345
      Left            =   10440
      TabIndex        =   22
      Text            =   "txtSourceName"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txtDestName 
      Height          =   345
      Left            =   11925
      TabIndex        =   21
      Text            =   "txtDestName"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1410
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
      Left            =   7245
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   18
      Top             =   1080
      Width           =   2355
   End
   Begin VB.ComboBox cboRoomID 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   9660
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   20
      Top             =   1155
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  '右揃え
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11565
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7485
      Width           =   900
   End
   Begin VB.TextBox txtDestJuken 
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
      Height          =   400
      Left            =   8625
      TabIndex        =   2
      Top             =   1650
      Width           =   1400
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
      ItemData        =   "frm2jikesseki.frx":3AD3
      Left            =   1770
      List            =   "frm2jikesseki.frx":3AD5
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1080
      Width           =   2355
   End
   Begin VB.TextBox txtSourceJuken 
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
      Left            =   1770
      TabIndex        =   1
      Top             =   1650
      Width           =   1400
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "2次 欠席者 確定"
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
      Height          =   525
      Left            =   5250
      TabIndex        =   9
      Top             =   8580
      Width           =   2205
   End
   Begin VB.CommandButton cmdDeselectall 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      TabIndex        =   6
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectall 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ListBox lstSelected 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4935
      Left            =   7080
      MultiSelect     =   2  '拡張
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   5370
   End
   Begin VB.ListBox lstExaminees 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4935
      Left            =   240
      MultiSelect     =   2  '拡張
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   5370
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "受験生数"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   240
      TabIndex        =   26
      Top             =   7515
      Width           =   1200
   End
   Begin VB.Label lblRoom 
      Alignment       =   1  '右揃え
      Caption         =   "会場名"
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
      Left            =   5280
      TabIndex        =   19
      Top             =   1140
      Width           =   1755
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  '透明
      Caption         =   "欠席者数"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   7095
      TabIndex        =   16
      Top             =   7515
      Width           =   1200
   End
   Begin VB.Label lblDestJuken 
      Caption         =   "受験番号"
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
      Left            =   7095
      TabIndex        =   15
      Top             =   1725
      Width           =   1365
   End
   Begin VB.Label lblJukenSource 
      BackStyle       =   0  '透明
      Caption         =   "受験番号"
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
      Left            =   270
      TabIndex        =   14
      Top             =   1725
      Width           =   1365
   End
   Begin VB.Label lblErrorDetails 
      BackStyle       =   0  '透明
      Caption         =   "lblErrorDetails"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   420
      TabIndex        =   13
      Top             =   9255
      Width           =   12015
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  '透明
      Caption         =   "科目選択"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1140
      Width           =   1110
   End
   Begin VB.Label lblSelectFrom 
      Alignment       =   2  '中央揃え
      Caption         =   "2次 受験生リスト"
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
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2235
      Width           =   5355
   End
   Begin VB.Label lblSelected 
      Alignment       =   2  '中央揃え
      Caption         =   "2次 欠席者リスト"
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
      Height          =   255
      Left            =   7080
      TabIndex        =   10
      Top             =   2235
      Width           =   5355
   End
End
Attribute VB_Name = "frm2jikesseki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*************************************************************************************************
'Form Name      : frm2jikesseki
'Author         : jhi
'Created On     : 2021.12.29
'Description    : 2次試験 - 欠席者入力
'Reference      :
'***************************************************************************************************

Private f_bln_SelectAll   As Boolean    'Shows the status of the Select All button
Private f_bln_Select      As Boolean    'Shows the status of the Select  button
Private f_bln_DeSelect    As Boolean    'Shows the status of the DeSelectAll button
Private f_bln_DeSelectAll As Boolean    'Shows the status of the DeSelect  button

Dim f_bln_DataChanged     As Boolean    'to enable/disable the save button
Dim f_bln_FormLoaded      As Boolean    'to check whether form is loaded or not

Public m_int_IntRpt       As Long       'form variable variable which indicated whether the form has to be instantiated for the "interview" or "report"
Public m_int_Action       As Long       'determine the action to be performed

''''Private Const prvcSubName_Language As String = "外国語"     ''''2021.12.14 del jhi(外国語はない)
Private Const prvcSubName_Language     As String = "英語"       ''''2021.12.14 add jhi(外国語->英語に変更)
Private Const prvcSubName_Science      As String = "理科"
Private Const prvcSubName_SecondExam   As String = "２次試験"


'*******************************************************************************
'* 2次試験 - 欠席者入力                                                        *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler


    LoadResStrings Me
''''Call g_void_SetFontProperties(Me)     ' set the font properties

    lblErrorDetails.Caption = ""


    f_bln_DataChanged = False
    f_bln_FormLoaded = False
    
    lblRoom.Visible = False
    cboRoom.Visible = False

    m_int_Action = 2 '2021.12.29 強制的 2次試験- 欠席者入力 flagをセット

    '---------------------------------------------------------------------------
    '2次試験:欠席者List
    '---------------------------------------------------------------------------
    Select Case m_int_Action
    Case 2

'        Me.Caption = "frm2jikesseki : 欠席者入力(2次試験)"      ''''LoadResString(1010)
'        lblSelectFrom.Caption = "受験生リスト(2次試験)"         ''''LoadResString(2408)
'        lblSelected.Caption = "欠席者リスト(2次試験)"           ''''LoadResString(2409)
'        lblTotal.Caption = "欠席者数"                           ''''LoadResString(2489)

        lblDestJuken.Visible = False '受験番号lable
        txtDestJuken.Visible = False '受験番号Text

'        cmdOK.Caption = "2次 欠席者 確定"
 
    End Select


    lstExaminees.Font = "ＭＳ ゴシック"
    lstSelected.Font = "ＭＳ ゴシック"

    lstExaminees.FontSize = 10
    lstSelected.FontSize = 10



    '---------------------------------------------------------------------------
    '- 科目を選択comboセット                                                   -
    '---------------------------------------------------------------------------
    '2次試験「欠席者List」入力では科目を表示する
    If m_int_Action = 2 Then                        '2: 2次試験　欠席者List syori

        'input absentee record
        Call f_void_cboSubject_Get
        cboSubject.ListIndex = 0

    End If



    '---------------------------------------------------------------------------
    ' 1次試験 欠席者入力, 2次試験 欠席者入力 データ設定
    '---------------------------------------------------------------------------
    If m_int_Action = 0 Or m_int_Action = 2 Then    '2: 2次試験　欠席者List syori
        Call f_void_SelectAbsentee
    Else
        Call f_void_Select
    End If

    cmdDeselect.Enabled = False
    cmdSelect.Enabled = False

    Call f_void_CheckButtonStatus

    txtJuTotal.Text = lstExaminees.ListCount '受験生count
    txtTotal.Text = lstSelected.ListCount    '欠席者count

    f_bln_FormLoaded = True

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim i As Long

    
    fMainForm.mnuTools.Enabled = False

    For i = 1 To fMainForm.Toolbar1.Buttons.Count
        ' disable the toolbar buttons
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next i

'    If m_int_Action = 0 Or m_int_Action = 2 Then
'        Call f_void_SelectAbsentee
'    Else
'        Call f_void_Select
'    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub cboRoom_Click()

    On Error GoTo ErrorHandler
    Dim L_str_temp As String

    cboRoomID.ListIndex = cboRoom.ListIndex
'    If f_bln_FormLoaded Then Call f_void_SelectAbsentee
    
    L_str_temp = UCase(LoadResString(2474)) & "*"
    lblErrorDetails.Caption = ""

    If m_int_Action = 2 Then
        If UCase(cboSubject) Like L_str_temp Then
            g_int_ExamType = 2
        Else
            g_int_ExamType = 3
        End If
    End If

    If f_bln_FormLoaded Then
        Call f_void_SelectAbsentee
    End If

    Exit Sub

ErrorHandler:
     MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_cboSubject_Get()

    On Error GoTo ErrorHandler
    Dim sSQL          As String                 ' SQL string
    Dim adoRs         As New ADODB.Recordset    ' recordset object
    Dim l_int_Counter As Long
 
   
    ' select all subjects that come under the selected exam type
    sSQL = "SELECT iSubjectprofileID,vSubjectName FROM tbSTESubjectProfile "

    If m_int_Action = 0 Then
        ''''sSQL = sSQL & "WHERE iExamType =" & g_int_ExamType
    ElseIf m_int_Action = 2 Or m_int_Action = 3 Then
        sSQL = sSQL & "WHERE iExamType IN(2,3,4,5)"
    End If

    sSQL = sSQL & " ORDER BY iDispOrder"

'-------------------------------------------------------------------------------
'2021.12.14 add jhi
'SELECT
'    --iSubjectprofileID
'   --,vSubjectName
'   *
'From
'    tbSTESubjectProfile
'Where
'    iExamType = 1
'ORDER BY iDispOrder
'-------------------------------------------------------------------------------
    
    Set adoRs = g_obj_Conn.Execute(sSQL)
    
    ' add the subjects to combo box
    Do While Not adoRs.EOF
        l_int_Counter = l_int_Counter + 1
        cboSubject.AddItem adoRs("vSubjectName")
        adoRs.MoveNext
    Loop
    
    ' release the object variables
    adoRs.Close
    Set adoRs = Nothing

    '1次試験の欠席者List
''''2021.12.28 del jhi なぜこれを追加するのか？
'    If m_int_Action = 0 Then
'        cboSubject.AddItem prvcSubName_Science, 0  '理科
'        cboSubject.AddItem prvcSubName_Language, 0 '英語
'    End If

    '2次試験の欠席者List
    If m_int_Action = 2 Then
        cboSubject.AddItem prvcSubName_SecondExam, 0 '2次試験
    End If

'    If l_int_Counter > 0 Then
'        cboSubject.ListIndex = 0
'    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)
End Sub


' The different values of m_int_action and what they stand for
'   0   -   Input Absentee Record for 1st exam
'   1   -   Input Passed Person data for 1st exam
'   2   -   Input absentee record for 2nd exam
'   3   -   Input Passed Person data for 2nd exam
'   4   -   Input waiting list for 2nd exam
'   5   -   upliftment from waiting list for Enter/Refuse phase
'   6   -   Input Refuse offer for Enter/Refuse phase

Private Sub cboSubject_Click()

    On Error GoTo ErrorHandler
    Dim L_str_temp As String
    
    L_str_temp = UCase(LoadResString(2474)) & "*"
    lblErrorDetails.Caption = ""

    If m_int_Action = 2 Then
        If UCase(cboSubject) Like L_str_temp Then
            g_int_ExamType = 2
        Else
            g_int_ExamType = 3
        End If
    End If

    If f_bln_FormLoaded Then
        Call f_void_SelectAbsentee
    End If

    Exit Sub

ErrorHandler:
     MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorHandler

    Dim l_int_Count              As Long                   ' counter
    Dim l_int_TempJuken          As Long                   ' to store the juken number
    Dim l_str_JukenNo            As String                 ' to store all the selected juken numbers as a string
    Dim l_str_NonSelected        As String                 ' to store all the non-selected juken numbers as a string
    Dim l_str_ExamineeID         As String                 ' string of examinee id's
    Dim l_obj_Rec                As ADODB.Recordset        ' recordset variable
    Dim l_str_Sql                As String                 ' to store the SQl string
    Dim l_str_MySql              As String
    Dim l_obj_Rst                As New ADODB.Recordset    ' recordset variable
    Dim l_obj_rst1               As New ADODB.Recordset
    Dim l_obj_rst2               As New ADODB.Recordset
    Dim l_obj_rst3               As New ADODB.Recordset
    Dim l_obj_rst4               As New ADODB.Recordset
    Dim l_str_ExamineeIDSql      As String                 ' to store the SQL string
    Dim l_int_subjectProfileId   As Long                   ' to store the subject profile Id
    Dim l_int_NewScoreProfileId  As Long                   ' to store the score profile Id
    Dim l_str_Sql1               As String                 ' to store the SQL string
    Dim l_str_sql2               As String

    Dim bRtn As Boolean


    ' マウスポインタを砂時計にします。
    Screen.MousePointer = vbHourglass
    lblErrorDetails.Caption = "欠席者の更新処理を行っています。しばらくお持ちください。"
    cmdOK.Enabled = False

    DoEvents


    
    ' get all the examinees in selected list box into a single string
    For l_int_Count = 0 To lstSelected.ListCount - 1
        l_int_TempJuken = Left(lstSelected.List(l_int_Count), 4)
        l_str_JukenNo = l_str_JukenNo & "," & l_int_TempJuken
    Next

    If Len(Trim(l_str_JukenNo)) > 0 Then
        l_str_JukenNo = Right(Trim(l_str_JukenNo), Len(Trim(l_str_JukenNo)) - 1)
    End If
    
    ' get all the examinees in non-selected examinees(left) list box into a single string
    For l_int_Count = 0 To lstExaminees.ListCount - 1
        l_int_TempJuken = Left(lstExaminees.List(l_int_Count), 4)
        l_str_NonSelected = l_str_NonSelected & "," & l_int_TempJuken
    Next

    If Len(Trim(l_str_NonSelected)) > 0 Then
        l_str_NonSelected = Right(Trim(l_str_NonSelected), Len(Trim(l_str_NonSelected)) - 1)
    End If
    
    If lstSelected.ListCount > 0 Or lstExaminees.ListCount > 0 Then
        
        g_obj_Conn.BeginTrans   ' start a transaction as there are multiple database table inserts/updates
        
        Select Case m_int_Action
        Case 2

            ' get the selected subject
            If Trim(cboSubject.Text) = prvcSubName_SecondExam Then
                l_obj_rst3.Open "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
                    " WHERE iExamType = 2 ", g_obj_Conn
            Else
                l_obj_rst3.Open "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
                    " WHERE vSubjectName='" & Trim(cboSubject.Text) & "'", g_obj_Conn
            End If

            Do Until l_obj_rst3.EOF

                l_int_subjectProfileId = l_obj_rst3("isubjectprofileid")

                ' insert/update details of selected examinees
                For l_int_Count = 0 To lstSelected.ListCount - 1

                    bRtn = getNewId("tbSTEScoreProfile", "iScoreProfileId", l_int_NewScoreProfileId)

                    l_int_TempJuken = Left(lstSelected.List(l_int_Count), 4)

                    l_str_ExamineeIDSql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                        " WHERE iNendo = " & g_int_CurrentNendo & _
                        " AND iJukenNumber = " & l_int_TempJuken
                    If m_int_Action = 0 Then
                        ' input absentee record for 1st exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                    Else
                        ' input absentee record for 2nd exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                    End If

                    l_obj_rst4.Open l_str_ExamineeIDSql, g_obj_Conn, adOpenStatic, adLockReadOnly

                    l_str_Sql = "SELECT iScoreProfileId FROM tbSTEScoreProfile" & _
                        " WHERE iExamineeProfileId=" & l_obj_rst4("iExamineeProfileId") & _
                        " AND iSubjectProfileId=" & l_int_subjectProfileId & _
                        " AND iAbsentFlag = 1"
                    l_obj_rst2.Open l_str_Sql, g_obj_Conn
                    If l_obj_rst2.EOF Then
                        l_str_Sql = "INSERT INTO tbSTEScoreProfile (iScoreProfileId,iSubjectProfileId,iExamineeProfileId,iAbsentFlag,dtCreate,dtUpdate) VALUES(" & _
                            l_int_NewScoreProfileId & "," & _
                            l_int_subjectProfileId & "," & _
                            l_obj_rst4("iExamineeProfileId") & ", 1,'" & _
                            Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
                    End If
                    l_obj_rst2.Close
                    Set l_obj_rst2 = Nothing
                    
                    g_obj_Conn.Execute l_str_Sql
                    
                    l_str_Sql = "UPDATE tbSTEExamineeProfile SET iAbsentFlag = 1, dtUpdate='" & Format(Date, "MM/DD/YYYY") & "' WHERE" & _
                        " iNendo = " & g_int_CurrentNendo & _
                        " AND iExamineeProfileId = " & l_obj_rst4("iExamineeProfileId")
                    If m_int_Action = 0 Then
                        ' input absentee record for 1st exam
                        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                    Else
                        ' input absentee record for 2nd exam
                        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                    End If
                    
                    g_obj_Conn.Execute l_str_Sql
                    
                    Set l_obj_rst4 = Nothing
                Next

                ' insert/update details of non-selected examinees
                For l_int_Count = 0 To lstExaminees.ListCount - 1
                    l_int_TempJuken = Left(lstExaminees.List(l_int_Count), 4)
                    
                    l_str_ExamineeIDSql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                        " WHERE iNendo = " & g_int_CurrentNendo & _
                        " AND iJukenNumber = " & l_int_TempJuken
                    If m_int_Action = 0 Then
                        ' input absentee record for 1st exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                    Else
                        ' input absentee record for 2nd exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                    End If
                    
                    l_obj_rst4.Open l_str_ExamineeIDSql, g_obj_Conn, adOpenStatic, adLockReadOnly
                                            
                    l_str_Sql = "DELETE FROM tbSTEScoreProfile WHERE iAbsentFlag = 1" & _
                        " AND iExamineeProfileId=" & l_obj_rst4("iExamineeProfileId") & _
                        " AND iSubjectProfileId=" & l_int_subjectProfileId
                      
                    g_obj_Conn.Execute l_str_Sql
                    
                    ' check whether the examinee is present for all other subjects
                    l_str_Sql1 = "SELECT iSubjectProfileId FROM tbSTEScoreProfile" & _
                        " WHERE iSubjectProfileId <>" & l_int_subjectProfileId & _
                        " AND iExamineeProfileId=" & l_obj_rst4("iExamineeProfileId") & _
                        " and iAbsentFlag=1"
                    l_obj_rst1.Open l_str_Sql1, g_obj_Conn
                    If l_obj_rst1.EOF Then
                                        
                        l_str_Sql = "UPDATE tbSTEExamineeProfile SET iAbsentFlag = 0," & _
                        " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                        " WHERE iNendo = " & g_int_CurrentNendo & _
                        " AND iExamineeProfileId = " & l_obj_rst4("iExamineeProfileId")
                        
                        If m_int_Action = 0 Then
                            ' input absentee record for 1st exam
                            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                        Else
                            ' input absentee record for 2nd exam
                            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                        End If
                        
                        g_obj_Conn.Execute l_str_Sql
                    End If
                    
                    l_obj_rst1.Close
                    Set l_obj_rst1 = Nothing
                    Set l_obj_rst4 = Nothing
                Next

                l_obj_rst3.MoveNext

            Loop

            l_obj_rst3.Close
            Set l_obj_rst3 = Nothing
            
        End Select
        
        g_obj_Conn.CommitTrans
        
        If f_bln_DataChanged Then
            f_bln_DataChanged = False
            cmdOK.Enabled = False
        End If
        lblErrorDetails.Caption = "欠席者更新処理が正常に終了しました。" ''''LoadResString(2404)
    End If

    ' マウスポインタを規定値に戻します。
    Screen.MousePointer = vbDefault

    Exit Sub

ErrorHandler:
    g_obj_Conn.RollbackTrans
    lblErrorDetails.Caption = "欠席者更新処理でエラーが発生しました。" ''''LoadResString(2405)
    MsgBox Err.Description, vbInformation, "エラー"

    ' マウスポインタを規定値に戻します。
    Screen.MousePointer = vbDefault


End Sub

Private Sub cmdSelectAll_Click()
    'On the click of this button all the Examinees from the lstExaminees will be transfered to lstSelectedInterviewers
    Dim l_int_Examinees As Long
    On Error GoTo ErrorHandler
    
    f_bln_SelectAll = True
    
    lblErrorDetails.Caption = ""
    If lstExaminees.ListCount >= 1 Then
        For l_int_Examinees = lstExaminees.ListCount - 1 To 0 Step -1
            lstSelected.AddItem lstExaminees.List(l_int_Examinees)
            lstExaminees.ListIndex = l_int_Examinees
            lstExaminees.RemoveItem l_int_Examinees
        Next
    End If

    f_void_CheckButtonStatus
    f_bln_SelectAll = False
    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtJuTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount

    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdSelect_Click()
    'on the click of this button only the Examinee selected from the lstExaminees will be transfered to
    'lstSelected
    Dim l_int_Count As Long
    On Error GoTo ErrorHandler
    
    f_bln_Select = True
    lblErrorDetails.Caption = ""
    If lstExaminees.SelCount > 0 Then
        For l_int_Count = lstExaminees.ListCount - 1 To 0 Step -1
            If lstExaminees.Selected(l_int_Count) Then
                lstSelected.AddItem lstExaminees.List(l_int_Count)
                lstExaminees.RemoveItem l_int_Count
            End If
        Next
    End If
    f_void_CheckButtonStatus
    f_bln_Select = False
    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtJuTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

'*******************************************************************************
'on the click of this button only the interviewer selected from the lstSelected will be
'transfered to lstExaminees
'*******************************************************************************
Private Sub cmdDeselect_Click()

    On Error GoTo ErrorHandler

    Dim l_int_Count As Long
    
    lblErrorDetails.Caption = ""
    f_bln_DeSelect = True
        If lstSelected.SelCount > 0 Then
            For l_int_Count = lstSelected.ListCount - 1 To 0 Step -1
                If lstSelected.Selected(l_int_Count) Then
                    lstExaminees.AddItem lstSelected.List(l_int_Count)
                    lstSelected.RemoveItem l_int_Count
                End If
            Next
        End If
    f_void_CheckButtonStatus
    f_bln_DeSelect = True
    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtJuTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'*******************************************************************************
'on the click of this button all the Examinees from the lstSelectedInterviewers will be moved to
'LstAllinterviewers
'*******************************************************************************
Private Sub cmdDeselectAll_Click()

    On Error GoTo ErrorHandler
    Dim l_int_InterviewerCount As Long
    
    lblErrorDetails.Caption = ""
    f_bln_DeSelectAll = True
    If lstSelected.ListCount >= 1 Then
       For l_int_InterviewerCount = lstSelected.ListCount - 1 To 0 Step -1
            lstExaminees.AddItem lstSelected.List(l_int_InterviewerCount)
            lstSelected.RemoveItem l_int_InterviewerCount
        Next
    End If
    f_void_CheckButtonStatus
    f_bln_DeSelectAll = False
    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtJuTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'*******************************************************************************
'Private Sub l_SetUpdateButtonEnabled()
'
'    If dd Then
'    End If
'
'End Sub

'*******************************************************************************
Public Sub f_void_CheckButtonStatus()

    'Procedure to check the status of the buttons
    'i.e enabling and disabling the buttons based on the presense
    'and selection of data in the list boxes

    If lstExaminees.ListCount = 0 Then
        cmdSelectall.Enabled = False
        cmdSelect.Enabled = False
    Else
        cmdSelectall.Enabled = True
        If lstExaminees.SelCount > 0 Then
            cmdSelect.Enabled = True
        Else
            cmdSelect.Enabled = False
        End If
    End If
    
    If lstSelected.ListCount = 0 Then
        cmdDeselectall.Enabled = False
        cmdDeselect.Enabled = False
    Else
        cmdDeselectall.Enabled = True
        If lstSelected.SelCount > 0 Then
            cmdDeselect.Enabled = True
        Else
            cmdDeselect.Enabled = False
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    f_bln_DataChanged = False
    Call g_void_CloseChildForm
    Unload Me
End Sub

Private Sub lstExaminees_Click()
    'Enables the cmdselect button when any element in the list box is selected else
    'button remains disabled
    f_void_CheckButtonStatus
End Sub

Private Sub lstExaminees_DblClick()
    cmdSelect_Click
    f_void_CheckButtonStatus
End Sub

Private Sub lstSelected_Click()
    'Enables the cmddeselect button when any element in the list box is selected else
    'button remains disabled
    f_void_CheckButtonStatus
End Sub

Private Sub lstSelected_DblClick()
    cmdDeselect_Click
    f_void_CheckButtonStatus
End Sub

Private Sub f_void_Select()

    Dim l_obj_Rst As ADODB.Recordset    ' recordset object
    Dim l_str_Sql As String             ' The SQL string
    Dim l_str_DisplayString As String   ' to form the display string in the list box
        
    lstExaminees.Clear
    lstSelected.Clear
        
    l_str_Sql = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber,substring(vExamineeName + '　　　　　　　　',1,10) as vExamineeName,iSex FROM tbSTEExamineeProfile WHERE" & _
        " iNendo = " & g_int_CurrentNendo & _
        " AND iAbsentFlag = 0"
    
    Select Case m_int_Action
    Case 1   ' 1st exam
        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default
    Case 3, 4    ' 2nd exam
        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    Case 5  ' enter/refuse phase
        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_2ndWait
    Case 6  ' enter/refuse phase
        l_str_Sql = l_str_Sql & " AND (iExamineeStatus = " & gclExamineeStatus_2ndPass & " or iExamineeStatus = " & gclExamineeStatus_2ndWaitPass & ") and iRejectFlag = 0"
    End Select
        
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
'    If l_obj_Rst.EOF Then
'        Set l_obj_Rst = Nothing
'        Exit Sub
'    End If
    Do While Not l_obj_Rst.EOF
        l_str_DisplayString = l_obj_Rst.Fields("iJukenNumber").Value & _
            " - " & l_obj_Rst.Fields("vExamineeName").Value
        If l_obj_Rst.Fields("iSex").Value = 0 Then
            l_str_DisplayString = l_str_DisplayString & " - (*)"
        Else
            l_str_DisplayString = l_str_DisplayString & "      "
        End If

        lstExaminees.AddItem l_str_DisplayString
        l_obj_Rst.MoveNext
    Loop
    Set l_obj_Rst = Nothing
    
    l_str_Sql = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber,substring(vExamineeName + '　　　　　　　　',1,10) as vExamineeName,iSex FROM tbSTEExamineeProfile  WHERE"
    l_str_Sql = l_str_Sql & " iNendo = " & g_int_CurrentNendo
    
    Select Case m_int_Action
    Case 1  ' input passed person data
        l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    Case 3  ' passed person data for 2nd phase
        l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_2ndPass
    Case 4  ' waiting list
        l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_2ndWait
    Case 5  ' upliftment from waiting list
        l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_2ndWaitPass
    Case 6  ' enter/refuse offer
        l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0" & _
            " AND iRejectFlag = 1" & _
            " AND iExamineeStatus IN (" & gclExamineeStatus_2ndPass & "," & gclExamineeStatus_2ndWaitPass & ")"
    End Select
        
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If l_obj_Rst.EOF Then
        Set l_obj_Rst = Nothing
        Exit Sub
    End If
    Do While Not l_obj_Rst.EOF
        l_str_DisplayString = l_obj_Rst.Fields("iJukenNumber").Value & " - " & l_obj_Rst.Fields("vExamineeName").Value
        If l_obj_Rst.Fields("iSex").Value = 0 Then
            l_str_DisplayString = l_str_DisplayString & " - (*)"
        Else
            l_str_DisplayString = l_str_DisplayString & "      "
        End If
        lstSelected.AddItem l_str_DisplayString
        l_obj_Rst.MoveNext
    Loop
    
    Set l_obj_Rst = Nothing
End Sub


Private Sub txtDestJuken_KeyPress(KeyAscii As Integer)
    ' move the input juken number from the non-selected listbox to the selected listbox
    Dim l_str_sqlExaminee As String             ' to form the examinee details query string
    Dim l_obj_rsExaminee As New ADODB.Recordset ' to hold the examinee details records
    Dim l_str_JukenNo As String                 ' to sotre the input juken number
    Dim l_int_counter1 As Long               ' to loop through the list box
    Dim l_int_counter2 As Long               ' to loop through the list box
    
    On Error GoTo ErrorHandler
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
        
    If KeyAscii = 13 Then
        
        If Trim(txtDestJuken.Text) = "" Then Exit Sub     ' exit if the textbox is empty
        
        ' enable the functionality only for the "Enter/Return key"
        l_str_sqlExaminee = "SELECT iJukenNumber, substring(vExamineeName + '　　　　　　　　',1,10) as vExamineeName FROM tbSTEExamineeProfile" & _
            " WHERE iJukenNumber=" & Trim(txtDestJuken.Text) & " AND iNendo=" & g_int_CurrentNendo
        l_obj_rsExaminee.Open l_str_sqlExaminee, g_obj_Conn
        
            
        If l_obj_rsExaminee.EOF Then
            ' the input juken number does not exist - display an error message
            lblErrorDetails.Caption = LoadResString(2473)
        Else
            lblErrorDetails.Caption = ""
            ' pad the input juken number with leading "0"
            l_str_JukenNo = g_str_LPad(Trim(txtDestJuken.Text), Len(Trim(txtDestJuken.Text)))
            
            For l_int_counter1 = 0 To lstSelected.ListCount - 1
                ' loop through the list box to check whether the juken number is present or not
                If Left(lstSelected.List(l_int_counter1), 4) = l_str_JukenNo Then
                    ' input juken is presnet
                    
                    ' display examinee name in the neme text box
                    txtDestName.Text = l_obj_rsExaminee.Fields("vExamineeName").Value
                    
                    ' make it the selected one
                    lstSelected.Selected(l_int_counter1) = True
                    
                    ' move it to the non-selected listbox
                    lblErrorDetails.Caption = ""
                    f_bln_DeSelect = True
                        
                    lstExaminees.AddItem lstSelected.List(l_int_counter1)
                    lstSelected.RemoveItem l_int_counter1
                                
                    f_void_CheckButtonStatus
                    f_bln_DeSelect = True
                    If Not f_bln_DataChanged Then
                        f_bln_DataChanged = True
                        cmdOK.Enabled = True
                    End If
                    txtTotal.Text = lstSelected.ListCount
                    
                    ' loop thourh the nonselected listbox, and highlight the input juken number
                    For l_int_counter2 = 0 To lstExaminees.ListCount - 1
                        If Left(lstExaminees.List(l_int_counter2), 4) = l_str_JukenNo Then
                            lstExaminees.Selected(l_int_counter2) = True
                        Else
                            lstExaminees.Selected(l_int_counter2) = False
                        End If
                    Next
                    txtDestJuken.Text = ""
                    txtDestName.Text = ""
                    Exit Sub
                End If
            Next
        End If
        l_obj_rsExaminee.Close
        Set l_obj_rsExaminee = Nothing
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub txtSourceJuken_KeyPress(KeyAscii As Integer)
    ' move the input juken number from the selected listbox to the non-selected listbox
    Dim l_str_sqlExaminee As String             ' to form the examinee details query string
    Dim l_obj_rsExaminee As New ADODB.Recordset ' to hold the examinee details records
    Dim l_str_JukenNo As String                 ' to sotre the input juken number
    Dim l_int_counter1 As Long               ' to loop through the list box
    Dim l_int_counter2 As Long               ' to loop through the list box
    
    On Error GoTo ErrorHandler
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
        
    If KeyAscii = 13 Then
        
        If Trim(txtSourceJuken.Text) = "" Then Exit Sub     ' exit if the textbox is empty
        
        ' enable the functionality only for the "Enter/Return key"
        l_str_sqlExaminee = "SELECT iJukenNumber, substring(vExamineeName + '　　　　　　　　',1,10) as vExamineeName FROM tbSTEExamineeProfile" & _
            " WHERE iJukenNumber=" & Trim(txtSourceJuken.Text) & " AND iNendo=" & g_int_CurrentNendo
        l_obj_rsExaminee.Open l_str_sqlExaminee, g_obj_Conn
            
        If l_obj_rsExaminee.EOF Then
            ' the input juken number does not exist - display an error message
            lblErrorDetails.Caption = LoadResString(2473)
        Else
            lblErrorDetails.Caption = ""
            ' pad the input juken number with leading "0"
            l_str_JukenNo = g_str_LPad(Trim(txtSourceJuken.Text), Len(Trim(txtSourceJuken.Text)))
            
            ' loop through the list box to check whether the juken number is present or not
            For l_int_counter1 = 0 To lstExaminees.ListCount - 1
                If Left(lstExaminees.List(l_int_counter1), 4) = l_str_JukenNo Then
                    ' input juken is presnet
                    
                    ' display examinee name in the name text box
                    txtSourceName.Text = l_obj_rsExaminee.Fields("vExamineeName").Value
                    
                    ' make it the selected one
                    lstExaminees.Selected(l_int_counter1) = True
                    
                    ' move it to the selected listbox
                    f_bln_Select = True
                    lblErrorDetails.Caption = ""
                    
                    lstSelected.AddItem lstExaminees.List(l_int_counter1)
                    lstExaminees.RemoveItem l_int_counter1
                           
                    f_void_CheckButtonStatus
                    f_bln_Select = False
                    If Not f_bln_DataChanged Then
                        f_bln_DataChanged = True
                        cmdOK.Enabled = True
                    End If
                    txtTotal.Text = lstSelected.ListCount
                    
                    ' loop thourh the selected listbox, and highlight the input juken number
                    For l_int_counter2 = 0 To lstSelected.ListCount - 1
                        If Left(lstSelected.List(l_int_counter2), 4) = l_str_JukenNo Then
                            lstSelected.Selected(l_int_counter2) = True
                        Else
                            lstSelected.Selected(l_int_counter2) = False
                        End If
                    Next
                    txtSourceJuken.Text = ""
                    txtSourceName.Text = ""
                End If
            Next
            
        End If
        l_obj_rsExaminee.Close
        Set l_obj_rsExaminee = Nothing
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

'*******************************************************************************
'* ListBoxにデータを表示する                                                   *
'* 2022.01.04 update                                                           *
'*******************************************************************************
Private Sub f_void_SelectAbsentee()

    Dim oRs              As ADODB.Recordset         ' recordset object
    Dim oRs_RoomName     As New ADODB.Recordset

    Dim sSQL             As String                  ' The SQL string
    Dim sSQLRoomName     As String

    Dim strDisp          As String                  ' to form the display string in the list box
    
    lstExaminees.Clear
    lstSelected.Clear

    sSQL = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber,substring(vExamineeName + '　　　　　　　　',1,10) as vExamineeName,iSex,iRoomProfileId"
    sSQL = sSQL & " FROM tbSTEExamineeProfile WHERE iNendo = " & g_int_CurrentNendo
    sSQL = sSQL & " AND iExamineeProfileId NOT IN("
    sSQL = sSQL & " SELECT iExamineeProfileId FROM tbSTEScoreProfile"
    sSQL = sSQL & " WHERE iSubjectProfileId in (SELECT iSubjectProfileId FROM tbSTESubjectProfile"

    Select Case Trim(cboSubject.Text)
    Case prvcSubName_Science
        sSQL = sSQL & " WHERE vSubjectName in ('物理' , '化学' , '生物' ) ) "
    Case prvcSubName_Language
        sSQL = sSQL & " WHERE vSubjectName in ('仏語' , '独語' , '英語' ) ) "
    Case prvcSubName_SecondExam
        sSQL = sSQL & " WHERE vSubjectName in ('面接Ⅰ' , '面接Ⅱ' , '小論文' ) ) "
    Case Else
        sSQL = sSQL & " WHERE vSubjectName='" & Trim(cboSubject.Text) & "' ) "
    End Select

    sSQL = sSQL & " AND tbSTEScoreProfile.iAbsentFlag=1) "

    If m_int_Action = 0 Then
        sSQL = sSQL & " AND iRoomProfileId=" & cboRoomID.Text & " "
    End If

    Select Case m_int_Action
    Case 0   ' 1st exam

        ' sSQL = sSQL & " AND iExamineeStatus = 0"
        ' modify form codesign 16/08/02
        '
        Select Case Trim(cboSubject.Text)
        Case "数学"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default
        Case "英語"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "')"
        Case "独語"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "')"
        Case "仏語"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "')"
        Case "物理"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "') in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
        Case "化学"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "') in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
        Case "生物"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "') in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
        Case prvcSubName_Science
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & _
            " ( iScienceSubjProfileId1 in (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName in ('物理' , '化学' , '生物' ) ) " & _
            " OR iScienceSubjProfileId2 in (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName in ('物理' , '化学' , '生物' ) ) ) "
        Case prvcSubName_Language
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & _
            " iLanguageSubjProfileId in (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName in ('仏語' , '独語' , '英語' ) ) "
        End Select

    Case 2    ' 2nd exam
        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
''''    sSQL = sSQL & " order by iJukenNumber" ''''2022.01.12 del
        sSQL = sSQL & " order by iRoomProfileId,iJukenNumber"
    End Select

    Set oRs = g_obj_Conn.Execute(sSQL)

    If oRs.EOF Then
        txtTotal.Text = lstSelected.ListCount

'        Set oRs = Nothing
'        Exit Sub
    End If

    Do While Not oRs.EOF
        ' form the string to be displayed in the listbox
        strDisp = oRs.Fields("iJukenNumber").Value & _
            " - " & oRs.Fields("vExamineeName").Value

        If oRs.Fields("iSex").Value = 0 Then
            strDisp = strDisp & " (*)"
        Else
            strDisp = strDisp & "    "
        End If
            
        ' check whether the examinee is allocated to any room or not
        If Trim(oRs.Fields("iRoomProfileId").Value) <> "" Then
            
            sSQLRoomName = "SELECT vRoomName FROM tbSTERoomProfile" & _
                " WHERE iRoomProfileId=" & oRs.Fields("iRoomProfileId").Value
            oRs_RoomName.Open sSQLRoomName, g_obj_Conn
            
            If Not oRs_RoomName.EOF Then
                strDisp = strDisp & " - " & oRs_RoomName.Fields("vRoomName").Value
            End If
            
            oRs_RoomName.Close
            Set oRs_RoomName = Nothing
        End If

        lstExaminees.AddItem strDisp
        oRs.MoveNext
    Loop
 
    oRs.Close
    Set oRs = Nothing

    '---------------------------------------------------------------------------
    ' 欠席者List SQL
    '---------------------------------------------------------------------------
    sSQL = ""
    sSQL = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber,substring(vExamineeName + '　　　　　　　　',1,10) as vExamineeName,iSex,iRoomProfileId"
    sSQL = sSQL & " FROM tbSTEExamineeProfile WHERE iNendo = " & g_int_CurrentNendo
    sSQL = sSQL & " AND exists ( SELECT 1 FROM tbSTEScoreProfile"
    sSQL = sSQL & " WHERE tbSTEScoreProfile.iExamineeProfileId = tbSTEExamineeProfile.iExamineeProfileId "
    sSQL = sSQL & " AND iSubjectProfileId in (SELECT iSubjectProfileId FROM tbSTESubjectProfile"

    Select Case cboSubject.Text
    Case prvcSubName_Science
        sSQL = sSQL & " WHERE vSubjectName in ('物理' , '化学' , '生物'  ) ) "
    Case prvcSubName_Language
        sSQL = sSQL & " WHERE vSubjectName in ('仏語' , '独語' , '英語' ) ) "
    Case prvcSubName_SecondExam
        sSQL = sSQL & " WHERE vSubjectName in ('面接Ⅰ' , '面接Ⅱ' , '小論文' ) ) "
    Case Else
        sSQL = sSQL & " WHERE vSubjectName = '" & cboSubject.Text & "' ) "
    End Select

    sSQL = sSQL & " AND iAbsentFlag=1)"

    If m_int_Action = 0 Then
        sSQL = sSQL & " AND iRoomProfileId=" & cboRoomID.Text & " "
    End If

    Select Case m_int_Action
    Case 0  ' input absentee in the 1st exam phase
        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default

    Case 2  ' input absentee in the 2nd exam phase
        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    End Select
        
    Set oRs = g_obj_Conn.Execute(sSQL)
    
    If oRs.EOF Then
        txtTotal.Text = lstSelected.ListCount
        Set oRs = Nothing
        Exit Sub
    End If
    
    Do While Not oRs.EOF
        strDisp = oRs.Fields("iJukenNumber").Value & _
            " - " & oRs.Fields("vExamineeName").Value
        

        If oRs.Fields("iSex").Value = 0 Then
            strDisp = strDisp & " (*)"
        Else
            strDisp = strDisp & "    "
        End If
                
        If Trim(oRs.Fields("iRoomProfileId").Value) <> "" Then
            sSQLRoomName = "SELECT vRoomName FROM tbSTERoomProfile" & _
                " WHERE iRoomProfileId=" & oRs.Fields("iRoomProfileId").Value
            oRs_RoomName.Open sSQLRoomName, g_obj_Conn
            
            If Not oRs_RoomName.EOF Then
                strDisp = strDisp & " - " & oRs_RoomName.Fields("vRoomName").Value
            End If
            
            oRs_RoomName.Close
            Set oRs_RoomName = Nothing
        End If
        
        lstSelected.AddItem strDisp
        oRs.MoveNext
    Loop

    oRs.Close
    Set oRs = Nothing

    txtTotal.Text = lstSelected.ListCount

End Sub


Public Sub f_void_LoadRoom()        'populate the room names

    On Error GoTo ErrorHandler

    Dim adoRs    As New ADODB.Recordset
    Dim sSQL     As String
    
    sSQL = "SELECT iRoomProfileid,vRoomName FROM tbSTERoomProfile" & _
        " WHERE iMaxCapacity > 0 "
    
    If m_int_IntRpt = 0 Then    ' change made on 31/07/02
        sSQL = sSQL & " AND iInterviewRoomFlag = 0"
    Else
        sSQL = sSQL & " AND iInterviewRoomFlag = 1"
    End If
    
    sSQL = sSQL & " ORDER BY iRoomProfileId"

'-------------------------------------------------------------------------------
'2021.12.14 add jhi
'SELECT
'    iRoomProfileid
'   ,vRoomName
'From
'    tbSTERoomProfile
'Where
'        iMaxCapacity > 0
'    AND iInterviewRoomFlag = 1
'Order By
'    iRoomProfileid
'-------------------------------------------------------------------------------
    
    adoRs.Open sSQL, g_obj_Conn

    Do While Not adoRs.EOF
        cboRoomID.AddItem adoRs.Fields("iRoomProfileid").Value    'hidden combo to keep the id's of rooms
        cboRoom.AddItem adoRs.Fields("vRoomName").Value           'combo which displays the rooms names
        adoRs.MoveNext
    Loop
    
    If cboRoom.ListCount > 0 Then
        cboRoom.ListIndex = 0
        cboRoomID.ListIndex = 0
        lblErrorDetails.Caption = ""
    Else
        lblErrorDetails.Caption = LoadResString(2010)
        Unload Me
    End If

    adoRs.Close
    Set adoRs = Nothing

    Exit Sub

ErrorHandler:
        MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)
End Sub

'*******************************************************************************
'* 2次受験生 List                                                              *
'* 2022.01.11 update jhi                                                       *
'*******************************************************************************
Private Sub cmdJukenList_Click()

    On Error GoTo ErrorHandler

    Dim fso                   As Object
    Dim objText               As Object
    Dim strFile               As String
    Dim blnOpenFile           As Boolean

    Dim sJukenNo              As String
    Dim sJukenNm              As String
    Dim icnt                  As Integer

    Dim strLine               As String



    If lstExaminees.ListCount < 1 Then
        cmdJukenList.Enabled = False
        Exit Sub
    End If

    cmdJukenList.Enabled = True


    blnOpenFile = False

    'FSOオブジェクットを初期化
    strFile = App.Path & "\Report\2次受験生一覧" & Format(Now(), "YYYYMMDDHHmmSS") & ".csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.CreateTextFile(strFile, True, False)

    blnOpenFile = True

    sJukenNo = ""    '受験生No
    sJukenNm = ""    '受験名


    'ファイルを出力
    For icnt = 0 To lstExaminees.ListCount - 1
        sJukenNo = Left(lstExaminees.List(icnt), 4)

        sJukenNm = Mid(lstExaminees.List(icnt), 7, 8)
        sJukenNm = Trim(sJukenNm)
        strLine = icnt + 1 & "," & g_int_CurrentNendo & "," & sJukenNo & "," & sJukenNm
        objText.WriteLine (strLine)
    Next

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If

    ShellExecute Me.hwnd, "open", strFile, vbNullString, vbNullString, 1

    Exit Sub


ErrorHandler:

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If
    MsgBox Err.Description, vbInformation, "2次受験生一覧表"


End Sub


'*******************************************************************************
'* 2次 欠席者 List                                                             *
'* 2022.01.15 add jhi                                                          *
'*******************************************************************************
Private Sub cmdKessekiList_Click()

    On Error GoTo ErrorHandler

    Dim fso                   As Object
    Dim objText               As Object
    Dim strFile               As String
    Dim blnOpenFile           As Boolean

    Dim sJukenNo              As String
    Dim sJukenNm              As String
    Dim icnt                  As Integer

    Dim strLine               As String


    If lstSelected.ListCount < 1 Then
        cmdKessekiList.Enabled = False
        Exit Sub
    End If

    cmdKessekiList.Enabled = True

    blnOpenFile = False

    'FSOオブジェクットを初期化
    strFile = App.Path & "\Report\2次欠席者一覧_" & Format(Now(), "YYYYMMDDHHmmSS") & ".csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.CreateTextFile(strFile, True, False)

    blnOpenFile = True

    sJukenNo = ""    '受験生No
    sJukenNm = ""    '受験名


    '---------------------------------------------------------------------------
    '設定パラメータをファイルに出力
    '---------------------------------------------------------------------------
''''    strLine = "科目: " & cboSubject.Text & "," & ",,会場名: " & cboRoom.Text
''''    objText.WriteLine (strLine)


    '---------------------------------------------------------------------------
    'Headerをファイルに出力
    '---------------------------------------------------------------------------
    strLine = "No,年度,受験番号,受験生名"
    objText.WriteLine (strLine)


    '---------------------------------------------------------------------------
    'ListBoxの内容をファイルに出力
    '---------------------------------------------------------------------------
    For icnt = 0 To lstSelected.ListCount - 1
        sJukenNo = Left(lstSelected.List(icnt), 4)

        sJukenNm = Mid(lstSelected.List(icnt), 7, 8)
        sJukenNm = Trim(sJukenNm)
        strLine = icnt + 1 & "," & g_int_CurrentNendo & "," & sJukenNo & "," & sJukenNm
        objText.WriteLine (strLine)
    Next

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If

    ShellExecute Me.hwnd, "open", strFile, vbNullString, vbNullString, 1

    Exit Sub


ErrorHandler:

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If
    MsgBox Err.Description, vbInformation, "2次欠席者一覧表"

End Sub

