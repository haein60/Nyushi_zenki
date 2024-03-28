VERSION 5.00
Begin VB.Form frmManualAllocationGrp 
   Caption         =   "frmManualAllocationGrp : 面接グループ変更 "
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmManualAllocationGrp.frx":0000
   ScaleHeight     =   9810
   ScaleWidth      =   12435
   Tag             =   "2431"
   WindowState     =   2  '最大化
   Begin VB.ComboBox cboSplRoomFrom 
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
      Left            =   1305
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   20
      Top             =   1965
      Width           =   2280
   End
   Begin VB.ComboBox cboSplDayFrom 
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
      Left            =   4770
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   19
      Top             =   1095
      Width           =   2265
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
      Left            =   7710
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   510
      Visible         =   0   'False
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
      Left            =   10290
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   7440
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
      Left            =   1305
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1095
      Width           =   2280
   End
   Begin VB.ListBox lstExaminee 
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
      Height          =   4740
      Left            =   240
      MultiSelect     =   2  '拡張
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2400
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
      Left            =   8370
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.ComboBox cboSplRoom 
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
      Left            =   8040
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1965
      Width           =   2280
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5340
      TabIndex        =   8
      Top             =   4815
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5325
      TabIndex        =   7
      Top             =   5415
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5325
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5325
      TabIndex        =   5
      Top             =   4215
      Width           =   1095
   End
   Begin VB.ListBox lstAllotted 
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
      Height          =   4740
      Left            =   6960
      MultiSelect     =   2  '拡張
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "会場名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   240
      TabIndex        =   22
      Top             =   2010
      Width           =   930
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "面接日"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   3885
      TabIndex        =   21
      Top             =   1155
      Width           =   855
   End
   Begin VB.Label lblMsg 
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
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   7920
      Width           =   11295
   End
   Begin VB.Label lblDayTotal 
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6945
      TabIndex        =   17
      Top             =   525
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblDayRoomTotal 
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   225
      TabIndex        =   15
      Top             =   7485
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
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "定員"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   9120
      TabIndex        =   12
      Top             =   525
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  '透明
      Caption         =   "科目名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1155
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "面接日"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   7230
      TabIndex        =   10
      Top             =   1260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "会場名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   6960
      TabIndex        =   9
      Top             =   2010
      Width           =   930
   End
End
Attribute VB_Name = "frmManualAllocationGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmManualAllocationGrp
'Author         :   Dileep Cherian
'Created On     :   16/05/02
'Description    :   This form is used for maual allocation of examiees for 2nd phase interviews
'Reference      :   Functional Specs Of Manual Allocation Ver 1.0
'***************************************************************************************************
Option Explicit

Dim f_dt_SplDay             As Date       ' to store the selected spl interview/report day
Dim f_int_SplDayMax         As Long       ' to store the max capacity of the selected interview/report day
Dim f_int_SplDayCount       As Long       ' counter to check the number of examinees allocated to the selected day
Dim f_int_SplRoomMax        As Long       ' max capacity of the selected spl interview/report room
Dim g_ExamDay1Count         As Long       ' count of examinees allocated to the selected spl interview/report room
Dim f_bln_DataChange        As Boolean    ' variable to indicate any change operations
Dim f_int_ExamType          As Long       ' to identify the exam type
Dim f_str_RoomStatus        As String     ' to store the room status before refreshing
Dim f_str_RoomFromStatus    As String     ' to store the room status before refreshing
Dim f_int_SplFromRoomMax    As Long       ' max capacity of the selected spl interview/report room
Dim f_int_SplFromRoomCount  As Long       ' count of examinees allocated to the selected spl interview/report room

Private prviCboFromIndex    As Long       'Fromグループ指定に変更があったかなかったかをチェック
Private prviCboIndex        As Long       'Toグループ指定に変更があったかなかったかをチェック
Private prvbStarttime       As Boolean    '画面ロード時のみの処理のため使用

Public gcnt                 As Integer
'*******************************************************************************
'* frmManualAllocationGrp : 面接グループ変更 Form_Load                         *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler


Call log("▼▼▼ Form_Load START(frmManualAllocationGrp) ▼▼▼")


    ''''2021.11.30 cboSplRoomFrom_Click 関数2回動作するので1回にするため設定flag
    gcnt = 0

    prvbStarttime = True

    LoadResStrings Me
    g_void_SetFontProperties Me


''''2021.11.26 add -> del mnuManualAllocationGrp_Click()に入れた。
''''If MsgBox("面接グループ変更処理を行いますか？", vbQuestion + vbOKCancel, "入試システム") <> vbOK Then
''''    Exit Sub
''''End If


    lblDayTotal.Caption = "該当日付受験者数"                  ''''LoadResString(2487)
    lblDayRoomTotal.Caption = "該当日付・グループ受験者数"    ''''LoadResString(2488)

    lstExaminee.Font = "ＭＳ ゴシック"
    lstAllotted.Font = "ＭＳ ゴシック"


    '---------------------------------------
    'cboSubject comboに科目をaddする
    '---------------------------------------
    Call Add_cboSubject
    cboSubject.ListIndex = 0 ''''2021.11.30 add jhi -- これをいれるとclick eventが発生するのだ

    '---------------------------------------
    ' cboSplDayFrom comboに面接日をaddする
    '---------------------------------------
    Call Add_cboSplDayFrom
    cboSplDayFrom.ListIndex = 0 ''''2021.11.30 add jhi -- これをいれるとclick eventが発生するのだ

    '---------------------------------------
    ' cboSplRoomFrom comboに会場名をaddする
    '---------------------------------------
    Call Add_cboSplRoomFrom(cboSplDayFrom.Text)
    cboSplRoomFrom.ListIndex = 0 ''''2021.11.30 add jhi -- これをいれるとclick eventが発生するのだ


    prvbStarttime = False



    '---------------------------------------
    ' ボタンの状態を変える関数-->たいした処理はなし
    '---------------------------------------
    Call f_void_CheckButtonStatus



Call log("▲▲▲ Form_Load END  (frmManualAllocationGrp) ▲▲▲")

    Exit Sub

ErrorHandler:
    prvbStarttime = False
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'*******************************************************************************
'* frmManualAllocationGrp : 面接グループ変更 Form_Activate                     *
'*******************************************************************************
Private Sub Form_Activate()

    Dim i As Long

'Call log("==== Form_Activate START ====")


    fMainForm.mnuTools.Enabled = False

    For i = 1 To fMainForm.Toolbar1.Buttons.Count
        fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next


'Call log("==== Form_Activate END   ====")

End Sub

'*******************************************************************************
'* frmManualAllocationGrp : 面接科目を取得する関数                             *
'* 取得した科目名 : 面接Ⅰ、面接Ⅱ                                             *
'*******************************************************************************
Private Sub Add_cboSubject()

    On Error GoTo ErrorHandler

    Dim strSQL     As String                 ' SQL string
    Dim oRs  As New ADODB.Recordset    ' recordset object


Call log("==== Add_cboSubject START ====")

    cboSubject.Clear


    strSQL = "SELECT vSubjectName FROM tbSTESubjectProfile"

'    If g_int_ExamType = 1 Then
'        strSQL = strSQL & " WHERE iExamType = 2" '面接のみ
'    Else
'        strSQL = strSQL & " WHERE iExamType in ( 2 , 4 )"    '面接とGroup面接
'        strSQL = strSQL & " WHERE iExamType = 4 "            'Group面接のみ
         strSQL = strSQL & " WHERE iSubType  = 3 "            'Group面接のみ。---> 面接Ⅰ、面接Ⅱ
'    End If

'''' SELECT * FROM tbSTESubjectProfile where iSubType=3; ''''2021.11.26 add jhi

    oRs.Open strSQL, g_obj_Conn


    '科目をcboSubjectに入れる
    Do While Not oRs.EOF
        ''''cboSubjectに科目を設定
        cboSubject.AddItem oRs.Fields("vSubjectName").Value
        oRs.MoveNext
    Loop
    
    If cboSubject.ListCount > 0 Then
        lblMsg.Caption = ""

        ''''del jhi
        ''''cboSubject.ListIndex = 0

        '---------------------------------------
        ' cboSplDayFrom comboに面接日をaddする
        '---------------------------------------
''''        Call Add_cboSplDayFrom
    Else
        lblMsg.Caption = "有効な科目がありません。" ''''LoadResString(2499)
    End If

    oRs.Close
    Set oRs = Nothing

Call log("==== Add_cboSubject END   ====")

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

Private Sub cboSubject_Click()

    On Error GoTo ErrorHandler

    Dim strSQL    As String                 ' SQl string
    Dim oRs As New ADODB.Recordset    ' recordset object
    

''''for debug
Call log("==== cboSubject_Click START ====")
    
    strSQL = "SELECT iExamType FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "'"

    oRs.Open strSQL, g_obj_Conn
    If Not oRs.EOF Then
        f_int_ExamType = oRs.Fields("iExamType").Value ''''面接Ⅰ= 2になる
    End If

    oRs.Close
    Set oRs = Nothing
    
'    Call l_void_PopulateDayCombo
''''Call Add_cboSplDayFrom        ''''2022.01.25 Form_Loadでやるのでここではやめた

    Exit Sub

''''for debug
Call log("==== cboSubject_Click END   ====")


ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)

End Sub

'*******************************************************************************
'* 面接日に第三日が設定されているかを判断して、設定する                        *
'* 2021.11.29 cyosa                                                            *
'*******************************************************************************
Private Sub Add_cboSplDayFrom()

    Dim oRs        As ADODB.Recordset
    Dim sSQL       As String

    Dim bThirdDay  As Boolean


''''for debug
Call log("==== Add_cboSplDayFrom  START ====")


    bThirdDay = False

    sSQL = "SELECT dtSecondExamDay3 FROM tbSTESecondExamProfile "
    sSQL = sSQL & " WHERE iSystemProfileId = "
    sSQL = sSQL & " (SELECT iSystemProfileId FROM tbSTESystemProfile WHERE iActiveFlag=1) "

'----------------------------------------------------------------------------------------------
'for debug '2021.11.29 cyosa ---> nullが返ってくる
'----------------------------------------------------------------------------------------------
'SELECT
'    --dtSecondExamDay3
'    *
'FROM tbSTESecondExamProfile
'WHERE iSystemProfileId = (SELECT iSystemProfileId FROM tbSTESystemProfile WHERE iActiveFlag=1)
'----------------------------------------------------------------------------------------------

    Set oRs = g_obj_Conn.Execute(sSQL)

    If Not oRs.EOF Then

        '面接日に3回が設定されているか？
        If Not IsNull(oRs.Fields(0)) Then
            bThirdDay = True
        End If
    End If

    oRs.Close
    Set oRs = Nothing


    '---------------------------------------------------------------------------
    '以下の処理が走ると、
    '以下のPrivate Sub cboSplDayFrom_Click()
    '関数が自動動作するのだ！
    '---------------------------------------------------------------------------
    With cboSplDayFrom
        .Clear
        .AddItem "第一日"                   ''''LoadResString(2424)
        .AddItem "第二日"                   ''''LoadResString(2425)
        If bThirdDay Then .AddItem "第三日" ''''LoadResString(2426)
''''jhi        .ListIndex = 0                      ''''★これを指定することにより、cboSplDayFrom_Click()関数が自動実行される
    End With


''''for debug
Call log("==== Add_cboSplDayFrom  END   ====")


End Sub

'*******************************************************************************
'* 面接日 combo                                                                *
'*-----------------------------------------------------------------------------*
'* 2021.11.29 cyosa                                                            *
'*******************************************************************************
Private Sub cboSplDayFrom_Click()


''''2022.01.25 add jhi 画面表示でかなり時間がかかるので改修
''''Exit Sub


Call log("===▼ cboSplDayFrom_Click START ▼===")


    prviCboFromIndex = -1
    prviCboIndex = -1

'   Call Add_cboSplRoom(cboSplDayFrom.Text)
    Call Add_cboSplRoomFrom(cboSplDayFrom.Text)

    If cboSplRoomFrom.ListCount > 0 Then
        prvbStarttime = True
        cboSplRoomFrom.ListIndex = 0 ''''これがあるから右側の会場名が表示される 2022.01.25 確認
        prvbStarttime = False
    Else
        Label4.Visible = False
        lblSourceCapacity.Visible = False
    End If

Call log("===▲ cboSplDayFrom_Click END   ▲===")

End Sub

Private Sub Add_cboSplRoomFrom(ByVal l_str_vDay As String)

    On Error GoTo ErrorHandler

    ' fill the room combo based on the day selected in the day combo
    Dim oRs             As New ADODB.Recordset    'recordset object
    Dim l_str_Sql       As String                 'SQL string
    Dim l_int_NoOfRooms As Long                   'to store the number of rooms
    Dim l_int_Counter   As Long                   'counter
    


''''for debug
Call log("●● Add_cboSplRoomFrom START ●●")

    
    cboSplRoomFrom.Clear
    
    ' get the current selected day and room, and their capacities
    l_str_Sql = "SELECT iNumberOfRoomDay1, iNumberOfRoomDay2, iNumberOfRoomDay3," & _
        " dtSecondExamDay1, dtSecondExamDay2, dtSecondExamDay3," & _
        " iNumberOfExamineeDay1, iNumberOfExamineeDay2, iNumberOfExamineeDay3" & _
        " FROM tbSTESecondExamProfile" & _
        " WHERE iSystemProfileId = (" & _
        " SELECT iSystemProfileId FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag = 1)"
    
    oRs.Open l_str_Sql, g_obj_Conn

    If Not oRs.EOF Then

        Select Case UCase(l_str_vDay)
        Case "第一日"    ''''UCase(LoadResString(2424))
            f_dt_SplDay = oRs("dtSecondExamDay1")
            f_int_SplDayMax = oRs("iNumberOfExamineeDay1")
            l_int_NoOfRooms = oRs("iNumberOfRoomDay1")

        Case "第二日"    ''''UCase(LoadResString(2425))
            f_dt_SplDay = oRs("dtSecondExamDay2")
            f_int_SplDayMax = oRs("iNumberOfExamineeDay2")
            l_int_NoOfRooms = oRs("iNumberOfRoomDay2")

        Case "第三日"    ''''UCase(LoadResString(2426))
            If Not IsNull(oRs("dtSecondExamDay3")) Then
                f_dt_SplDay = oRs("dtSecondExamDay3")
            End If

            If IsNull(oRs("iNumberOfExamineeDay3")) Then
                f_int_SplDayMax = 0
            Else
                f_int_SplDayMax = oRs("iNumberOfExamineeDay3")
            End If

            If IsNull(oRs("iNumberOfRoomDay3")) Then
                l_int_NoOfRooms = 0
            Else
                l_int_NoOfRooms = oRs("iNumberOfRoomDay3")
            End If
        End Select

    End If

    oRs.Close
    Set oRs = Nothing
    
    ' to check whether the max capacity of the room is reached or not
    ' change made on 31/07/02
    l_str_Sql = "SELECT iRoomProfileId, vRoomName FROM tbSTERoomProfile" & _
        " WHERE iInterviewRoomFlag = 0" & _
        " ORDER BY iRoomProfileId"

    oRs.Open l_str_Sql, g_obj_Conn
    
    l_int_Counter = 1

    Do While Not oRs.EOF And l_int_Counter <= l_int_NoOfRooms

        ''''会場名のcomboに会場名を設定
        cboSplRoomFrom.AddItem oRs("vRoomName")
        l_int_Counter = l_int_Counter + 1

''''for debug
''''Call log("cnt" & CStr(l_int_Counter) & ":" & oRs("vRoomName"))

        oRs.MoveNext
    Loop


    If cboSplRoomFrom.ListCount > 0 Then
''        Label4.Visible = True
''        lblSourceCapacity.Visible = True
''        cboSplRoomFrom.ListIndex = 0
    Else
        Label4.Visible = False
        lblSourceCapacity.Visible = False
        oRs.Close
        Set oRs = Nothing
        Exit Sub
    End If
        
    oRs.Close
    Set oRs = Nothing
    
    ' to check whether the max capacity of the day is reached or not
    l_str_Sql = "SELECT r.iExamineeProfileId FROM tbSTEExamineeRoomProfile r, tbSTEExamineeProfile e"

    Select Case UCase(l_str_vDay)
    Case "第一日"    ''''UCase(LoadResString(2424))
        l_str_Sql = l_str_Sql & " WHERE CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay1,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"

    Case "第二日"    ''''UCase(LoadResString(2425))
        l_str_Sql = l_str_Sql & " WHERE CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay2,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"

    Case "第三日"    ''''UCase(LoadResString(2426))
        l_str_Sql = l_str_Sql & " WHERE CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay3,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"
    End Select
    
    l_str_Sql = l_str_Sql & " AND r.iSubjectProfileId = (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & cboSubject.Text & "')" & _
        "  AND e.iExamineeProfileId = r.iExamineeProfileId"
    
    oRs.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly

    If Not oRs.EOF Then
        f_int_SplDayCount = oRs.RecordCount
    Else
        f_int_SplDayCount = 0
    End If

 
    txtTotalExamineesDay.Text = f_int_SplDayCount

Call log("txtTotalExamineesDay.Text ---->" & f_int_SplDayCount)


Call log("●● Add_cboSplRoomFrom END   ●●")

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

'*******************************************************************************
'* 会場名 combo 処理                                                           *
'*******************************************************************************
Private Sub cboSplRoomFrom_Click()

    Dim bExec As Boolean

Call log("○ cboSplRoomFrom_Click START ○")

    gcnt = gcnt + 1

    If gcnt > 1 Then
        Call log("---- ★cboSplRoomFrom_Click end(2回だから) ----")
        gcnt = 0
        Exit Sub

    End If

    '---------------------------------------------------------------------------
    'かなり時間がかかる。(1:30程度) 改善が必要！
    '---------------------------------------------------------------------------
    Call l_void_PopulateFromList(cboSplRoomFrom.Text, f_dt_SplDay)


'   lblSourceCapacity.Caption = CStr(f_int_SplFromRoomMax)
    If cboSplRoom.ListCount > 0 Then
        If cboSplRoom.ListIndex = cboSplRoomFrom.ListIndex Then
            bExec = True
        End If
    End If

    If prviCboIndex = -1 Then
        bExec = True
    End If

    If bExec Then


        If prviCboFromIndex <> -1 Or (prviCboFromIndex = -1 And prvbStarttime) Then

            prviCboIndex = -1

            Call Add_cboSplRoom(cboSplDayFrom.Text)

            If cboSplRoom.ListCount > 0 Then
                cboSplRoom.ListIndex = IIf(cboSplRoomFrom.ListIndex = 0, 1, 0)
                ''''この次cboSplRoom_Clickが走る
Call log("cboSplRoom.ListIndex=" & cboSplRoom.ListIndex & " --->この次cboSplRoom_Clickが走る")
            End If

        End If

    End If

    prviCboFromIndex = cboSplRoomFrom.ListIndex


Call log("○ cboSplRoomFrom_Click END   ○")


End Sub

'---------------------------------------------------------------------------
'かなり時間がかかる。(1:30程度) 改善が必要！
'---------------------------------------------------------------------------
Private Sub l_void_PopulateFromList(ByVal l_str_vRoom As String, ByVal l_dt_dtDay As Date)

    On Error GoTo ErrorHandler

    ' populate the list box based on selection made in the day and room combos
    Dim oRs        As New ADODB.Recordset       ' recordset object
    Dim l_obj_rsExaminee As New ADODB.Recordset       ' recordset object
    Dim l_str_Sql        As String                    ' SQL string
    Dim strSQL           As String                    ' SQL string
    
    
    lstExaminee.Clear
    f_int_SplFromRoomCount = 0


Call log("1. ---- l_void_PopulateFromList start ----")

    strSQL = " SELECT iExamineeProfileId FROM tbSTEExamineeProfile as ep where exists ("
    strSQL = strSQL & " SELECT iExamineeProfileId FROM tbSTEExamineeRoomProfile as er "
    strSQL = strSQL & " WHERE iRoomProfileId=("
    strSQL = strSQL & " SELECT iRoomProfileId FROM tbSTERoomProfile"
    strSQL = strSQL & " WHERE vRoomName='" & l_str_vRoom & "')"
    strSQL = strSQL & " AND iSubjectProfileId = (SELECT iSubjectProfileId FROM tbSTESubjectProfile"
    strSQL = strSQL & " WHERE vSubjectName='" & cboSubject.Text & "')"
    strSQL = strSQL & " AND not exists ( select 1 from tbSTEScoreProfile as sc where sc.iExamineeProfileID = er.iExamineeProfileId and sc.iSubjectProfileId = er.iSubjectProfileId and iAbsentFlag=1 ) "

    ''''2022.01.25 add jhi S 4:30かかる現象を直す
    strSQL = strSQL & " AND ep.iNendo=" & g_int_CurrentNendo
    strSQL = strSQL & " AND convert(varchar(4),er.dtCreate,112)='" & g_int_CurrentNendo & "'"
    ''''2022.01.25 add jhi E

    strSQL = strSQL & " AND er.iExamineeProfileId = ep.iExamineeProfileId ) "
    strSQL = strSQL & " AND iNendo=" & g_int_CurrentNendo

Call log("2. ---- l_void_PopulateFromList strSQL   ----")
Call log("3. strSQL=" & strSQL)

    oRs.Open strSQL, g_obj_Conn
''''oRs.Open strSQL, g_obj_Conn, adOpenStatic, adLockReadOnly
''''oRs.Open strSQL, g_obj_Conn, adOpenForwardOnly, adLockReadOnly ''''ODBC Timeout



Call log("4. strSQL OPEN ato")

    
    Do While Not oRs.EOF
        l_str_Sql = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber, substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName,iPreferenceDay1Flag,iPreferenceDay2Flag,iPreferenceDay3Flag,iSex" & _
            " FROM tbSTEExamineeProfile" & _
            " WHERE iExamineeProfileId=" & oRs.Fields("iExamineeProfileId").Value & _
            " AND dtSecondExamDay='" & Format(l_dt_dtDay, "MM/DD/YYYY") & "'" & _
            " AND iExamineeStatus = " & gclExamineeStatus_1stPass & _
            " AND iNendo=" & g_int_CurrentNendo

Call log("4.1. l_str_Sql=" & l_str_Sql)

        l_obj_rsExaminee.Open l_str_Sql, g_obj_Conn

        If Not l_obj_rsExaminee.EOF Then
            lstExaminee.AddItem l_obj_rsExaminee.Fields("iJukenNumber").Value & _
            " - " & l_obj_rsExaminee.Fields("vExamineeName").Value & _
            " -" & l_obj_rsExaminee.Fields("iPreferenceDay1Flag").Value & _
            "-" & l_obj_rsExaminee.Fields("iPreferenceDay2Flag").Value & _
            "-" & l_obj_rsExaminee.Fields("iPreferenceDay3Flag").Value & _
            "-" & IIf(l_obj_rsExaminee.Fields("iSex") = 0, "(*)", "")
            g_ExamDay1Count = g_ExamDay1Count + 1
        End If

        l_obj_rsExaminee.Close
        Set l_obj_rsExaminee = Nothing
        
        oRs.MoveNext
    Loop

Call log("5. Do While ato")

    
    oRs.Close
    Set oRs = Nothing
    
    l_str_Sql = "SELECT iMaxCapacity FROM tbSTERoomProfile WHERE vRoomName='" & _
        cboSplRoom.Text & "'"
        
    oRs.Open l_str_Sql, g_obj_Conn

    If Not oRs.EOF Then
        f_int_SplFromRoomMax = oRs("ImaxCapacity")
    Else
        f_int_SplFromRoomMax = 0
    End If
    
    oRs.Close
    Set oRs = Nothing
    
'    Call f_void_PopulateExaminee
    txtTotal.Text = lstAllotted.ListCount

Call log("6. ---- l_void_PopulateFromList end   ----")

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub Add_cboSplRoom(ByVal l_str_vDay As String)

    On Error GoTo ErrorHandler

    ' fill the room combo based on the day selected in the day combo
    Dim oRs       As New ADODB.Recordset    ' recordset object
    Dim l_str_Sql       As String                 ' SQL string
    Dim l_int_NoOfRooms As Long                   ' to store the number of rooms
    Dim l_int_Counter   As Long                   ' counter
    

''''for debug
Call log("==== Add_cboSplRoom START ====")
    
    cboSplRoom.Clear
    
    ' get the current selected day and room, and their capacities
    l_str_Sql = "SELECT iNumberOfRoomDay1, iNumberOfRoomDay2, iNumberOfRoomDay3," & _
        " dtSecondExamDay1, dtSecondExamDay2, dtSecondExamDay3," & _
        " iNumberOfExamineeDay1, iNumberOfExamineeDay2, iNumberOfExamineeDay3" & _
        " FROM tbSTESecondExamProfile" & _
        " WHERE iSystemProfileId = (" & _
        " SELECT iSystemProfileId FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag = 1)"
    
    oRs.Open l_str_Sql, g_obj_Conn

    If Not oRs.EOF Then
        Select Case UCase(l_str_vDay)
        Case UCase(LoadResString(2424))
            f_dt_SplDay = oRs("dtSecondExamDay1")
            f_int_SplDayMax = oRs("iNumberOfExamineeDay1")
            l_int_NoOfRooms = oRs("iNumberOfRoomDay1")
        Case UCase(LoadResString(2425))
            f_dt_SplDay = oRs("dtSecondExamDay2")
            f_int_SplDayMax = oRs("iNumberOfExamineeDay2")
            l_int_NoOfRooms = oRs("iNumberOfRoomDay2")
        Case UCase(LoadResString(2426))
            f_dt_SplDay = oRs("dtSecondExamDay3")
            f_int_SplDayMax = oRs("iNumberOfExamineeDay3")
            l_int_NoOfRooms = oRs("iNumberOfRoomDay3")
        End Select
    End If
    oRs.Close
    Set oRs = Nothing
    
    ' to check whether the max capacity of the room is reached or not
    ' change made on 31/07/02
    l_str_Sql = "SELECT iRoomProfileId, vRoomName FROM tbSTERoomProfile" & _
        " WHERE iInterviewRoomFlag = 0" & _
        " ORDER BY iRoomProfileId"
    oRs.Open l_str_Sql, g_obj_Conn
    
    l_int_Counter = 1
    Do While Not oRs.EOF And l_int_Counter <= l_int_NoOfRooms
        cboSplRoom.AddItem oRs("vRoomName")
        l_int_Counter = l_int_Counter + 1
        oRs.MoveNext
    Loop
     
    If cboSplRoom.ListCount > 0 Then
'        Label4.Visible = True
'        lblSourceCapacity.Visible = True
'        cboSplRoom.ListIndex = 0
    Else
        Label4.Visible = False
        lblSourceCapacity.Visible = False
        oRs.Close
        Set oRs = Nothing
        Exit Sub
    End If
        
    oRs.Close
    Set oRs = Nothing
    
    ' to check whether the max capacity of the day is reached or not
    l_str_Sql = "SELECT r.iExamineeProfileId FROM tbSTEExamineeRoomProfile r, tbSTEExamineeProfile e"
    Select Case UCase(l_str_vDay)
    Case UCase(LoadResString(2424))
        l_str_Sql = l_str_Sql & " WHERE CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay1,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"
    Case UCase(LoadResString(2425))
        l_str_Sql = l_str_Sql & " WHERE CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay2,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"
    Case UCase(LoadResString(2426))
        l_str_Sql = l_str_Sql & " WHERE CONVERT(VARCHAR(10),e.dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay3,101) FROM tbSTESecondExamProfile" & _
            " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
            " WHERE iActiveFlag=1))"
    End Select
    
    l_str_Sql = l_str_Sql & " AND r.iSubjectProfileId = (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & cboSubject.Text & "')" & _
        "  AND e.iExamineeProfileId = r.iExamineeProfileId"
    
    oRs.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
    If Not oRs.EOF Then
        f_int_SplDayCount = oRs.RecordCount
    Else
        f_int_SplDayCount = 0
    End If
    txtTotalExamineesDay.Text = f_int_SplDayCount

''''for debug
Call log("==== Add_cboSplRoom END   ====")


    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub


'*******************************************************************************
'* f_void_CheckButtonStatus                                                    *
'* ボタンの状態をEnable or Disableにする関数                                   *
'*******************************************************************************
Public Sub f_void_CheckButtonStatus()

    'Procedure to check the status of the buttons
    'i.e enabling and disabling the buttons based on the presense
    'and selection of data in the list boxes

''''for debug
Call log("==== f_void_CheckButtonStatus START ====")


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


''''for debug
Call log("==== f_void_CheckButtonStatus END   ====")


End Sub


Private Sub cmdDeselect_Click()

    'on the click of this button only the Interviewer selected from the lstExaminee
    ' will be transfered to lstAllotted
    Dim l_bln_existing As Boolean           ' to see whether the examinee is already existing or not
    Dim l_int_Counter As Long             ' counter
    Dim l_int_Count As Long               ' counter
    Dim l_bln_Flag As Boolean               ' to see whether the examinee is already existing or not
    Dim l_int_Start As Long               ' to extract the juken number from the combined string
    Dim l_int_End As Long                 ' to extract the juken number from the combined string
    Dim l_int_JukenNo As Long             ' to store the juken number
    Dim l_int_RetVal As Long              ' to track the return value of the function call
    Dim l_bln_Status As Boolean             ' to track the return value of the function call
    Dim l_str_Sql As String                 ' SQL string
    Dim oRs As New ADODB.Recordset    ' recordset object
    Dim l_int_ExamineeId As Long          ' to store the examinee Id

    On Error GoTo ErrorHandler


''''for debug
Call log("==== cmdDeselect_Click START ====")


    If lstAllotted.SelCount > 0 Then
        For l_int_Count = 0 To lstAllotted.ListCount - 1
            If l_int_Count > lstAllotted.ListCount - 1 Then Exit For
            If lstAllotted.Selected(l_int_Count) Then
                For l_int_Counter = 0 To lstExaminee.ListCount - 1
                    If lstExaminee.List(l_int_Counter) = lstAllotted.List(l_int_Count) Then
                        l_bln_existing = True
                        Exit For
                    End If
                Next
                
                If Not l_bln_existing Then
                    l_int_JukenNo = Left(lstAllotted.List(l_int_Count), 4)
                    If f_int_SplFromRoomCount + 1 <= f_int_SplFromRoomMax Then
                        l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                            " WHERE iJukenNumber=" & l_int_JukenNo & _
                            " AND iNendo=" & g_int_CurrentNendo
                        oRs.Open l_str_Sql, g_obj_Conn
                        If Not oRs.EOF Then
                            l_int_ExamineeId = oRs("iExamineeProfileId")
                        End If
                        oRs.Close
                        Set oRs = Nothing
                        
                        l_bln_Status = f_bln_UpdateDatabase(l_int_ExamineeId, 0)
                        If l_bln_Status Then
                            If Not f_bln_DataChange Then f_bln_DataChange = True
                            lstExaminee.AddItem lstAllotted.List(l_int_Count)
                            lstAllotted.RemoveItem (l_int_Count)
                            f_int_SplFromRoomCount = f_int_SplFromRoomCount + 1
                            g_ExamDay1Count = g_ExamDay1Count - 1
                            l_int_Count = l_int_Count - 1   ' because an item is removed from the list
                        End If
                    Else
                        MsgBox LoadResString(2419), vbCritical
                    End If
                End If
            End If
        Next
    End If
    f_void_CheckButtonStatus
    txtTotal.Text = lstAllotted.ListCount
    f_str_RoomFromStatus = cboSplRoomFrom.Text  ' save the room status
    f_str_RoomStatus = cboSplRoom.Text  ' save the room status
    ' refresh the room combo after an examinee is moved from one list bot to another
'    Call Add_cboSplRoom(cboSplDayFrom.Text)
'    Call Add_cboSplRoomFrom(cboSplDayFrom.Text)
    Call l_void_PopulateFromList(cboSplRoomFrom.Text, f_dt_SplDay)
    Call l_void_PopulateList(cboSplRoom.Text, f_dt_SplDay)
'    cboSplRoom.Text = f_str_RoomStatus  ' set the status back
    Call ls_cboSplRoomFromTextSet(f_str_RoomFromStatus)
    Call ls_cboSplRoomTextSet(f_str_RoomStatus)

''''for debug
Call log("==== cmdDeselect_Click END   ====")

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub ls_cboSplRoomTextSet(psText As String)

    Dim ii As Long

''''for debug
Call log("==== ls_cboSplRoomTextSet START ====")

    If cboSplRoom.ListCount < 0 Then
        Exit Sub
    End If

    For ii = 0 To cboSplRoom.ListCount - 1
        If cboSplRoom.List(ii) = psText Then
            cboSplRoom.ListIndex = ii
            Exit For
        End If
    Next

''''for debug
Call log("==== ls_cboSplRoomTextSet END   ====")


End Sub

Private Sub cmdDeselectAll_Click()

    'On the click of this button all the Interviewers from the lstExaminee
    'will be transfered to lstAllotted
    Dim l_int_AllExaminee As Long                   ' counter
    Dim l_int_Start       As Long                   ' to extract the juken number from the combined string
    Dim l_int_End         As Long                   ' to extract the juken number from the combined string
    Dim l_int_JukenNo     As Long                   ' to store the juken number
    Dim l_bln_Status      As Boolean                ' to track the return value odf the function call
    Dim l_str_Sql         As String                 ' SQL string
    Dim oRs         As New ADODB.Recordset    ' recordset object
    Dim l_int_ExamineeId  As Long                   ' to store the examinee Id
    
    On Error GoTo ErrorHandler
        
''''for debug
Call log("==== cmdDeselectAll_Click START ====")


    If lstAllotted.ListCount >= 1 Then
        For l_int_AllExaminee = 0 To lstAllotted.ListCount - 1
            If l_int_AllExaminee > lstAllotted.ListCount - 1 Then Exit For
            
            l_int_JukenNo = Left(lstAllotted.List(l_int_AllExaminee), 4)
                            
            l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                " WHERE iJukenNumber=" & l_int_JukenNo & _
                " AND iNendo=" & g_int_CurrentNendo

            oRs.Open l_str_Sql, g_obj_Conn

            If Not oRs.EOF Then
                l_int_ExamineeId = oRs("iExamineeProfileId")
            End If

            oRs.Close
            Set oRs = Nothing
            
            l_bln_Status = f_bln_FreeExaminee(l_int_ExamineeId)
            If l_bln_Status Then
                If Not f_bln_DataChange Then f_bln_DataChange = True
                lstExaminee.AddItem lstAllotted.List(l_int_AllExaminee)
                lstAllotted.RemoveItem (l_int_AllExaminee)
                f_int_SplDayCount = f_int_SplDayCount - 1
                l_int_AllExaminee = l_int_AllExaminee - 1   ' because an item is removed from the list
            Else
                MsgBox "手続きを完了できませんでした。しばらくしてからもう一度試してください。" ''''LoadResString(2416)
            End If
        Next
    End If

    f_void_CheckButtonStatus
    txtTotal.Text = lstAllotted.ListCount
    f_str_RoomStatus = cboSplRoom.Text  ' save the room status

    ' refresh the room combo after an examinee is moved from one list bot to another
    Call Add_cboSplRoom(cboSplDay.Text)
    Call Add_cboSplRoomFrom(cboSplDayFrom.Text)

''''for debug
Call log("==== cmdDeselectAll_Click END   ====")

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)
End Sub

Private Sub cmdSelectAll_Click()

    'On the click of this button all the Interviewers from the lstExaminee
    'will be transfered to lstAllotted
    Dim l_bln_existing    As Boolean                ' to see whether the examinee is already existing or not
    Dim l_int_Counter     As Long                   ' counter
    Dim l_int_AllExaminee As Long                   ' counter
    Dim l_int_Start       As Long                   ' to extract the juken number from the combined string
    Dim l_int_End         As Long                   ' to extract the juken number from the combined string
    Dim l_int_JukenNo     As Long                   ' to store the juken number
    Dim l_bln_Flag        As Boolean                ' to track the return value of the function call
    Dim l_bln_Status      As Boolean                ' to track the return value of the function call
    Dim l_int_RetVal      As Long                   ' to track the return value of the function call
    Dim l_str_Sql         As String                 ' SQL string
    Dim oRs         As New ADODB.Recordset    ' recordset object
    Dim l_int_ExamineeId  As Long                   ' to store the examinee Id
    
    On Error GoTo ErrorHandler

''''for debug
Call log("==== cmdSelectAll_Click START ====")

        
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
                            If g_ExamDay1Count + 1 <= f_int_SplRoomMax Then
                                l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                                    " WHERE iJukenNumber=" & l_int_JukenNo & _
                                    " AND iNendo=" & g_int_CurrentNendo
                                oRs.Open l_str_Sql, g_obj_Conn
                                If Not oRs.EOF Then
                                    l_int_ExamineeId = oRs("iExamineeProfileId")
                                End If
                                oRs.Close
                                Set oRs = Nothing
                                
                                l_bln_Status = f_bln_UpdateDatabase(l_int_ExamineeId, 0)
                                If l_bln_Status Then
                                    If Not f_bln_DataChange Then f_bln_DataChange = True
                                    lstAllotted.AddItem lstExaminee.List(l_int_AllExaminee)
                                    lstExaminee.RemoveItem (l_int_AllExaminee)
                                    g_ExamDay1Count = g_ExamDay1Count + 1
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
                        If g_ExamDay1Count + 1 <= f_int_SplRoomMax Then
                            l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                                " WHERE iJukenNumber=" & l_int_JukenNo & _
                                " AND iNendo=" & g_int_CurrentNendo
                            oRs.Open l_str_Sql, g_obj_Conn
                            If Not oRs.EOF Then
                                l_int_ExamineeId = oRs("iExamineeProfileId")
                            End If
                            oRs.Close
                            Set oRs = Nothing
                            
                            l_bln_Status = f_bln_UpdateDatabase(l_int_ExamineeId, 0)
                            If l_bln_Status Then
                                If Not f_bln_DataChange Then f_bln_DataChange = True
                                lstAllotted.AddItem lstExaminee.List(l_int_AllExaminee)
                                lstExaminee.RemoveItem (l_int_AllExaminee)
                                g_ExamDay1Count = g_ExamDay1Count + 1
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
    f_str_RoomStatus = cboSplRoom.Text  ' save the room status
    ' refresh the room combo after an examinee is moved from one list bot to another
    Call Add_cboSplRoom(cboSplDayFrom.Text)
    Call Add_cboSplRoomFrom(cboSplDayFrom.Text)
    cboSplRoom.Text = f_str_RoomStatus  ' set the status back

''''for debug
Call log("==== cmdSelectAll_Click END   ====")


    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub cmdSelect_Click()

    'on the click of this button only the Interviewer selected from the lstExaminee
    ' will be transfered to lstAllotted
    Dim l_bln_existing   As Boolean                ' to see whether the examinee is already existing or not
    Dim l_int_Counter    As Long                   ' counter
    Dim l_int_Count      As Long                   ' counter
    Dim l_bln_Flag       As Boolean                ' to see whether the examinee is already existing or not
    Dim l_int_Start      As Long                   ' to extract the juken number from the combined string
    Dim l_int_End        As Long                   ' to extract the juken number from the combined string
    Dim l_int_JukenNo    As Long                   ' to store the juken number
    Dim l_int_RetVal     As Long                   ' to track the return value of the function call
    Dim l_bln_Status     As Boolean                ' to track the return value of the function call
    Dim l_str_Sql        As String                 ' SQL string
    Dim oRs        As New ADODB.Recordset    ' recordset object
    Dim l_int_ExamineeId As Long                   ' to store the examinee Id
    
    On Error GoTo ErrorHandler


''''for debug
Call log("==== cmdSelect_Click START ====")

    
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
                    If f_int_SplDayCount + 1 <= f_int_SplDayMax Then
                        If g_ExamDay1Count + 1 <= f_int_SplRoomMax Then
                            l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                                " WHERE iJukenNumber=" & l_int_JukenNo & _
                                " AND iNendo=" & g_int_CurrentNendo
                            oRs.Open l_str_Sql, g_obj_Conn
                            If Not oRs.EOF Then
                                l_int_ExamineeId = oRs("iExamineeProfileId")
                            End If
                            oRs.Close
                            Set oRs = Nothing
                            
                            l_bln_Status = f_bln_UpdateDatabase(l_int_ExamineeId, 1)
                            If l_bln_Status Then
                                If Not f_bln_DataChange Then f_bln_DataChange = True
                                lstAllotted.AddItem lstExaminee.List(l_int_Count)
                                lstExaminee.RemoveItem (l_int_Count)
                                g_ExamDay1Count = g_ExamDay1Count + 1
                                f_int_SplFromRoomCount = f_int_SplFromRoomCount - 1
                                l_int_Count = l_int_Count - 1   ' because an item is removed from the list
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
    f_str_RoomFromStatus = cboSplRoomFrom.Text  ' save the room status
    f_str_RoomStatus = cboSplRoom.Text          ' save the room status

    ' refresh the room combo after an examinee is moved from one list bot to another
'    Call Add_cboSplRoom(cboSplDayFrom.Text)
'    Call Add_cboSplRoomFrom(cboSplDayFrom.Text)

    Call l_void_PopulateFromList(cboSplRoomFrom.Text, f_dt_SplDay)
    Call l_void_PopulateList(cboSplRoom.Text, f_dt_SplDay)

'    cboSplRoom.Text = f_str_RoomStatus  ' set the status back

    Call ls_cboSplRoomFromTextSet(f_str_RoomFromStatus)
    Call ls_cboSplRoomTextSet(f_str_RoomStatus)


''''for debug
Call log("==== cmdSelect_Click END   ====")


    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub ls_cboSplRoomFromTextSet(psText As String)

    Dim ii As Long


''''for debug
Call log("==== ls_cboSplRoomFromTextSet START ====")


    If cboSplRoomFrom.ListCount < 0 Then
        Exit Sub
    End If

    For ii = 0 To cboSplRoomFrom.ListCount - 1
        If cboSplRoomFrom.List(ii) = psText Then
            cboSplRoomFrom.ListIndex = ii
            Exit For
        End If
    Next

''''for debug
Call log("==== ls_cboSplRoomFromTextSet END   ====")


End Sub

Private Function f_bln_UpdateDatabase(ByVal iExamineeId As Long, iRoomFlag As Long) As Boolean

    ' update the database with the current changes
    ' value has to be inserted in tbSTEExamineeRoomProfile
    ' also updation in tbSTEExamineeProfile
    Dim l_str_Sql         As String
    Dim oRs         As New ADODB.Recordset
    Dim oRs1        As New ADODB.Recordset
    Dim oRs2        As New ADODB.Recordset
    Dim l_int_NewId       As Long
    Dim l_int_SubjectId   As Long
    Dim l_int_RoomId      As Long
    Dim l_int_ExamDate    As Date
    Dim l_dt_IntvDate     As Date
    Dim l_int_Counter     As Long
    Dim l_int_LoopCounter As Long
    Dim l_str_SubjId()    As String

    Dim bRtn              As Boolean
    
    On Error GoTo ErrorHandler
    

''''for debug
Call log("==== f_bln_UpdateDatabase START ====")


    g_obj_Conn.BeginTrans

    l_str_Sql = "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & cboSubject.Text & "'"
    oRs.Open l_str_Sql, g_obj_Conn

    If Not oRs.EOF Then
        l_int_SubjectId = oRs.Fields("iSubjectProfileId").Value
    End If

    oRs.Close
    Set oRs = Nothing
    
    l_str_Sql = "SELECT iRoomProfileId FROM tbSTERoomProfile" & _
        " WHERE vRoomName='" & IIf(iRoomFlag = 0, cboSplRoomFrom.Text, cboSplRoom.Text) & "'"
    oRs.Open l_str_Sql, g_obj_Conn

    If Not oRs.EOF Then
        l_int_RoomId = oRs("iRoomProfileId")
    End If
    oRs.Close
    Set oRs = Nothing
        
    l_str_Sql = "SELECT iExamineeRoomProfileId FROM tbSTEExamineeRoomProfile" & _
        " WHERE iExamineeProfileId=" & iExamineeId & _
        " AND iSubjectProfileId =" & l_int_SubjectId

    oRs.Open l_str_Sql, g_obj_Conn

    If Not oRs.EOF Then
        l_str_Sql = "UPDATE tbSTEExamineeRoomProfile" & _
            " SET iRoomProfileId = " & l_int_RoomId & "," & _
            " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
            " WHERE iExamineeProfileId=" & iExamineeId & _
            " AND iSubjectProfileId =" & l_int_SubjectId
        g_obj_Conn.Execute l_str_Sql

    Else
'        l_str_Sql = "SELECT iExamineeRoomProfileId FROM tbSTEExamineeRoomProfile"
'        oRs1.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
'        If Not oRs1.EOF Then
'            oRs1.MoveLast
'            l_int_NewId = oRs1("iExamineeRoomProfileId") + 1
'        Else
'            l_str_Sql = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping" & _
'                " WHERE vTableName='tbSTEExamineeRoomProfile'"
'            oRs2.Open l_str_Sql, g_obj_Conn
'            If Not oRs2.EOF Then
'                l_int_NewId = oRs2("iTableCounterIdMapping")
'            Else
'                l_int_NewId = 1
'            End If
'            oRs2.Close
'            Set oRs2 = Nothing
'        End If
'        oRs1.Close
'        Set oRs1 = Nothing


        bRtn = getNewId("tbSTEExamineeRoomProfile", "iExamineeRoomProfileId", l_int_NewId)

        ' insert data into tbSTEExamineeroomProfile table
        l_str_Sql = "INSERT INTO tbSTEExamineeRoomProfile VALUES(" & _
            l_int_NewId & "," & _
            iExamineeId & "," & _
            l_int_RoomId & "," & _
            l_int_SubjectId & ",'" & _
            Format(Date, "MM/DD/YYYY") & "','" & _
            Format(Date, "MM/DD/YYYY") & "')"
        g_obj_Conn.Execute l_str_Sql
    End If
        
    oRs.Close
    Set oRs = Nothing
    
    Select Case UCase(cboSplDayFrom.Text)
    Case UCase(LoadResString(2424))
        l_str_Sql = "SELECT dtSecondExamDay1 FROM tbSTESecondExamProfile"
    Case UCase(LoadResString(2425))
        l_str_Sql = "SELECT dtSecondExamDay2 FROM tbSTESecondExamProfile"
    Case UCase(LoadResString(2426))
        l_str_Sql = "SELECT dtSecondExamDay3 FROM tbSTESecondExamProfile"
    End Select
    
    l_str_Sql = l_str_Sql & " WHERE iSystemProfileId=(SELECT iSystemProfileId" & _
        " FROM tbSTESystemProfile WHERE iActiveFlag=1)"
    oRs1.Open l_str_Sql, g_obj_Conn
    If Not oRs1.EOF Then
        l_int_ExamDate = oRs1(0)
    End If
    oRs1.Close
    Set oRs1 = Nothing

'部屋変更
    l_str_Sql = "UPDATE tbSTEExamineeProfile SET iRoomProfileID='" & l_int_RoomId & "'," & _
        " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
        " WHERE iExamineeProfileId=" & iExamineeId
    g_obj_Conn.Execute l_str_Sql
    g_obj_Conn.CommitTrans
    f_bln_UpdateDatabase = True

''''for debug
Call log("==== f_bln_UpdateDatabase END   ====")

    Exit Function

ErrorHandler:
    g_obj_Conn.RollbackTrans
    MsgBox LoadResString(2416), vbInformation, LoadResString(1729)
    f_bln_UpdateDatabase = False

End Function


Private Function f_void_CheckPreferenceViolation(ByVal l_int_iJukenNo As Long) As Boolean

    ' check whteher the movement of selected examinee from unallocated list box & _
    to allocated listbox will cause in violation in the preference day mentioned by & _
    the examinee at the time of registration
    Dim l_str_Sql As String
    Dim oRs As New ADODB.Recordset
    
    On Error GoTo ErrorHandler


''''for debug
Call log("==== f_void_CheckPreferenceViolation START ====")


    l_str_Sql = "SELECT iPreferenceDay1Flag,iPreferenceDay2Flag,iPreferenceDay3Flag FROM tbSTEExamineeProfile" & _
        " WHERE iJukenNumber=" & l_int_iJukenNo & _
        " AND iNendo=" & g_int_CurrentNendo
    
    oRs.Open l_str_Sql, g_obj_Conn

    If Not oRs.EOF Then
        Select Case UCase(cboSplDayFrom.Text)
        Case UCase(LoadResString(2424))
            If oRs("iPreferenceDay1Flag") = 1 Then
                f_void_CheckPreferenceViolation = True
            Else
                f_void_CheckPreferenceViolation = False
            End If
        Case UCase(LoadResString(2425))
            If oRs("iPreferenceDay2Flag") = 1 Then
                f_void_CheckPreferenceViolation = True
            Else
                f_void_CheckPreferenceViolation = False
            End If
        Case UCase(LoadResString(2426))
            If oRs("iPreferenceDay3Flag") = 1 Then
                f_void_CheckPreferenceViolation = True
            Else
                f_void_CheckPreferenceViolation = False
            End If
        End Select
    Else
        f_void_CheckPreferenceViolation = False
    End If
    oRs.Close
    Set oRs = Nothing

''''for debug
Call log("==== f_void_CheckPreferenceViolation END   ====")

    Exit Function

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
    f_void_CheckPreferenceViolation = False

End Function

Private Sub l_void_PopulateDayCombo()

    Dim sSQL       As String
    Dim oRs        As ADODB.Recordset
    Dim bThirdDay  As Boolean

''''for debug
Call log("==== l_void_PopulateDayCombo START ====")

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

''''for debug
Call log("==== l_void_PopulateDayCombo END   ====")



End Sub

Private Sub cboSplRoom_Click()  ' for special interview/report


''''for debug
Call log("==== cboSplRoom_Click START ====")

    Call l_void_PopulateList(cboSplRoom.Text, f_dt_SplDay)

    lblSourceCapacity.Caption = CStr(f_int_SplRoomMax)

    If cboSplRoomFrom.ListCount > 0 Then
        If cboSplRoomFrom.ListIndex = cboSplRoom.ListIndex Then
            If prviCboIndex <> -1 Then
                prviCboFromIndex = -1
                Call Add_cboSplRoomFrom(cboSplDayFrom.Text)
                If cboSplRoomFrom.ListCount > 0 Then
                    cboSplRoomFrom.ListIndex = IIf(cboSplRoom.ListIndex = 0, 1, 0)
                End If
            End If
        End If
    End If

    prviCboIndex = cboSplRoom.ListIndex

''''for debug
Call log("==== cboSplRoom_Click END   ====")

End Sub

Private Sub l_void_PopulateList(ByVal l_str_vRoom As String, ByVal l_dt_dtDay As Date)

    On Error GoTo ErrorHandler

    Dim oRs              As New ADODB.Recordset ''''2021.11.30 時間短縮調査
    Dim l_obj_rsExaminee As New ADODB.Recordset       ' recordset object
    Dim l_str_Sql        As String                    ' SQL string
    Dim strSQL           As String                    ' SQL string



    

''''for debug
Call log("==== l_void_PopulateList START ====")

    
    lstAllotted.Clear
    g_ExamDay1Count = 0


    ''''------------------------------------------------------------------------
    ''''以下のsql文はかなり時間がかかる 2021.11.29 cyosa
    ''''------------------------------------------------------------------------
    strSQL = ""
    strSQL = " SELECT iExamineeProfileId FROM tbSTEExamineeProfile as ep where exists ("
    strSQL = strSQL & " SELECT iExamineeProfileId FROM tbSTEExamineeRoomProfile as er "
    strSQL = strSQL & " WHERE iRoomProfileId=("
    strSQL = strSQL & " SELECT iRoomProfileId FROM tbSTERoomProfile"
    strSQL = strSQL & " WHERE vRoomName='" & l_str_vRoom & "')"
    strSQL = strSQL & " AND iSubjectProfileId = (SELECT iSubjectProfileId FROM tbSTESubjectProfile"
    strSQL = strSQL & " WHERE vSubjectName='" & cboSubject.Text & "')"

    ''''2022.01.25 add jhi S 4:30かかる現象を直す
    strSQL = strSQL & " AND ep.iNendo=" & g_int_CurrentNendo
    strSQL = strSQL & " AND convert(varchar(4),er.dtCreate,112)='" & g_int_CurrentNendo & "'"
    ''''2022.01.25 add jhi E

    strSQL = strSQL & " AND er.iExamineeProfileId = ep.iExamineeProfileId ) "
    strSQL = strSQL & " AND ep.iNendo=" & g_int_CurrentNendo
    strSQL = strSQL & " AND ep.iAbsentFlag=0"

''''2021.11.30 del
    oRs.Open strSQL, g_obj_Conn


    '2021.11.30 add jhi
''''    strSQL = ""
''''    strSQL = strSQL & "SELECT" & vbCrLf
''''    strSQL = strSQL & "    ep.iExamineeProfileId iExamineeProfileId" & vbCrLf
''''    strSQL = strSQL & "FROM" & vbCrLf
''''    strSQL = strSQL & "    tbSTEExamineeProfile     ep" & vbCrLf
''''    strSQL = strSQL & "   ,tbSTEExamineeRoomProfile er" & vbCrLf
''''    strSQL = strSQL & "   ,tbSTERoomProfile         rp" & vbCrLf
''''    strSQL = strSQL & "   ,tbSTESubjectProfile      sp" & vbCrLf
''''    strSQL = strSQL & "WHERE" & vbCrLf
''''    strSQL = strSQL & "        ep.iNendo             = " & g_int_CurrentNendo & vbCrLf
''''    strSQL = strSQL & "    and ep.iAbsentFlag        =0" & vbCrLf
''''    strSQL = strSQL & "    and er.iExamineeProfileId = ep.iExamineeProfileId" & vbCrLf
''''    strSQL = strSQL & "    and er.iRoomProfileId     = rp.iRoomProfileId" & vbCrLf
''''    strSQL = strSQL & "    and rp.vRoomName          ='" & l_str_vRoom & "'" & vbCrLf
''''    strSQL = strSQL & "    and er.iSubjectProfileId  = sp.iSubjectProfileId" & vbCrLf
''''    strSQL = strSQL & "    and sp.vSubjectName       ='" & cboSubject.Text & "'" & vbCrLf
''''    strSQL = strSQL & "ORDER BY ep.iExamineeProfileId"

''''oRs.Open strSQL, g_obj_Conn

    Do While Not oRs.EOF

        l_str_Sql = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber, substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName,iPreferenceDay1Flag,iPreferenceDay2Flag,iPreferenceDay3Flag,iSex" & _
            " FROM tbSTEExamineeProfile" & _
            " WHERE iExamineeProfileId=" & oRs.Fields("iExamineeProfileId").Value & _
            " AND dtSecondExamDay='" & Format(l_dt_dtDay, "MM/DD/YYYY") & "'" & _
            " AND iExamineeStatus = " & gclExamineeStatus_1stPass & _
            " AND iNendo=" & g_int_CurrentNendo

        l_obj_rsExaminee.Open l_str_Sql, g_obj_Conn

        If Not l_obj_rsExaminee.EOF Then
            lstAllotted.AddItem l_obj_rsExaminee.Fields("iJukenNumber").Value & _
            " - " & l_obj_rsExaminee.Fields("vExamineeName").Value & _
            " -" & l_obj_rsExaminee.Fields("iPreferenceDay1Flag").Value & _
            "-" & l_obj_rsExaminee.Fields("iPreferenceDay2Flag").Value & _
            "-" & l_obj_rsExaminee.Fields("iPreferenceDay3Flag").Value & _
            "-" & IIf(l_obj_rsExaminee.Fields("iSex") = 0, "(*)", "")
            g_ExamDay1Count = g_ExamDay1Count + 1
        End If

        l_obj_rsExaminee.Close
        Set l_obj_rsExaminee = Nothing

        oRs.MoveNext

    Loop
    
    oRs.Close
    Set oRs = Nothing
    
    l_str_Sql = "SELECT iMaxCapacity FROM tbSTERoomProfile WHERE vRoomName='" & _
        cboSplRoom.Text & "'"
        
    oRs.Open l_str_Sql, g_obj_Conn
    If Not oRs.EOF Then
        f_int_SplRoomMax = oRs("ImaxCapacity")
    Else
        f_int_SplRoomMax = 0
    End If
    
    oRs.Close
    Set oRs = Nothing
    
'    Call f_void_PopulateExaminee
    txtTotal.Text = lstAllotted.ListCount

''''for debug
Call log("==== l_void_PopulateList END   ====")

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "l_void_PopulateList:エラー"

End Sub


Private Sub f_void_PopulateExaminee()

    ' pick up all the unallocated examinees and populate them in the & _
    unllocated (left) List box
    Dim l_int_Count   As Long                   ' counter
    Dim l_str_Arr()   As String                 ' to store the examinee id of all examinees
    Dim l_int_Start   As Long                   ' to extract the juken number
    Dim l_int_End     As Long                   ' to extract the juken number
    Dim l_int_JukenNo As Long                   ' to store the examinee number
    Dim oRs     As New ADODB.Recordset    ' recordset object
    Dim l_str_Sql     As String                 ' SQL string
    
    On Error GoTo ErrorHandler
    

''''for debug
Call log("==== f_void_PopulateExaminee START ====")


    l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeRoomProfile" & _
        " WHERE iSubjectProfileId=(" & _
        " SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & cboSubject.Text & "')"
    oRs.Open l_str_Sql, g_obj_Conn
    
    Do While Not oRs.EOF
        ReDim Preserve l_str_Arr(l_int_Count)
        l_str_Arr(l_int_Count) = oRs("iExamineeProfileId")
        l_int_Count = l_int_Count + 1
        oRs.MoveNext
    Loop
    
    oRs.Close
    Set oRs = Nothing
            
    l_str_Sql = "SELECT substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName,iJukenNumber,iPreferenceDay1Flag,iPreferenceDay2Flag,iPreferenceDay3Flag,iSex" & _
        " FROM tbSTEExamineeProfile" & _
        " WHERE iAbsentFlag = 0 AND iExamineeStatus = " & gclExamineeStatus_1stPass & _
        " AND iNendo=" & g_int_CurrentNendo & _
        " AND dtsecondexamday=" & f_dt_SplDay
    
    If l_int_Count > 0 Then
        l_str_Sql = l_str_Sql & " AND iExamineeProfileId NOT IN (" & Join(l_str_Arr, ",") & ")"
    End If
    
    With oRs
        .Open l_str_Sql, g_obj_Conn
        lstExaminee.Clear
        
        Do While Not .EOF
            lstExaminee.AddItem g_str_LPad(oRs.Fields("iJukenNumber").Value, Len(oRs.Fields("iJukenNumber").Value)) & _
            " - " & oRs.Fields("vExamineeName").Value & _
            " -" & oRs.Fields("iPreferenceDay1Flag").Value & _
            "-" & oRs.Fields("iPreferenceDay2Flag").Value & _
            "-" & oRs.Fields("iPreferenceDay3Flag").Value & _
            "-" & IIf(oRs.Fields("iSex") = 0, "(*)", "")
            .MoveNext
        Loop
        
        .Close
        Set oRs = Nothing
    End With

''''for debug
Call log("==== f_void_PopulateExaminee END   ====")

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)

End Sub


Private Sub cboSplDay_Click()   ' for special interview/report

'    Call Add_cboSplRoom(cboSplDay.Text)

End Sub

Private Function f_bln_FreeExaminee(ByVal l_int_iExamineeId As Long) As Boolean

    Dim l_str_Sql         As String                 ' SQL string
    Dim oRsExaminee As New ADODB.Recordset    ' recordset object
    Dim l_int_RecCount    As Long                   ' to store the total no of records
    
    On Error GoTo ErrorHandler


    l_str_Sql = "SELECT iExamineeRoomProfileId FROM tbSTEExamineeRoomProfile" & _
        " WHERE iExamineeProfileId=" & l_int_iExamineeId

    oRsExaminee.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
    l_int_RecCount = oRsExaminee.RecordCount
        
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
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)
    f_bln_FreeExaminee = False

End Function

Private Sub Form_Unload(Cancel As Integer)

''''for debug
Call log("=== Form_Unload START   ===")

    Call g_void_CloseChildForm

''''for debug
Call log("=== Form_Unload END   ===")


End Sub

Private Sub lstAllotted_Click()

    Call f_void_CheckButtonStatus

End Sub

Private Sub lstAllotted_DblClick()

    Call cmdDeselect_Click

End Sub

Private Sub lstExaminee_Click()

    Call f_void_CheckButtonStatus

End Sub

Private Sub lstExaminee_DblClick()

    Call cmdSelect_Click

End Sub

