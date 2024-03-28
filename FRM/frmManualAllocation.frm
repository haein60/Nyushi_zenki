VERSION 5.00
Begin VB.Form frmManualAllocation 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14895
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmManualAllocation.frx":0000
   ScaleHeight     =   9810
   ScaleWidth      =   14895
   Tag             =   "2431"
   Visible         =   0   'False
   WindowState     =   2  '最大化
   Begin VB.TextBox txtExamineeID2 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8280
      MaxLength       =   4
      TabIndex        =   24
      Top             =   2025
      Width           =   1095
   End
   Begin VB.CommandButton cmdReDisp 
      Caption         =   "受験生表示"
      Height          =   495
      Left            =   3510
      TabIndex        =   23
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "表示オプション"
      Height          =   660
      Left            =   510
      TabIndex        =   19
      Top             =   1080
      Width           =   2880
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "併願なし"
         Height          =   270
         Left            =   1650
         TabIndex        =   22
         Top             =   285
         Width           =   960
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "併願"
         Height          =   315
         Left            =   870
         TabIndex        =   21
         Top             =   255
         Width           =   705
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "全体"
         Height          =   420
         Left            =   105
         TabIndex        =   20
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.TextBox txtWemenDay1 
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
      Left            =   4605
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   8085
      Width           =   900
   End
   Begin VB.CommandButton cmdJukenList1 
      Caption         =   "受験生リストCSV出力"
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
      Left            =   510
      TabIndex        =   16
      Top             =   8520
      Width           =   2900
   End
   Begin VB.CommandButton cmdJukenList2 
      Caption         =   "受験生リストCSV出力"
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
      Left            =   7095
      TabIndex        =   15
      Top             =   8520
      Width           =   2900
   End
   Begin VB.TextBox txtTotalDay2 
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
      Left            =   8175
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   8085
      Width           =   900
   End
   Begin VB.TextBox txtTotalDay1 
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
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   8085
      Width           =   900
   End
   Begin VB.TextBox txtExamineeID1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1665
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2025
      Width           =   1095
   End
   Begin VB.TextBox txtWemenDay2 
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
      Left            =   11205
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   8085
      Width           =   900
   End
   Begin VB.ListBox lstDay1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   5325
      Left            =   510
      MultiSelect     =   2  '拡張
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2745
      Width           =   5000
   End
   Begin VB.ListBox lstDay2 
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
      Height          =   5325
      Left            =   7110
      MultiSelect     =   2  '拡張
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2745
      Width           =   5000
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
      Left            =   5775
      TabIndex        =   3
      Top             =   5490
      Width           =   1095
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
      Left            =   5760
      TabIndex        =   1
      Top             =   4575
      Width           =   1095
   End
   Begin VB.Label lblMsg2 
      BackStyle       =   0  '透明
      Caption         =   "lblMsg2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   7065
      TabIndex        =   26
      Top             =   9240
      Width           =   7515
   End
   Begin VB.Label lblJuken2 
      BackStyle       =   0  '透明
      Caption         =   "受験番号"
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
      Height          =   225
      Left            =   7110
      TabIndex        =   25
      Top             =   2070
      Width           =   1050
   End
   Begin VB.Label lblWemenDay1 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "女性"
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
      Height          =   300
      Left            =   3780
      TabIndex        =   17
      Top             =   8115
      Width           =   705
   End
   Begin VB.Label lblTotalDay1 
      Alignment       =   1  '右揃え
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
      Left            =   450
      TabIndex        =   13
      Top             =   8115
      Width           =   1020
   End
   Begin VB.Label lblExamDay2 
      Alignment       =   2  '中央揃え
      Caption         =   "試験日2"
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
      Left            =   7110
      TabIndex        =   11
      Top             =   2460
      Width           =   4980
   End
   Begin VB.Label lblExamDay1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "試験日1"
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
      Left            =   510
      TabIndex        =   10
      Top             =   2460
      Width           =   4980
   End
   Begin VB.Label lblJuken1 
      BackStyle       =   0  '透明
      Caption         =   "受験番号"
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
      Height          =   225
      Left            =   510
      TabIndex        =   9
      Top             =   2070
      Width           =   1050
   End
   Begin VB.Label lblWemenDay2 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "女性"
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
      Height          =   300
      Left            =   10410
      TabIndex        =   7
      Top             =   8115
      Width           =   645
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  '透明
      Caption         =   "lblMsg"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   480
      TabIndex        =   5
      Top             =   9240
      Width           =   8130
   End
   Begin VB.Label lblTotalDay2 
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
      Left            =   7125
      TabIndex        =   4
      Top             =   8115
      Width           =   1020
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

Dim f_bln_DataChange   As Boolean    'variable to indicate any change operations
Dim f_int_ExamType     As Long       'to identify the exam type

Dim g_iSystemProfileId As Integer


Dim g_dExamDay1        As Date
Dim g_iExamDay1Max     As Long
Dim g_iExamDay1RoomCnt As Long

Dim g_dExamDay2        As Date
Dim g_iExamDay2Max     As Long
Dim g_iExamDay2RoomCnt As Long

Dim g_dExamDay3        As Date
Dim g_iExamDay3Max     As Long
Dim g_iExamDay3RoomCnt As Long

Dim g_ExamDay1Count    As Long       'count of examinees allocated to the selected spl interview/report room

'指定のウィンドウにメッセージを送る(P750)
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As String) As Long

Private Const LB_FINDSTRING = &H18F         '先頭一致検索(P816)
Private Const LB_FINDSTRINGEXACT = &H1A2    '完全一致検索(P816)



Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Dim kubun    As Integer
    Dim rinf     As Integer


    Call f_void_CheckButtonStatus


    '----------------------
    '表示設定
    '----------------------
    Option1.Value = True
    Option2.Value = False
    Option3.Value = False


    '----------------------
    '初期データ取得関数
    '----------------------
    g_iSystemProfileId = Get_iSystemProfileId(1)
    rinf = Get_Init_Hensu(kubun)

    lblExamDay1.Caption = "試験日1 (" & Format(g_dExamDay1, "YYYY/MM/DD") & ")"
    lblExamDay2.Caption = "試験日2 (" & Format(g_dExamDay2, "YYYY/MM/DD") & ")"


    '----------------------
    '初期データ取得関数
    '----------------------
    lstDay1.Clear ''''2022.01.15 add jhi
    lstDay2.Clear ''''2022.01.15 add jhi

    lblMsg.Caption = ""
    lblMsg2.Caption = ""
 

    Call Set_lstDay1(1)
    Call Set_lstDay2(1)

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"


End Sub

'*******************************************************************************
'* 受験生再表示                                                                *
'*******************************************************************************
Private Sub cmdReDisp_Click()

    lblMsg.Caption = ""
    lblMsg2.Caption = ""

''''Debug.Print "option1.Value=" & Option1.Value & " option2.Value=" & Option2.Value & " option3.Value=" & Option3.Value

   If (Option1.Value = True) Then
      Call Set_lstDay1(1)
      Call Set_lstDay2(1)

   ElseIf (Option2.Value = True) Then
      Call Set_lstDay1(2)
      Call Set_lstDay2(2)

   ElseIf (Option3.Value = True) Then
      Call Set_lstDay1(3)
      Call Set_lstDay2(3)

   End If


End Sub

'*******************************************************************************
'* ListBoxにデータを設定                                                       *
'*******************************************************************************
Private Sub Set_lstDay1(kubun As Integer)

    On Error GoTo ErrorHandler

    Dim oRs     As New ADODB.Recordset    ' recordset object
    Dim sSQL    As String                 ' SQL string
    
    
    lstDay1.Clear
    g_ExamDay1Count = 0


    sSQL = ""
    sSQL = Make_Day1_sSQL(kubun)

    oRs.Open sSQL, g_obj_Conn

    Do While Not oRs.EOF
        lstDay1.AddItem oRs.Fields("iJukenNumber").Value & _
        " - " & oRs.Fields("vExamineeName").Value & _
        " -" & oRs.Fields("iPreferenceDay1Flag").Value & _
        "-" & oRs.Fields("iPreferenceDay2Flag").Value & _
        "-" & oRs.Fields("iMultipleApplyFlag").Value & _
        "-" & IIf(oRs.Fields("iSex") = 0, "(*)", "")

        g_ExamDay1Count = g_ExamDay1Count + 1
        oRs.MoveNext
    Loop

    oRs.Close
    Set oRs = Nothing


    txtTotalDay1.Text = lstDay1.ListCount
    txtWemenDay1.Text = Get_WemenCount(1)


    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

'******************************************************************************
'* ListBox lstDay2を設定する                                                  *
'******************************************************************************
Private Sub Set_lstDay2(kubun As Integer)

    On Error GoTo ErrorHandler

    Dim oRs           As New ADODB.Recordset    ' recordset object
    Dim sSQL          As String                 ' SQL string
    

    lstDay2.Clear


    sSQL = ""
    sSQL = Make_Day2_sSQL(kubun)

    oRs.Open sSQL, g_obj_Conn
        
    Do While Not oRs.EOF
        lstDay2.AddItem oRs.Fields("iJukenNumber").Value & _
        " - " & oRs.Fields("vExamineeName").Value & _
        " -" & oRs.Fields("iPreferenceDay1Flag").Value & _
        "-" & oRs.Fields("iPreferenceDay2Flag").Value & _
        "-" & oRs.Fields("iMultipleApplyFlag").Value & _
        "-" & IIf(oRs.Fields("iSex") = 0, "(*)", "")
         oRs.MoveNext
    Loop
       
    oRs.Close
    Set oRs = Nothing

    txtTotalDay2.Text = lstDay2.ListCount
    txtWemenDay2.Text = Get_WemenCount(2)

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

'*******************************************************************************
'* 【>】 選択処理                                                              *
'*******************************************************************************
Private Sub cmdSelect_Click()

    On Error GoTo ErrorHandler

    Dim oRs              As New ADODB.Recordset    ' recordset object
    Dim sSQL             As String                 ' SQL string

    Dim i                As Long                   ' counter
    Dim icnt             As Long                   ' counter

    Dim l_bln_existing   As Boolean                ' to see whether the examinee is already existing or not
    Dim l_bln_Flag       As Boolean                ' to see whether the examinee is already existing or not
    Dim l_bln_Status     As Boolean                ' to track the return value of the function call
    Dim borinf           As Boolean                ' to track the return value of the function call

    Dim iJukenNo         As Long                   ' to store the juken number
    Dim iExamineeId      As Long                   ' to store the examinee Id
    Dim rinf             As Long                   ' to track the return value of the function call

    Dim sMsg             As String


    
    If lstDay1.SelCount <= 0 Then
        Exit Sub
    End If


    For icnt = 0 To lstDay1.ListCount - 1

        If icnt > lstDay1.ListCount - 1 Then
            Exit For
        End If


        If lstDay1.Selected(icnt) Then

            For i = 0 To lstDay2.ListCount - 1
                If lstDay2.List(i) = lstDay1.List(icnt) Then
                    l_bln_existing = True
                    Exit For
                 End If
            Next i

''''Call log("icnt=" & icnt)

                
            If Not l_bln_existing Then

                 iJukenNo = CLng(Left(lstDay1.List(icnt), 4))
                    
''''2022.02.01 del jhi
''''                 sMsg = ""
''''                 sMsg = sMsg & "選択した受験生の試験日を変更します。(受験番号=" & iJukenNo & ")" & vbCrLf
''''                 sMsg = sMsg & "実行しますか？"
''''
''''                 rinf = MsgBox(sMsg, vbQuestion + vbYesNo, "確認")
''''                 If rinf = vbNo Then
''''                     Exit Sub
''''                 End If

                 iExamineeId = Get_iExamineeProfileId(iJukenNo)

                 borinf = ExamDay_Update(iExamineeId, 2)
                 If borinf Then

                     If Not f_bln_DataChange Then
                         f_bln_DataChange = True
                     End If

                     lstDay2.AddItem lstDay1.List(icnt)
                     lstDay1.RemoveItem (icnt)
                     icnt = icnt - 1   ' because an item is removed from the list

                     lblMsg.Caption = "選択した受験生の試験日を変更しました。(受験番号=" & Format(iJukenNo, "000#") & ")"
                 End If

            End If

        End If

    Next

''''Call log("2 icnt=" & icnt)


    txtTotalDay1.Text = lstDay1.ListCount
    txtWemenDay1.Text = Get_WemenCount(1)

    txtTotalDay2.Text = lstDay2.ListCount
    txtWemenDay2.Text = Get_WemenCount(2)

''''Call cmdReDisp_Click ''''refresh gamen

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

'*******************************************************************************
'* 【<】 選択処理                                                              *
'*******************************************************************************
Private Sub cmdDeselect_Click()

    On Error GoTo ErrorHandler

    Dim oRs              As New ADODB.Recordset    ' recordset object
    Dim sSQL             As String                 ' SQL string

    Dim i                As Long                   ' counter
    Dim icnt             As Long                   ' counter

    Dim l_bln_existing   As Boolean                ' to see whether the examinee is already existing or not
    Dim l_bln_Flag       As Boolean                ' to see whether the examinee is already existing or not
    Dim l_bln_Status     As Boolean                ' to track the return value of the function call
    Dim borinf           As Boolean                ' to track the return value of the function call

    Dim iJukenNo         As Long                   ' to store the juken number
    Dim iExamineeId      As Long                   ' to store the examinee Id
    Dim rinf             As Long                   ' to track the return value of the function call

    Dim sMsg             As String


    
    If lstDay2.SelCount <= 0 Then
        Exit Sub
    End If


''''Call log("ListCount-1=" & lstDay2.ListCount - 1)

    For icnt = 0 To (lstDay2.ListCount - 1)

''''Call log("icnt=" & icnt)

        If icnt > lstDay2.ListCount - 1 Then
            Exit For
        End If


        If lstDay2.Selected(icnt) Then

            For i = 0 To (lstDay1.ListCount - 1)
                If lstDay1.List(i) = lstDay2.List(icnt) Then
                    l_bln_existing = True
                    Exit For
                 End If
            Next i
                
            If Not l_bln_existing Then

                 iJukenNo = CLng(Left(lstDay2.List(icnt), 4))
                    
''''2022.02.01 del jhi
''''                 sMsg = ""
''''                 sMsg = sMsg & "選択した受験生の試験日を変更します。(受験番号=" & iJukenNo & ")" & vbCrLf
''''                 sMsg = sMsg & "実行しますか？"
''''
''''                 rinf = MsgBox(sMsg, vbQuestion + vbYesNo, "確認")
''''                 If rinf = vbNo Then
''''                     Exit Sub
''''                 End If

                 iExamineeId = Get_iExamineeProfileId(iJukenNo)

                 borinf = ExamDay_Update(iExamineeId, 1)
                 If borinf Then

                     If Not f_bln_DataChange Then
                         f_bln_DataChange = True
                     End If

                     lstDay1.AddItem lstDay2.List(icnt)
                     lstDay2.RemoveItem (icnt)

                     'because an item is removed from the list
                     icnt = icnt - 1

                     lblMsg2.Caption = "選択した受験生の試験日を変更しました。(受験番号=" & Format(iJukenNo, "000#") & ")"

                 End If

            End If

        End If

    Next icnt


    txtTotalDay1.Text = lstDay1.ListCount
    txtWemenDay1.Text = Get_WemenCount(1)

    txtTotalDay2.Text = lstDay2.ListCount
    txtWemenDay2.Text = Get_WemenCount(2)

''''Call cmdReDisp_Click ''''refresh gamen

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

'*******************************************************************************
'* 面接データ更新処理                                                          *
'*******************************************************************************
Private Function ExamDay_Update(ByVal iExamineeId As Long, iExamDay_flag As Integer) As Boolean

    On Error GoTo ErrorHandler

    Dim oRs    As New ADODB.Recordset
    Dim sSQL   As String
    
    Dim dtTemp As Date



    g_obj_Conn.BeginTrans


    sSQL = ""
    If iExamDay_flag = 1 Then
        sSQL = sSQL & "UPDATE tbSTEExamineeProfile" & vbCrLf
        sSQL = sSQL & "SET" & vbCrLf
        sSQL = sSQL & "    dtSecondExamDay    = '" & Format(g_dExamDay1, "YYYY/MM/DD") & "'" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay1Flag=1 " & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay2Flag=0 " & vbCrLf
        sSQL = sSQL & "   ,dtUpdate           = '" & Format(Date, "YYYY/MM/DD") & "'" & vbCrLf
        sSQL = sSQL & "WHERE" & vbCrLf
        sSQL = sSQL & "        iExamineeProfileId=" & iExamineeId & vbCrLf
        sSQL = sSQL & "    and iNendo            =" & g_int_CurrentNendo
    Else
        sSQL = sSQL & "UPDATE tbSTEExamineeProfile" & vbCrLf
        sSQL = sSQL & "SET" & vbCrLf
        sSQL = sSQL & "    dtSecondExamDay    = '" & Format(g_dExamDay2, "YYYY/MM/DD") & "'" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay1Flag=0 " & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay2Flag=1 " & vbCrLf
        sSQL = sSQL & "   ,dtUpdate           = '" & Format(Date, "YYYY/MM/DD") & "'" & vbCrLf
        sSQL = sSQL & "WHERE" & vbCrLf
        sSQL = sSQL & "        iExamineeProfileId=" & iExamineeId
        sSQL = sSQL & "    and iNendo            =" & g_int_CurrentNendo
    End If

    g_obj_Conn.Execute sSQL

    g_obj_Conn.CommitTrans

    ExamDay_Update = True
    Exit Function

ErrorHandler:
    g_obj_Conn.RollbackTrans
    MsgBox "２次試験日変更処理でエラーが発生しました。(ExamDay_Update)", vbInformation, "エラー"
    ExamDay_Update = False

End Function

Private Function Get_WemenCount(kubun As Integer) As Integer

    Dim i         As Long
    Dim iWemenCnt As Long

    iWemenCnt = 0

    If kubun = 1 Then
        For i = 0 To lstDay1.ListCount - 1
            If InStr(lstDay1.List(i), "(*)") = 0 Then
                iWemenCnt = iWemenCnt + 1
            End If
        Next
    Else
        For i = 0 To lstDay2.ListCount - 1
            If InStr(lstDay2.List(i), "(*)") = 0 Then
                iWemenCnt = iWemenCnt + 1
            End If
        Next
    End If


    Get_WemenCount = iWemenCnt


End Function

Private Sub Form_Activate()

    Dim i As Long

    fMainForm.mnuTools.Enabled = False

    For i = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next i

End Sub

'2022.01.21 add jhi
Private Function Get_iSystemProfileId(id As Integer) As Integer

    Get_iSystemProfileId = id

End Function

Private Function Get_Init_Hensu(ikubun As Integer) As Integer

    On Error GoTo ErrorHandler

    Dim oRs    As New ADODB.Recordset
    Dim sSQL   As String

    Dim rinf   As Integer


    rinf = 0

    '2次試験日の日付を取得する
    sSQL = ""
    sSQL = sSQL & "SELECT" & vbCrLf
    sSQL = sSQL & "    dtSecondExamDay1" & vbCrLf
    sSQL = sSQL & "   ,dtSecondExamDay2" & vbCrLf
    sSQL = sSQL & "   ,dtSecondExamDay3" & vbCrLf
    sSQL = sSQL & "   ,iNumberOfExamineeDay1" & vbCrLf
    sSQL = sSQL & "   ,iNumberOfExamineeDay2" & vbCrLf
    sSQL = sSQL & "   ,iNumberOfExamineeDay3" & vbCrLf
    sSQL = sSQL & "   ,iNumberOfRoomDay1" & vbCrLf
    sSQL = sSQL & "   ,iNumberOfRoomDay2" & vbCrLf
    sSQL = sSQL & "   ,iNumberOfRoomDay3" & vbCrLf
    sSQL = sSQL & " FROM" & vbCrLf
    sSQL = sSQL & "     tbSTESecondExamProfile" & vbCrLf
    sSQL = sSQL & " WHERE" & vbCrLf
    sSQL = sSQL & "     iSystemProfileId = 1" & vbCrLf

'SELECT
'    dtSecondExamDay1
'   ,dtSecondExamDay2
'   ,dtSecondExamDay3
'   ,iNumberOfExamineeDay1
'   ,iNumberOfExamineeDay2
'   ,iNumberOfExamineeDay3
'   ,iNumberOfRoomDay1
'   ,iNumberOfRoomDay2
'   ,iNumberOfRoomDay3
'From
'    tbSTESecondExamProfile
'Where
'    iSystemProfileId = 1

    oRs.Open sSQL, g_obj_Conn


    '2次試験日を取得する
    If Not oRs.EOF Then

        '試験日: 第一日　日付
        g_dExamDay1 = oRs("dtSecondExamDay1")
        g_iExamDay1Max = oRs("iNumberOfExamineeDay1")
        g_iExamDay1RoomCnt = oRs("iNumberOfRoomDay1")

        '試験日: 第二日　日付
        g_dExamDay2 = oRs("dtSecondExamDay2")
        g_iExamDay2Max = oRs("iNumberOfExamineeDay2")
        g_iExamDay2RoomCnt = oRs("iNumberOfRoomDay2")

        '試験日: 第三日　日付
'        g_dExamDay3 = oRS("dtSecondExamDay3")
'        g_iExamDay3Max = oRS("iNumberOfExamineeDay3")
'        g_iExamDay3RoomCnt = oRS("iNumberOfRoomDay3")

    End If

    oRs.Close
    Set oRs = Nothing


    Get_Init_Hensu = rinf

    Exit Function

ErrorHandler:
    MsgBox Err.Description, vbInformation, "Get_Init_Hensuエラー"


End Function

Private Function Get_iExamineeProfileId(ByVal iJukenNo As Long) As Long

    On Error GoTo ErrorHandler

    Dim oRs            As New ADODB.Recordset
    Dim sSQL           As String
    
    Dim iExamineeId    As Long


    sSQL = ""
    sSQL = sSQL & "SELECT" & vbCrLf
    sSQL = sSQL & "    iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "FROM" & vbCrLf
    sSQL = sSQL & "    tbSTEExamineeProfile" & vbCrLf
    sSQL = sSQL & "WHERE" & vbCrLf
    sSQL = sSQL & "        iNendo      =" & g_int_CurrentNendo & vbCrLf
    sSQL = sSQL & "    and iJukenNumber=" & iJukenNo

    oRs.Open sSQL, g_obj_Conn

    If Not oRs.EOF Then
        iExamineeId = oRs("iExamineeProfileId")
    End If

    oRs.Close
    Set oRs = Nothing

    Get_iExamineeProfileId = iExamineeId
    Exit Function

ErrorHandler:
    MsgBox Err.Description, vbInformation, "Get_iExamineeProfileIdエラー"
    Get_iExamineeProfileId = 0

End Function

'*******************************************************************************
'*                                                                             *
'*******************************************************************************
Public Sub f_void_CheckButtonStatus()

    If lstDay1.ListCount = 0 Then
        cmdSelect.Enabled = False
    Else
        If lstDay1.SelCount > 0 Then
            cmdSelect.Enabled = True
        Else
            cmdSelect.Enabled = False
        End If
    End If

    
    If lstDay2.ListCount = 0 Then
        cmdDeselect.Enabled = False
    Else
        If lstDay2.SelCount > 0 Then
            cmdDeselect.Enabled = True
        Else
            cmdDeselect.Enabled = False
        End If
    End If

End Sub

Private Sub lstDay1_Click()
    Call f_void_CheckButtonStatus
End Sub

Private Sub lstDay2_Click()
    Call f_void_CheckButtonStatus
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call g_void_CloseChildForm

End Sub

'*******************************************************************************
'* lstDay1のデータ抽出SQL文作成                                                *
'*******************************************************************************
Private Function Make_Day1_sSQL(kubun As Integer) As String

    Dim sSQL As String


    sSQL = ""

    If kubun = 1 Then
       
        sSQL = sSQL & "SELECT" & vbCrLf
        sSQL = sSQL & "    dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber" & vbCrLf
        sSQL = sSQL & "   ,substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay1Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay2Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay3Flag" & vbCrLf
        sSQL = sSQL & "   ,iMultipleApplyFlag" & vbCrLf    '併願flag
        sSQL = sSQL & "   ,iSex" & vbCrLf
        sSQL = sSQL & "FROM" & vbCrLf
        sSQL = sSQL & "     tbSTEExamineeProfile" & vbCrLf
        sSQL = sSQL & "WHERE" & vbCrLf
        sSQL = sSQL & "        iNendo          =" & g_int_CurrentNendo & vbCrLf
        sSQL = sSQL & "    AND iExamineeStatus =" & gclExamineeStatus_1stPass & vbCrLf
        sSQL = sSQL & "    AND dtSecondExamDay ='" & Format(g_dExamDay1, "YYYY/MM/DD") & "'"

    ElseIf (kubun = 2) Then '併願者のみ

        sSQL = sSQL & "SELECT" & vbCrLf
        sSQL = sSQL & "    dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber" & vbCrLf
        sSQL = sSQL & "   ,substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay1Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay2Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay3Flag" & vbCrLf
        sSQL = sSQL & "   ,iMultipleApplyFlag" & vbCrLf    '併願flag
        sSQL = sSQL & "   ,iSex" & vbCrLf
        sSQL = sSQL & "FROM" & vbCrLf
        sSQL = sSQL & "     tbSTEExamineeProfile" & vbCrLf
        sSQL = sSQL & "WHERE" & vbCrLf
        sSQL = sSQL & "        iNendo          =" & g_int_CurrentNendo & vbCrLf
        sSQL = sSQL & "    AND iExamineeStatus =" & gclExamineeStatus_1stPass & vbCrLf
        sSQL = sSQL & "    AND dtSecondExamDay ='" & Format(g_dExamDay1, "YYYY/MM/DD") & "'"
        sSQL = sSQL & "    AND iMultipleApplyFlag =1" '併願者

    ElseIf (kubun = 3) Then '併願者ではない

        sSQL = sSQL & "SELECT" & vbCrLf
        sSQL = sSQL & "    dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber" & vbCrLf
        sSQL = sSQL & "   ,substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay1Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay2Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay3Flag" & vbCrLf
        sSQL = sSQL & "   ,iMultipleApplyFlag" & vbCrLf    '併願flag
        sSQL = sSQL & "   ,iSex" & vbCrLf
        sSQL = sSQL & "FROM" & vbCrLf
        sSQL = sSQL & "     tbSTEExamineeProfile" & vbCrLf
        sSQL = sSQL & "WHERE" & vbCrLf
        sSQL = sSQL & "        iNendo          =" & g_int_CurrentNendo & vbCrLf
        sSQL = sSQL & "    AND iExamineeStatus =" & gclExamineeStatus_1stPass & vbCrLf
        sSQL = sSQL & "    AND dtSecondExamDay ='" & Format(g_dExamDay1, "YYYY/MM/DD") & "'"
        sSQL = sSQL & "    AND iMultipleApplyFlag =0" '併願者ではない
    Else
        MsgBox "Make_Day1_sSQL: パラメータエラー(" & kubun & ")"
    End If


    Make_Day1_sSQL = sSQL
  

    Exit Function

End Function

'*******************************************************************************
'* lstDay2のデータ抽出SQL文作成                                                *
'*******************************************************************************
Private Function Make_Day2_sSQL(kubun As Integer) As String

    Dim sSQL As String


    sSQL = ""

    If kubun = 1 Then
       
        sSQL = sSQL & "SELECT" & vbCrLf
        sSQL = sSQL & "    dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber" & vbCrLf
        sSQL = sSQL & "   ,substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay1Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay2Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay3Flag" & vbCrLf
        sSQL = sSQL & "   ,iMultipleApplyFlag" & vbCrLf    '併願flag
        sSQL = sSQL & "   ,iSex" & vbCrLf
        sSQL = sSQL & "FROM" & vbCrLf
        sSQL = sSQL & "     tbSTEExamineeProfile" & vbCrLf
        sSQL = sSQL & "WHERE" & vbCrLf
        sSQL = sSQL & "        iNendo          =" & g_int_CurrentNendo & vbCrLf
        sSQL = sSQL & "    AND iExamineeStatus =" & gclExamineeStatus_1stPass & vbCrLf
        sSQL = sSQL & "    AND dtSecondExamDay ='" & Format(g_dExamDay2, "YYYY/MM/DD") & "'"

    ElseIf (kubun = 2) Then '併願者のみ

        sSQL = sSQL & "SELECT" & vbCrLf
        sSQL = sSQL & "    dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber" & vbCrLf
        sSQL = sSQL & "   ,substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay1Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay2Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay3Flag" & vbCrLf
        sSQL = sSQL & "   ,iMultipleApplyFlag" & vbCrLf    '併願flag
        sSQL = sSQL & "   ,iSex" & vbCrLf
        sSQL = sSQL & "FROM" & vbCrLf
        sSQL = sSQL & "     tbSTEExamineeProfile" & vbCrLf
        sSQL = sSQL & "WHERE" & vbCrLf
        sSQL = sSQL & "        iNendo          =" & g_int_CurrentNendo & vbCrLf
        sSQL = sSQL & "    AND iExamineeStatus =" & gclExamineeStatus_1stPass & vbCrLf
        sSQL = sSQL & "    AND dtSecondExamDay ='" & Format(g_dExamDay2, "YYYY/MM/DD") & "'"
        sSQL = sSQL & "    AND iMultipleApplyFlag =1" '併願者

    ElseIf (kubun = 3) Then '併願者ではない

        sSQL = sSQL & "SELECT" & vbCrLf
        sSQL = sSQL & "    dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber" & vbCrLf
        sSQL = sSQL & "   ,substring( vExamineeName + '　　　　　　　　　　' , 1 , 8 ) as vExamineeName" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay1Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay2Flag" & vbCrLf
        sSQL = sSQL & "   ,iPreferenceDay3Flag" & vbCrLf
        sSQL = sSQL & "   ,iMultipleApplyFlag" & vbCrLf    '併願flag
        sSQL = sSQL & "   ,iSex" & vbCrLf
        sSQL = sSQL & "FROM" & vbCrLf
        sSQL = sSQL & "     tbSTEExamineeProfile" & vbCrLf
        sSQL = sSQL & "WHERE" & vbCrLf
        sSQL = sSQL & "        iNendo          =" & g_int_CurrentNendo & vbCrLf
        sSQL = sSQL & "    AND iExamineeStatus =" & gclExamineeStatus_1stPass & vbCrLf
        sSQL = sSQL & "    AND dtSecondExamDay ='" & Format(g_dExamDay2, "YYYY/MM/DD") & "'"
        sSQL = sSQL & "    AND iMultipleApplyFlag =0" '併願者ではない
    Else
        MsgBox "Make_Day2_sSQL: パラメータエラー(" & kubun & ")"
    End If


    Make_Day2_sSQL = sSQL
  

    Exit Function

End Function


'*******************************************************************************
'* 試験日変更                                                                  *
'* 左ListBoxデータcsv出力                                                      *
'* 2022.01.16 update jhi                                                       *
'*******************************************************************************
Private Sub cmdJukenList1_Click()

    On Error GoTo ErrorHandler

    Dim fso                   As Object
    Dim objText               As Object
    Dim strFile               As String
    Dim blnOpenFile           As Boolean

    Dim sJukenNo              As String
    Dim sJukenNm              As String
    Dim icnt                  As Integer

    Dim strLine               As String


    If lstDay1.ListCount < 1 Then
        cmdJukenList1.Enabled = False
        Exit Sub
    End If

    cmdJukenList1.Enabled = True


    blnOpenFile = False

    'FSOオブジェクットを初期化
    strFile = App.Path & "\Report\受験生一覧_" & "第一日" & "_" & Format(Now(), "YYYYMMDDHHmmSS") & ".csv"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.CreateTextFile(strFile, True, False)

    blnOpenFile = True

    sJukenNo = ""    '受験生No
    sJukenNm = ""    '受験名


    '---------------------------------------------------------------------------
    '設定パラメータをファイルに出力
    '---------------------------------------------------------------------------
    strLine = "面接日: " & g_dExamDay1 & "," & ",,"
    objText.WriteLine (strLine)


    '---------------------------------------------------------------------------
    'Headerをファイルに出力
    '---------------------------------------------------------------------------
    strLine = "No,年度,受験番号,受験生名"
    objText.WriteLine (strLine)


    '---------------------------------------------------------------------------
    'ListBoxの内容をファイルに出力
    '---------------------------------------------------------------------------
    For icnt = 0 To lstDay1.ListCount - 1
        sJukenNo = Left(lstDay1.List(icnt), 4)

        sJukenNm = Mid(lstDay1.List(icnt), 7, 8)
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

    ShellExecute Me.hWnd, "open", strFile, vbNullString, vbNullString, 1

    Exit Sub


ErrorHandler:

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If
    MsgBox Err.Description, vbInformation, "エラー(第一日)"

End Sub

'*******************************************************************************
'* 試験日変更                                                                  *
'* 右ListBoxデータcsv出力                                                      *
'* 2022.01.16 update jhi                                                       *
'*******************************************************************************
Private Sub cmdJukenList2_Click()

    On Error GoTo ErrorHandler

    Dim fso                   As Object
    Dim objText               As Object
    Dim strFile               As String
    Dim blnOpenFile           As Boolean

    Dim sJukenNo              As String
    Dim sJukenNm              As String
    Dim icnt                  As Integer

    Dim strLine               As String
    

    If lstDay2.ListCount < 1 Then
        cmdJukenList2.Enabled = False
        Exit Sub
    End If

    cmdJukenList2.Enabled = True

    blnOpenFile = False


    'FSOオブジェクットを初期化
    strFile = App.Path & "\Report\受験生一覧_" & "第二日" & "_" & Format(Now(), "YYYYMMDDHHmmSS") & ".csv"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.CreateTextFile(strFile, True, False)

    blnOpenFile = True

    sJukenNo = ""    '受験生No
    sJukenNm = ""    '受験名


    '---------------------------------------------------------------------------
    '設定パラメータをファイルに出力
    '---------------------------------------------------------------------------
    strLine = "面接日: " & g_dExamDay2 & "," & ",,"
    objText.WriteLine (strLine)


    '---------------------------------------------------------------------------
    'Headerをファイルに出力
    '---------------------------------------------------------------------------
    strLine = "No,年度,受験番号,受験生名"
    objText.WriteLine (strLine)


    '---------------------------------------------------------------------------
    'ListBoxの内容をファイルに出力
    '---------------------------------------------------------------------------
    For icnt = 0 To lstDay2.ListCount - 1

        sJukenNo = Left(lstDay2.List(icnt), 4)

        sJukenNm = Mid(lstDay2.List(icnt), 7, 8)
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

    ShellExecute Me.hWnd, "open", strFile, vbNullString, vbNullString, 1

    Exit Sub


ErrorHandler:

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If
    MsgBox Err.Description, vbInformation, "エラー(第二日)"

End Sub

'******************************************************************************
'* 受験番号入力で試験日の変更                                                 *
'* 2022.02.01 add jhi                                                         *
'******************************************************************************
Private Sub txtExamineeID1_KeyPress(KeyAscii As Integer)
 
    On Error GoTo ErrorHandler

    Dim oRs              As New ADODB.Recordset    ' recordset object
    Dim sSQL             As String                 ' SQL string

    Dim i                As Long                   ' counter
    Dim icnt             As Long                   ' counter
    Dim idx              As Integer

    Dim l_bln_existing   As Boolean                ' to see whether the examinee is already existing or not
    Dim borinf           As Boolean                ' to track the return value of the function call

    Dim sJukenNo         As String                 ' to store the juken number
    Dim iExamineeId      As Long                   ' to store the examinee Id
    Dim rinf             As Long                   ' to track the return value of the function call

    Dim sMsg             As String
    

    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If


    lblMsg.Caption = ""

    '---------------------------------------------------------------------------
    'Enter Keyの処理
    '---------------------------------------------------------------------------
    If KeyAscii = 13 Then
        
        If Trim(txtExamineeID1.Text) = "" Then
            Exit Sub
        End If

        lblMsg.Caption = ""
        lblMsg2.Caption = ""
    

 
        'リスト件数分ループしてSelectedをFalseにする
        For i = 0 To Me.lstDay1.ListCount - 1
            lstDay1.Selected(i) = False
        Next i


        sJukenNo = Format(txtExamineeID1.Text, "000#")
        idx = fLBSearch(lstDay1, sJukenNo, 0)
        If idx = -1 Then
            MsgBox "指定の受験番号は登録されていません。ご確認ください。"
            Exit Sub
        Else
''''        MsgBox "すでに登録されています。"
            lstDay1.Selected(idx) = True
        End If

        Call cmdSelect_Click

    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

'******************************************************************************
'* 受験番号入力で試験日の変更                                                 *
'* 2022.02.01 add jhi                                                         *
'******************************************************************************
Private Sub txtExamineeID2_KeyPress(KeyAscii As Integer)
 
    On Error GoTo ErrorHandler

    Dim oRs              As New ADODB.Recordset    ' recordset object
    Dim sSQL             As String                 ' SQL string

    Dim i                As Long                   ' counter
    Dim icnt             As Long                   ' counter
    Dim idx              As Integer

    Dim l_bln_existing   As Boolean                ' to see whether the examinee is already existing or not
    Dim borinf           As Boolean                ' to track the return value of the function call

    Dim sJukenNo         As String                 ' to store the juken number
    Dim iExamineeId      As Long                   ' to store the examinee Id
    Dim rinf             As Long                   ' to track the return value of the function call

    Dim sMsg             As String
    

    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If


    lblMsg2.Caption = ""

    '---------------------------------------------------------------------------
    'Enter Keyの処理
    '---------------------------------------------------------------------------
    If KeyAscii = 13 Then
        
        If Trim(txtExamineeID2.Text) = "" Then
            Exit Sub
        End If
 
        lblMsg.Caption = ""
        lblMsg2.Caption = ""


        'リスト件数分ループしてSelectedをFalseにする
        For i = 0 To Me.lstDay1.ListCount - 1
            lstDay1.Selected(i) = False
        Next i


        sJukenNo = Format(txtExamineeID2.Text, "000#")
        idx = fLBSearch(lstDay2, sJukenNo, 0)
        If idx = -1 Then
            MsgBox "指定受験番号は登録されていません。ご確認ください。"
            Exit Sub
        Else
''''        MsgBox "すでに登録されています。"
            lstDay2.Selected(idx) = True
        End If

        Call cmdDeselect_Click

    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub



'=================================================================
' 指定の文字列がリストボックス内にあるか検索する関数
'　LBox　　 　：検索するリストボックス名
'　SearchStr　：検索する文字列
'　Exact　　　：検索方法　1<>前方一致検索　1=完全一致検索
'　fLBSearch　：戻り値　見つかった場合=インデックス それ以外 = -1
'==================================================================
Private Function fLBSearch(LBox As ListBox, ByVal SearchStr As String, Optional ByVal Exact As Integer = 0) As Integer
    Dim Ret As Long
    If Exact = 1 Then
        Ret = SendMessage(LBox.hWnd, LB_FINDSTRINGEXACT, -1, SearchStr)
    Else
        Ret = SendMessage(LBox.hWnd, LB_FINDSTRING, -1, SearchStr)
    End If

    LBox.ListIndex = Ret
    fLBSearch = Ret

End Function

