VERSION 5.00
Begin VB.Form frm2jiHoketusyaJuni 
   AutoRedraw      =   -1  'True
   Caption         =   "frmHoketusyaJuni : 補欠者順位"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm2jiHoketusyaJuni.frx":0000
   ScaleHeight     =   10305
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
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   16
      Top             =   8205
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "補欠者 順位 確定"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6390
      TabIndex        =   8
      Top             =   8760
      Width           =   2205
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "閉じる"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9780
      TabIndex        =   7
      Top             =   8760
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "2次 補欠者リストCSV出力"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   660
      TabIndex        =   6
      Top             =   8520
      Width           =   2800
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "↑"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6210
      TabIndex        =   5
      Top             =   3870
      Width           =   480
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6195
      TabIndex        =   4
      Top             =   5265
      Width           =   480
   End
   Begin VB.ListBox lstKuriage 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   660
      TabIndex        =   3
      Top             =   1695
      Width           =   5325
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "表示"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4680
      TabIndex        =   2
      Top             =   1125
      Width           =   1300
   End
   Begin VB.TextBox txtNendo 
      Alignment       =   2  '中央揃え
      Enabled         =   0   'False
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
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1815
      MaxLength       =   4
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "[iNendo]"
      Top             =   915
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "補欠者数"
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
      Left            =   660
      TabIndex        =   17
      Top             =   8205
      Width           =   1200
   End
   Begin VB.Label lblka01 
      BackStyle       =   0  '透明
      Caption         =   "①総合計"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   8160
      TabIndex        =   15
      Top             =   2370
      Width           =   1590
   End
   Begin VB.Label lblNendo 
      BackStyle       =   0  '透明
      Caption         =   "YYYY"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1830
      TabIndex        =   14
      Top             =   1290
      Width           =   1170
   End
   Begin VB.Line Line4 
      X1              =   7920
      X2              =   11050
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Line Line3 
      X1              =   11040
      X2              =   11040
      Y1              =   1710
      Y2              =   3710
   End
   Begin VB.Line Line2 
      X1              =   7920
      X2              =   11050
      Y1              =   3705
      Y2              =   3705
   End
   Begin VB.Line Line1 
      X1              =   7920
      X2              =   7920
      Y1              =   1710
      Y2              =   3710
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "合計得点の高い順に表示します。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8145
      TabIndex        =   13
      Top             =   2040
      Width           =   3030
   End
   Begin VB.Label lblka02 
      BackStyle       =   0  '透明
      Caption         =   "②面接Ⅰ＋小論文"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8160
      TabIndex        =   12
      Top             =   2655
      Width           =   1800
   End
   Begin VB.Label lblka04 
      BackStyle       =   0  '透明
      Caption         =   "④数学＋英語"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8160
      TabIndex        =   11
      Top             =   3255
      Width           =   1590
   End
   Begin VB.Label lblka03 
      BackStyle       =   0  '透明
      Caption         =   "③面接Ⅰ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   2955
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "表示順は、以下の科目の"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8145
      TabIndex        =   9
      Top             =   1830
      Width           =   2535
   End
   Begin VB.Label lblTit 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "処理年度"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   660
      TabIndex        =   1
      Tag             =   "1804"
      Top             =   1290
      Width           =   1080
   End
End
Attribute VB_Name = "frm2jiHoketusyaJuni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''Public g_obj_Conn               As ADODB.Connection   'connection object
''''Public g_void_OpenConnection    As Boolean


'*******************************************************************************
'* 3.10 補欠者順位(sub-systemより統合)                                         *
'*-----------------------------------------------------------------------------*
'* Form Load                                                                   *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Dim sConn        As String
    Dim sPWD         As String
    Dim sUser        As String
    Dim sDatabase    As String
    Dim sMachine     As String


    lblNendo.Caption = g_int_CurrentNendo & "年"
    txtNendo.Text = g_int_CurrentNendo ''''Year(Date) 隠しフィードになっている


    cmdUp.Enabled = False
    cmdDown.Enabled = False

''''    g_void_OpenConnection = False
''''
''''
''''    sPWD = GetSetting("Nyushi_zenki", "Settings", "DatabasePassword", "")
''''    sUser = GetSetting("Nyushi_zenki", "Settings", "DatabaseUser", "")
''''    sDatabase = GetSetting("Nyushi_zenki", "Settings", "DatabaseName", "")
''''    sMachine = GetSetting("Nyushi_zenki", "Settings", "MachineName", "")
''''
''''    If Trim(sUser) = "" Or Trim(sDatabase) = "" Or Trim(sMachine) = "" Then
''''        MsgBox "データベースの設定に誤りがあります。", vbInformation, "補欠者順位"
''''        Exit Sub
''''    End If
''''
''''    sConn = ";DSN=" & sMachine & ";UID=" & sUser & ";PWD=" & sPWD & ";Database=" & sDatabase
''''
''''    Set g_obj_Conn = New ADODB.Connection
''''    g_obj_Conn.CursorLocation = adUseClient
''''    g_obj_Conn.Open sConn ''''Database Open
''''
''''    If Err.Number <> 0 Then
''''        g_void_OpenConnection = False
''''    Else
''''        g_void_OpenConnection = True
''''    End If
    

    txtJuTotal.Text = lstKuriage.ListCount

    Exit Sub

ErrorHandler:
        MsgBox Err.Description, vbInformation, "補欠者順位"

End Sub

Private Sub Form_Unload(Cancel As Integer)

''''2022.01.25 del jhi 別のExeの場合有効だったのでこのProject内ではいらないのだ!
''''    If g_void_OpenConnection = True Then
''''        g_void_OpenConnection = False
''''        g_obj_Conn.Close
''''        Set g_obj_Conn = Nothing
''''    End If

    Call g_void_CloseChildForm
    Unload Me

End Sub

'*******************************************************************************
'* 補欠者 【表示】ボタン処理                                                   *
'*******************************************************************************
Private Sub cmdShow_Click()
   
    On Error GoTo ErrorHandler

'   Dim blnOpenDB              As Boolean
    Dim strNendo               As String
    Dim oRs                    As New ADODB.Recordset    'recordset object
    Dim strTemp                As String             'to form the display string in the list box
    Dim sSQL                   As String             'The SQL string

    
    'check nendo
    cmdUp.Enabled = False
    cmdDown.Enabled = False
        
''''    strNendo = txtNendo.Text
    strNendo = g_int_CurrentNendo

''''    If strNendo = "" Then
''''        MsgBox "年度を入力してください。", vbInformation, "補欠者順位"
''''        Exit Sub
''''    End If
''''
''''    strNendo = Trim(strNendo)
''''    If strNendo = "" Then
''''        MsgBox "年度を入力してください。", vbInformation, "補欠者順位"
''''        Exit Sub
''''    End If
''''
''''    If strNendo >= "2101" Or strNendo < "2010" Then
''''        MsgBox "年度入力に誤りがあります。(2010～2100年を指定してください)", vbInformation, "補欠者順位"
''''        Exit Sub
''''    End If
 
    
    
    Me.lstKuriage.Clear


    cmdShow.Enabled = False    '2021.11.17 add jhi

    'getData

'     sSQL = "SELECT iJukenNumber,substring(vExamineeName + '　　　　　　　　',1,10) as vExamineeName,iSex FROM tbSTEExamineeProfile WHERE" & _
'    " iNendo = " & strNendo & _
'    " AND iAbsentFlag = 0"
'    sSQL = sSQL & " AND iExamineeStatus = 3"

    sSQL = "Exec uspSTEGetExamineeOrder " & strNendo
    Set oRs = g_obj_Conn.Execute(sSQL)
    
    If oRs.EOF Then
        Set oRs = Nothing
        MsgBox "指定年度に該当するデータはありません。", vbInformation, "補欠者順位" '2021.11.17 add jhi
        cmdShow.Enabled = True
        Exit Sub
    End If


    Do While Not oRs.EOF

        strTemp = g_str_LPad(oRs.Fields("iJukenNumber").Value, Len(oRs.Fields("iJukenNumber").Value)) & _
            " - " & oRs.Fields("vExamineeName").Value

'        If oRs.Fields("iSex").Value = 0 Then
'            strTemp = strTemp & " - (*)"
'        End If
        
        lstKuriage.AddItem strTemp
        oRs.MoveNext

    Loop

    oRs.Close
    Set oRs = Nothing

    txtJuTotal.Text = lstKuriage.ListCount '2022.01.15 add jhi

    cmdShow.Enabled = True                 '2021.11.17 add jhi
  
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "補欠者順位"

End Sub

Private Sub cmdDown_Click()

    On Error GoTo ErrorHandler

    Dim l_int_Count As Integer


    If lstKuriage.ListCount < 1 Then
       Exit Sub
    End If

    If lstKuriage.SelCount < 0 Then
        Exit Sub
    End If

    'lstKuriage
    For l_int_Count = 0 To lstKuriage.ListCount - 1
        If lstKuriage.Selected(l_int_Count) Then
            lstKuriage.AddItem lstKuriage.List(l_int_Count), l_int_Count + 2
            lstKuriage.RemoveItem l_int_Count
            lstKuriage.Selected(l_int_Count + 1) = True
            Exit Sub
        End If
    Next
    
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "補欠者順位"

End Sub

Private Sub cmdExcel_Click()

    On Error GoTo ErrorHandler

    Dim fso                   As Object
    Dim objText               As Object
    Dim strFile               As String
    Dim blnOpenFile           As Boolean
    Dim l_str_JukenNo         As String
    Dim l_str_ExamineeName    As String
    Dim l_int_Count           As Integer
    Dim strLine               As String


    If lstKuriage.ListCount < 1 Then
        Exit Sub
    End If


    blnOpenFile = False

    'FSOオブジェクットを初期化
    strFile = App.Path & "\Report\補欠者順位" & Format(Now(), "YYYYMMDDHHmmSS") & ".csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.CreateTextFile(strFile, True, False)

    blnOpenFile = True

    l_str_JukenNo = ""
    l_str_ExamineeName = ""


    'ファイルを出力
    For l_int_Count = 0 To lstKuriage.ListCount - 1
        l_str_JukenNo = Left(lstKuriage.List(l_int_Count), 4)

        l_str_ExamineeName = Mid(lstKuriage.List(l_int_Count), 7)
        l_str_ExamineeName = Trim(l_str_ExamineeName)
        strLine = l_int_Count + 1 & "," & l_str_JukenNo & "," & l_str_ExamineeName
'       strLine = lstKuriage.List(l_int_Count)
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
    MsgBox Err.Description, vbInformation, "補欠者順位"

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorHandler

    Dim strNendo              As String
    Dim l_str_Sql             As String             ' The SQL string
    
    Dim l_int_TempJuken       As Integer            ' to store the juken number
    Dim l_str_JukenNo         As String             ' to store all the lstThisTimeSelected juken numbers as a string
    Dim l_str_ExamineeName    As String
    Dim blnTrans              As Boolean
    Dim l_int_Count           As Integer
    
    
    'check nendo
    cmdUp.Enabled = False
    cmdDown.Enabled = False
        
    strNendo = txtNendo.Text
    If strNendo = "" Then
        MsgBox "年度を入力してください。", vbInformation, "補欠者順位"
        Exit Sub
    End If

    strNendo = Trim(strNendo)
    If strNendo = "" Then
        MsgBox "年度を入力してください。", vbInformation, "補欠者順位"
        Exit Sub
    End If

    If strNendo >= "9999" And strNendo < "2010" Then
        MsgBox "年度入力欄に誤りがあります。", vbInformation, "補欠者順位"
        Exit Sub
    End If

    'getData

    
    blnTrans = False

'   l_str_Sql = "Exec uspSTESetExamineeOrder " & strNendo

    g_obj_Conn.BeginTrans

    l_str_Sql = "DELETE FROM tbSTEExamineeOrder WHERE  iNendo=" & strNendo
    g_obj_Conn.Execute (l_str_Sql)

    blnTrans = True

    For l_int_Count = 0 To lstKuriage.ListCount - 1
        l_int_TempJuken = Left(lstKuriage.List(l_int_Count), 4)
        l_str_JukenNo = l_int_TempJuken
        
        l_str_ExamineeName = Mid(lstKuriage.List(l_int_Count), 7)
        l_str_ExamineeName = Trim(l_str_ExamineeName)
        l_str_Sql = "INSERT INTO tbSTEExamineeOrder(iJukenNumber,iNendo,vExamineeName)"
        l_str_Sql = l_str_Sql & "VALUES(" & l_str_JukenNo & "," & strNendo
        l_str_Sql = l_str_Sql & ",'" & l_str_ExamineeName & "'"
        l_str_Sql = l_str_Sql & ")"
        g_obj_Conn.Execute (l_str_Sql)
    Next

    g_obj_Conn.CommitTrans
    blnTrans = False
    
    MsgBox "補欠者順番を更新しました。", vbInformation, "補欠者順位"

    Exit Sub

ErrorHandler:
    If blnTrans = True Then
        blnTrans = False
        g_obj_Conn.RollbackTrans
    End If

    MsgBox Err.Description, vbInformation, "補欠者順位"

End Sub


Public Function g_str_LPad(ByVal str As String, ByVal iLen As Integer) As String

    '-------------------------------------------------------------
    'Left pads a string with 0 up to iLen.
    '-------------------------------------------------------------
    Select Case iLen
    Case 1
        g_str_LPad = "000" & str
    Case 2
        g_str_LPad = "00" & str
    Case 3
        g_str_LPad = "0" & str
    Case 4
        g_str_LPad = str
    End Select

End Function

Private Sub cmdUp_Click()

    On Error GoTo ErrorHandler

    Dim l_int_Count As Integer



    If lstKuriage.ListCount < 1 Then
        Exit Sub
    End If

    If lstKuriage.SelCount < 0 Then
        Exit Sub
    End If


    'lstKuriage

    For l_int_Count = 0 To lstKuriage.ListCount - 1
        If lstKuriage.Selected(l_int_Count) Then
            lstKuriage.AddItem lstKuriage.List(l_int_Count), l_int_Count - 1
            lstKuriage.RemoveItem l_int_Count + 1
            lstKuriage.Selected(l_int_Count - 1) = True
            Exit Sub
        End If
    Next
    
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "補欠者順位"

End Sub

'*******************************************************************************
'* 別のExeの場合有効だったのでこのProject内ではいらないのだ!!                  *
'* このボタン処理は未使用にした。                                              *
'* 2022.01.25 del jhi                                                          *
'*******************************************************************************
Private Sub cmdClose_Click()

    On Error GoTo ErrorHandler

''''    If g_void_OpenConnection = True Then
''''        g_void_OpenConnection = False
''''        g_obj_Conn.Close
''''        Set g_obj_Conn = Nothing
''''    End If

''''End          '2021.11.17 del jhi

    Unload Me    '2021.11.17 add jhi
    Exit Sub     '2021.11.17 add jhi

ErrorHandler:
    MsgBox Err.Description, vbInformation, "cmdClose_Click:補欠者順位"

End Sub


Private Sub lstKuriage_Click()

    On Error GoTo ErrorHandler

    If lstKuriage.ListCount < 1 Then
        Exit Sub
    End If
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    If lstKuriage.Selected(0) Then
        cmdUp.Enabled = False
'        Exit Sub
    End If
    If lstKuriage.Selected(lstKuriage.ListCount - 1) Then
        cmdDown.Enabled = False
'        Exit Sub
    End If
    Exit Sub
ErrorHandler:
        MsgBox Err.Description, vbInformation, "補欠者順位"
End Sub


