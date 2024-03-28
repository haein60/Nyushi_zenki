VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmImportSyoronbun 
   Caption         =   $"frmImportSyoronbun.frx":0000
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmImportSyoronbun.frx":0032
   ScaleHeight     =   10755
   ScaleWidth      =   13980
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdDataSet 
      Caption         =   "CSVデータをDBに反映"
      Default         =   -1  'True
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
      Left            =   5490
      TabIndex        =   7
      Top             =   5865
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6700
      Left            =   555
      TabIndex        =   6
      Top             =   2055
      Width           =   4525
      _ExtentX        =   7990
      _ExtentY        =   11827
      _Version        =   393216
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdDataDisp 
      Caption         =   "CSVデータ表示"
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
      Left            =   5505
      TabIndex        =   2
      Top             =   3615
      Width           =   2775
   End
   Begin VB.CommandButton cmdFileSentaku 
      Caption         =   "..."
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
      Left            =   10815
      TabIndex        =   1
      Top             =   1455
      Width           =   450
   End
   Begin VB.TextBox txtCSVPathFile 
      BackColor       =   &H00C0C0C0&
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
      Height          =   390
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1470
      Width           =   7455
   End
   Begin MSComDlg.CommonDialog cdlCSVPath 
      Left            =   12630
      Top             =   435
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select CSV File"
      Filter          =   "CSV Files (*.csv)|*.csv|その他テキストファイル(*)|*.*|"
   End
   Begin VB.Label lblGuid1 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "ヘッダなしのcsvファイルをご指定ください。"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6375
      TabIndex        =   8
      Top             =   1200
      Width           =   4890
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "CSVデータファイルを選択"
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
      Height          =   390
      Left            =   450
      TabIndex        =   5
      Top             =   1515
      Width           =   2880
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
      Height          =   345
      Left            =   540
      TabIndex        =   0
      Top             =   9000
      Width           =   9360
   End
   Begin VB.Label lblCSVPathFile 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "素点入力(小論文)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   465
      TabIndex        =   4
      Top             =   1230
      Width           =   2190
   End
End
Attribute VB_Name = "frmImportSyoronbun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************************
'Form Name   : 素点入力(小論文)_import(frmImportSyoronbun)
'Author      : jhi
'Created On  : 2021.12.21
'Update  On  : 2022.01.04
'Description :
'Reference   :
'*******************************************************************************

Private m_SecondExam_Type    As Long      '面接か小論文かflag
Private CurrentRowNo         As Integer   'active cellの行を取得

Dim FN_CSV                   As String    'Importするcsvファイル名
Dim giNendo                  As Long      '処理年度
Dim gupKensu                 As Long      'updateした件数

Dim gDestFile                As String    '小論文csvファイルをサーバに入れるファイル名


Private Type SyoData_Type
    No         As Integer
    iNendo     As String
    juno       As Integer
    fScore     As Single
    idbsetflg  As Integer
End Type
Private SyoData()    As SyoData_Type   '小論文 data


'*******************************************************************************
'* Form_Load 関数 frmImportSYoronbun : 素点入力(小論文)_Import                 *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler
    Dim i    As Long
    Dim rinf As Long


    Me.Caption = "frmImportSyoronbun : 素点入力(小論文)_import"

    lblMsg.Caption = ""


    '---------------------------------------------------------------------------
    ' MSFlexGrid1の初期処理をする
    '---------------------------------------------------------------------------
    Call MSFlexGrid1_Syokisyori

    If Trim(txtCSVPathFile.Text) = "" Then
        cmdDataDisp.Enabled = False
        cmdDataSet.Enabled = False
    End If


    giNendo = g_int_CurrentNendo
    ''''MsgBox (g_int_CurrentNendo) 'global variable in form load


    rinf = DB_Data_Disp_Syo


    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

'*******************************************************************************
'* Form_Activate 関数                                                          *
'*******************************************************************************
Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim i As Integer

    fMainForm.mnuTools.Enabled = False  ' disable tools menu

    For i = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next i

    Exit Sub


ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)

End Sub

Private Sub cmdFileSentaku_Click()

    On Error GoTo ErrorHandler


    lblMsg.Caption = ""

    Err.Clear
    cdlCSVPath.ShowOpen


    ' check for cancel error
    If Err.Number = 0 Then
''''    txtCSVPathFile.Text = Left(cdlCSVPath.FileName, InStrRev(cdlCSVPath.FileName, "\"))
        txtCSVPathFile.Text = cdlCSVPath.FileName
    End If

    'csv file名をセット
    FN_CSV = txtCSVPathFile.Text

    If Trim(FN_CSV) <> "" Then
        cmdDataDisp.Enabled = True
        cmdDataSet.Enabled = True

        '小論文の成績import ファイルをサーバ側にcopyする
        Call fCopy(FN_CSV, "W:\score_syoronbun30_" & giNendo & ".csv")
    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

'*******************************************************************************
'* MSFlexGridにcsvデータを表示する 関数                                        *
'*******************************************************************************
Private Sub cmdDataDisp_Click()

    On Error GoTo ErrorHandler

    'データ読込表示処理
    Dim iNendo      As Integer    '年度
    Dim juken_no    As Integer    '受験番号
    Dim soten       As Single     '素点

    Dim cnt         As Integer    'データのカウント
    Dim intFileNo   As Integer    'ファイルNo

    Dim sTmp        As String
    Dim step_no     As Integer

    Dim rinf        As Long


step_no = 1

    cnt = 0

    lblMsg.Caption = ""

    'importするcsvファイル指定有無チェック
    sTmp = txtCSVPathFile.Text

    If (sTmp = "") Then
        MsgBox "Importするcsvファイルを指定してください。"
        Exit Sub
    End If

    MSFlexGrid1.Clear
    MSFlexGrid1.Refresh

    '---------------------------------------------------------------------------
    ' MSFlexGrid1 初期設定
    '---------------------------------------------------------------------------
    Call MSFlexGrid1_Syokisyori


    ''''csvファイルの列数をcheckする関数
    rinf = ReadCsvFile(FN_CSV)
    If rinf <> 0 Then
        step_no = 3
        GoTo ErrorHandler
    End If


    With MSFlexGrid1

step_no = 2

        .Visible = False        '一旦非表示に(読込が早くなる)
        intFileNo = FreeFile

        'シーケンシャル入力モードでSeiseki.txtをオープン
        'ファイルのPATHは別途設定して下さい。
        Open FN_CSV For Input As #intFileNo

        Do Until EOF(intFileNo)   'EOF(intFileNo)が True になるまで実行

step_no = 3

            Input #intFileNo, iNendo, juken_no, soten

            '読込んだデータをセルに代入
            .Rows = cnt + 2
            .Row = cnt + 1
            .RowHeight(.Row) = 320

step_no = 4
            .Col = 0
            .Text = Format$(cnt + 1, "###0")     'no

step_no = 5
            .Col = 1
            .Text = Format$(iNendo, "###0")      '年度

step_no = 6
            .Col = 2
            .Text = Format$(juken_no, "000#")    '受験番号

step_no = 7
            .Col = 3
            .Text = Format$(soten, "#0.0")       '素点

step_no = 8

            cnt = cnt + 1
        Loop

step_no = 9
        Close #intFileNo


step_no = 10
        'カレントセルをホームポジションに
        .Row = 1
        .Col = 1
        .TopRow = 1
        .Visible = True         '再表示
''''    .SetFocus               'errorが発生するのでやめた

    End With

    '--------------------------------------------------------------------------
    '小論文のwork Table作成関数を呼出す
    '--------------------------------------------------------------------------
    Call Set_work_table(FN_CSV)

    Exit Sub


ErrorHandler:
    If step_no = 1 Then
        MsgBox "importするcsvファイルの取得に失敗しました。(ファイル名=" & sTmp & ")"

    ElseIf step_no = 2 Then
        MsgBox "importするcsvファイルのOpenで失敗しました。"

    ElseIf step_no = 3 Then
        MsgBox "importするcsvファイルの列数で誤りがあります。csvファイルをご確認ください。(No=" & rinf & ")"

    ElseIf step_no = 4 Then 'no check
        MsgBox "importするcsvファイルからNo作成に失敗しました。"

    ElseIf step_no = 5 Then '受験番号
        MsgBox "受験番号設定に誤りがありました。(No=" & cnt & ")"

    ElseIf step_no = 6 Then
        MsgBox "素点設定に誤りがありました。(No=" & cnt & ")"

    ElseIf step_no = 7 Then
        MsgBox "importするcsvファイルからcnt作成に失敗しました。(cnt=" & cnt & ")"

    ElseIf step_no = 8 Then
        MsgBox "importするcsvファイルのCloseで失敗しました"

    ElseIf step_no = 9 Then
        MsgBox "カレントセルをホームポジションに処理でエラーが発生しました。"

    Else
        MsgBox "importするcsvファイルからエラーが発生しました。(step_no=" & step_no & ")"
    End If


    Call MSFlexGrid1_Syokisyori


End Sub

Private Sub MSFlexGrid1_Syokisyori()

    Dim i As Integer


    'MSFlexGrid の初期設定
    With MSFlexGrid1

        .Rows = 21                  '行の総数（固定行含む）
        .cols = 4                   '列の総数（固定列含む）
        .FixedRows = 1              '固定行の数 Rowsより１以上少ない事
        .FixedCols = 1              '固定列の数 Colsより１以上少ない事
        .Row = 0
        .ColWidth(0) = 900          'Noの列幅
        .ColWidth(1) = 900          '年度
        .ColWidth(2) = 1200         '受験番号
        .ColWidth(3) = 1200         '素点


        .RowHeight(0) = 320         '行の高さ

        .Col = 0
        .Text = "No"
        .CellAlignment = flexAlignCenterCenter '該当セルを　中寄／中寄　表示

        .Col = 1
        .Text = "年度"
        .CellAlignment = flexAlignCenterCenter

        .Col = 2
        .Text = "受験番号"
        .CellAlignment = flexAlignCenterCenter

        .Col = 3
        .Text = "素点"
        .CellAlignment = flexAlignCenterCenter


        .Col = 0
        For i = 1 To .Rows - 1
            .RowHeight(i) = 320     '行の高さ
            .Row = i
            .Text = i               '行番号を表示
        Next i

        .Col = 1
        .Row = 1

        'カレントセルを反転表示（強調表示すればカレントセルが解りやすい）
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways

    End With


End Sub

Private Function ReadCsvFile(fn As String) As Long

    Dim strFilename As String
    Dim intFileNo   As Integer
    Dim blnOpenFlg  As Boolean
    Dim vntBuf      As Variant
    Dim strBuf      As String
    Dim lngCnt      As Long
    Dim lngDataCnt  As Long
    
    '初期値設定
    blnOpenFlg = False
    
    'ファイル名設定
    strFilename = fn
    intFileNo = FreeFile()
    
    'ファイルオープン
    Open strFilename For Input As #intFileNo
    'ファイルオープンしたらフラグOn
    blnOpenFlg = True
    
    lngCnt = 1
    Do While Not EOF(intFileNo)
        'データを一行単位で読込
        Line Input #intFileNo, strBuf
        'Splitをつかって、取り出した文字列を分解
        vntBuf = Split(strBuf, ",")
        '要素数を取得
        lngDataCnt = UBound(vntBuf) + 1
        
''''    MsgBox CStr(lngCnt) & "行目の項目数は" & CStr(lngDataCnt) & "個です。"
        If (lngDataCnt <> 3) Then
            GoTo GO_ERR
        End If
        
        lngCnt = lngCnt + 1
    Loop
    

    'ファイルを閉じる
    If blnOpenFlg = True Then
        Close #intFileNo
        blnOpenFlg = False
    End If

    ReadCsvFile = 0
    Exit Function


GO_ERR:
    
    'ファイルを閉じる
    If blnOpenFlg = True Then
        Close #intFileNo
        blnOpenFlg = False
    End If

    ReadCsvFile = lngCnt

End Function

'*******************************************************************************
' csv fileをTableに入れる処理
'*******************************************************************************
Private Sub Set_work_table(csvFName As String)

    On Error GoTo ErrorHandler
   
    Dim oRs  As New ADODB.Recordset
    Dim sSQL As String

    Dim strDestFile As String


    gDestFile = "C:\share\score_syoronbun30_" & giNendo & ".csv"


    'tmp_csvscore30 Table clear
    sSQL = ""
    sSQL = sSQL & "delete tmp_csvscore30" & vbCrLf
    sSQL = sSQL & "where iNendo=" & giNendo

    Set oRs = g_obj_Conn.Execute(sSQL)

    'release the object variable
    Set oRs = Nothing


    '---------------------------------------------------------------------------
    'csv file内容をtmp_csvscore30 tableに入れるsql文処理
    '---------------------------------------------------------------------------
    sSQL = ""
    sSQL = sSQL & "BULK INSERT tmp_csvscore30" & vbCrLf
''''sSQL = sSQL & "FROM '" & csvFName & "'" & vbCrLf     ''''2022.02.09 del jhi
    sSQL = sSQL & "FROM '" & gDestFile & "'" & vbCrLf    ''''2022.02.09 add jhi
    sSQL = sSQL & "WITH" & vbCrLf
    sSQL = sSQL & "(" & vbCrLf
    sSQL = sSQL & "   FIELDTERMINATOR = ','," & vbCrLf
    sSQL = sSQL & "   ROWTERMINATOR = '\n'" & vbCrLf
    sSQL = sSQL & ");"

    Set oRs = g_obj_Conn.Execute(sSQL)
    
    'release the object variable
    Set oRs = Nothing



    'tbSTEcsvscore30 Table clear
    sSQL = ""
    sSQL = sSQL & "delete tbSTEcsvscore30" & vbCrLf
    sSQL = sSQL & "where iNendo=" & giNendo

    Set oRs = g_obj_Conn.Execute(sSQL)

    'release the object variable
    Set oRs = Nothing


    sSQL = ""
    sSQL = sSQL & "insert into tbSTEcsvscore30" & vbCrLf
    sSQL = sSQL & "select" & vbCrLf
    sSQL = sSQL & "    a.*" & vbCrLf
    sSQL = sSQL & "   ,b.iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "   ,0" & vbCrLf                      'idbsetflg 初期は0をセットする
    sSQL = sSQL & "   ,GETDATE()" & vbCrLf              'dtCreate
    sSQL = sSQL & "   ,GETDATE()" & vbCrLf              'dtUpdate

    sSQL = sSQL & "from" & vbCrLf
    sSQL = sSQL & "    tmp_csvscore30       a" & vbCrLf
    sSQL = sSQL & "   ,tbSTEExamineeProfile b" & vbCrLf
    sSQL = sSQL & "where" & vbCrLf
    sSQL = sSQL & "        a.iJukenno = b.iJukenNumber" & vbCrLf
    sSQL = sSQL & "    and a.iNendo   = b.iNendo" & vbCrLf
    sSQL = sSQL & "order by" & vbCrLf
    sSQL = sSQL & "    a.iNendo" & vbCrLf
    sSQL = sSQL & "   ,a.iJukenno" & vbCrLf

    Set oRs = g_obj_Conn.Execute(sSQL)
    
    'release the object variable
    Set oRs = Nothing

    Exit Sub


ErrorHandler:

    MsgBox "tbSTEcsvscore30 work Table(小論文)作成時エラーが発生しました。"


End Sub

'*******************************************************************************
'* 小論文 素点を DB tbSTEScoreProfile tableに反映する                          *
'*******************************************************************************
Private Sub cmdDataSet_Click()

    On Error GoTo ErrorHandler
    Dim rinf As Long


    lblMsg.Caption = ""

    rinf = myMsgBox("Importしました小論文、素点をDBに反映します。よろしいですか？", gTit)
    If rinf = vbCancel Then
        Exit Sub
    End If


    Call Set_Hon_table

    lblMsg.Caption = gupKensu & "件の小論文、素点をDBに正常に反映しました。"

    Exit Sub


ErrorHandler:

    MsgBox "小論文、素点をDBに反映処理関数でエラーが発生しました。"



End Sub

'-------------------------------------------------------------------------------
' csv fileの指定 fRawScoreをtbSTEScoreProfile TableのfRawScoreを反映する
'-------------------------------------------------------------------------------
Private Sub Set_Hon_table()

    On Error GoTo ErrorHandler

    Dim oRs    As New ADODB.Recordset
    Dim sSQL   As String


    g_obj_Conn.BeginTrans

#If 0 Then
    '***************************************************************************
    '* 以下 update文のみだけではなくinsertもありますので以下は未使用になる     *
    '* 2022.01.26 del jhi                                                      *
    '***************************************************************************

    sSQL = ""
    sSQL = sSQL & "update ta" & vbCrLf
    sSQL = sSQL & "   set ta.fRawScore = so.fRawScore" & vbCrLf
    sSQL = sSQL & "from" & vbCrLf
    sSQL = sSQL & "    tbSTEScoreProfile ta" & vbCrLf
    sSQL = sSQL & "    inner join" & vbCrLf
    sSQL = sSQL & "    tbSTEcsvscore30 so on" & vbCrLf
    sSQL = sSQL & "            ta.iExamineeProfileId = so.iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "        and substring(convert(nvarchar,ta.dtCreate,111),1,4)=" & giNendo & vbCrLf
    sSQL = sSQL & "        and ta.iSubjectProfileId = 30" & vbCrLf
    sSQL = sSQL & "        and ta.iAbsentFlag       = 0" & vbCrLf
    sSQL = sSQL & "        and so.iNendo            = " & giNendo & vbCrLf


'-------------------------------------------------------------------------------
'2021.12.17 add jhi
'-------------------------------------------------------------------------------
Update ta
   Set ta.fRawScore = so.fRawScore
From
    tbSTEScoreProfile ta
    Inner Join
    tbSTEcsvscore30 so on
            ta.iExamineeProfileId = so.iExamineeProfileId
        and substring(convert(nvarchar,ta.dtUpdate,111),1,4)='2020'
        and ta.iSubjectProfileId = 30
        and ta.iAbsentFlag       = 0
        and so.iNendo            = 2020
--where
--    so.iNendo =2020
-------------------------------------------------------------------------------
#End If

    '***************************************************************************
    '* insertもありますのでSQL文を修正                                         *
    '* 2022.01.26 add jhi                                                      *
    '***************************************************************************
    sSQL = ""
    sSQL = sSQL & "MERGE INTO" & vbCrLf
    sSQL = sSQL & "    tbSTEScoreProfile AS sp" & vbCrLf
    sSQL = sSQL & "USING (" & vbCrLf
    sSQL = sSQL & "    select" & vbCrLf
    sSQL = sSQL & "        (ROW_NUMBER() OVER(ORDER BY iJukenno) + sp.iScoreProfileId )as iScoreProfileId" & vbCrLf
    sSQL = sSQL & "       ,30        as iSubjectProfileId" & vbCrLf
    sSQL = sSQL & "       ,iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "       ,fRawScore" & vbCrLf
    sSQL = sSQL & "       ,0.00      as fChoseiScore" & vbCrLf
    sSQL = sSQL & "       ,0         as iAbsentFlag" & vbCrLf
    sSQL = sSQL & "       ,GETDATE() as dtCreate" & vbCrLf
    sSQL = sSQL & "       ,GETDATE() as dtUpdate" & vbCrLf
    sSQL = sSQL & "    from" & vbCrLf
    sSQL = sSQL & "        tbSTEcsvscore30" & vbCrLf
    sSQL = sSQL & "       ,(select" & vbCrLf
    sSQL = sSQL & "             MAX(iScoreProfileId) as iScoreProfileId" & vbCrLf
    sSQL = sSQL & "         from" & vbCrLf
    sSQL = sSQL & "             tbSTEScoreProfile" & vbCrLf
    sSQL = sSQL & "         where" & vbCrLf
    sSQL = sSQL & "             convert(varchar(4),dtcreate,112) >= '" & g_int_CurrentNendo & "'" & vbCrLf
    sSQL = sSQL & "        ) sp" & vbCrLf
    sSQL = sSQL & "    where" & vbCrLf
    sSQL = sSQL & "        iNendo=" & g_int_CurrentNendo & vbCrLf
    sSQL = sSQL & "    ) cs" & vbCrLf
    sSQL = sSQL & "    on sp.iExamineeProfileId = cs.iExamineeProfileId and sp.iSubjectProfileID=30" & vbCrLf
    sSQL = sSQL & "WHEN MATCHED THEN" & vbCrLf
    sSQL = sSQL & "    UPDATE SET" & vbCrLf
    sSQL = sSQL & "       fRawScore=cs.fRawScore" & vbCrLf
    sSQL = sSQL & "      ,dtUpdate =GETDATE()" & vbCrLf
    sSQL = sSQL & "WHEN NOT MATCHED THEN" & vbCrLf
    sSQL = sSQL & "    INSERT(iScoreProfileId,iSubjectProfileId,iExamineeProfileId,fRawScore,fChoseiScore,iAbsentFlag,dtCreate,dtUpdate)" & vbCrLf
    sSQL = sSQL & "    VALUES" & vbCrLf
    sSQL = sSQL & "    (" & vbCrLf
    sSQL = sSQL & "        cs.iScoreProfileId" & vbCrLf
    sSQL = sSQL & "       ,cs.iSubjectProfileId" & vbCrLf
    sSQL = sSQL & "       ,cs.iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "       ,cs.fRawScore" & vbCrLf
    sSQL = sSQL & "       ,cs.fChoseiScore" & vbCrLf
    sSQL = sSQL & "       ,cs.iAbsentFlag" & vbCrLf
    sSQL = sSQL & "       ,cs.dtCreate" & vbCrLf
    sSQL = sSQL & "       ,cs.dtUpdate" & vbCrLf
    sSQL = sSQL & "    )" & vbCrLf
    sSQL = sSQL & ";"

'*******************************************************************************
#If 0 Then

Merge Into
    tbSTEScoreProfile As sp
USING (
    select
        (ROW_NUMBER() OVER(ORDER BY iJukenno)  + sp.iScoreProfileId )as iScoreProfileId
       ,30        as iSubjectProfileId
--     ,iNendo
--     ,iJukenno
       ,iExamineeProfileId
       ,fRawScore
       ,0.00      as fChoseiScore
       ,0         as iAbsentFlag
       ,GETDATE() as dtCreate
       ,GETDATE() as dtUpdate
    From
        tbSTEcsvscore30
       ,(select
             MAX(iScoreProfileId) As iScoreProfileId
         From
             tbSTEScoreProfile
         Where
             convert(varchar(4),dtcreate,112) >= '2022'
        ) sp
    Where
        iNendo = 2022
    ) cs
    on sp.iExamineeProfileId = cs.iExamineeProfileId and sp.iSubjectProfileID=30
WHEN MATCHED THEN
    Update
    SET
       fRawScore = cs.fRawScore
      ,dtUpdate =GETDATE()
WHEN NOT MATCHED THEN
    INSERT(iScoreProfileId,iSubjectProfileId,iExamineeProfileId,fRawScore,fChoseiScore,iAbsentFlag,dtCreate,dtUpdate)
    Values
    (
        cs.iScoreProfileId
       ,cs.iSubjectProfileId
       ,cs.iExamineeProfileId
       ,cs.fRawScore
       ,cs.fChoseiScore
       ,cs.iAbsentFlag
       ,cs.dtCreate
       ,cs.dtUpdate
    )
;
#End If
'*******************************************************************************

    Set oRs = g_obj_Conn.Execute(sSQL)


    Set oRs = Nothing


    sSQL = ""
    sSQL = sSQL & "select @@ROWCOUNT;"
    oRs.Open sSQL, g_obj_Conn

    gupKensu = 0
    gupKensu = oRs.Fields(0)

    oRs.Close
    Set oRs = Nothing


    lblMsg.Caption = gupKensu & "件の小論文素点をDBに反映しました。"


    sSQL = ""
    sSQL = sSQL & "update tbSTEcsvscore30" & vbCrLf
    sSQL = sSQL & "   set idbsetflg = 1" & vbCrLf
    sSQL = sSQL & "where" & vbCrLf
    sSQL = sSQL & "    iNendo = " & giNendo & vbCrLf

    Set oRs = g_obj_Conn.Execute(sSQL)

    Set oRs = Nothing

    g_obj_Conn.CommitTrans

    Exit Sub


ErrorHandler:
    g_obj_Conn.RollbackTrans
    MsgBox "Set_Hon_table()関数処理時エラーが発生しました。"

End Sub

''''2021.12.12 add jhi
Public Sub gsSetSecondType(piSType As Long)

    If piSType = 0 Then
        m_SecondExam_Type = 0 '面接
    Else
        m_SecondExam_Type = 1 '小論文
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call g_void_CloseChildForm

End Sub

'*******************************************************************************
'* MSFlexGridにDBに小論文importデータを反映していれば表示する 関数             *
'*******************************************************************************
Private Function DB_Data_Disp_Syo() As Long

    On Error GoTo ErrorHandler

    Dim oRs         As New ADODB.Recordset
    Dim sSQL        As String

    Dim step_no     As Integer
    Dim icnt        As Integer    'データのカウント
    Dim i           As Integer    'loopカウント

    Dim rinf        As Long



step_no = 1

    DB_Data_Disp_Syo = 0


    MSFlexGrid1.Clear
    MSFlexGrid1.Refresh

    '---------------------------------------------------------------------------
    ' MSFlexGrid1 初期設定
    '---------------------------------------------------------------------------
    Call MSFlexGrid1_Syokisyori


    sSQL = ""
    sSQL = sSQL & "select" & vbCrLf
    sSQL = sSQL & "    iNendo" & vbCrLf
    sSQL = sSQL & "   ,iJukenno" & vbCrLf
    sSQL = sSQL & "   ,fRawScore" & vbCrLf
    sSQL = sSQL & "   ,idbsetflg" & vbCrLf
    sSQL = sSQL & "from" & vbCrLf
    sSQL = sSQL & "    tbSTEcsvscore30" & vbCrLf
    sSQL = sSQL & "where" & vbCrLf
    sSQL = sSQL & "    iNendo=" & giNendo & vbCrLf

    Set oRs = g_obj_Conn.Execute(sSQL)
    
    If oRs.EOF Then
        DB_Data_Disp_Syo = 0
        oRs.Close
        Set oRs = Nothing
        Exit Function
    Else
        oRs.MoveFirst
    End If
 

    icnt = 0
    Do While Not oRs.EOF

        ReDim Preserve SyoData(icnt)

        SyoData(icnt).No = icnt + 1
        SyoData(icnt).iNendo = oRs.Fields(0)       'iNendo
        SyoData(icnt).juno = oRs.Fields(1)         'iJukenno
        SyoData(icnt).fScore = oRs.Fields(2)       'fRawScore
        SyoData(icnt).idbsetflg = oRs.Fields(3)    'idbsetflg

        If SyoData(icnt).idbsetflg = 0 Then
            GoTo DBSET_NASI
        End If

        icnt = icnt + 1
        oRs.MoveNext
    Loop

    oRs.Close
    Set oRs = Nothing


    MSFlexGrid1.Visible = False        '一旦非表示に(読込が早くなる)

    For i = 0 To UBound(SyoData)

        MSFlexGrid1.Rows = i + 2
        MSFlexGrid1.Row = i + 1
        MSFlexGrid1.RowHeight(i + 1) = 320

        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = Format$(i + 1, "###0")                   'no

        MSFlexGrid1.Col = 1
        MSFlexGrid1.Text = Format$(SyoData(i).iNendo, "###0")       '年度

        MSFlexGrid1.Col = 2
        MSFlexGrid1.Text = Format$(SyoData(i).juno, "000#")         '受験番号

        MSFlexGrid1.Col = 3
        MSFlexGrid1.Text = Format$(SyoData(i).fScore, "#0.0")       '素点

    Next i


    'カレントセルをホームポジションに
    MSFlexGrid1.Row = 1
    MSFlexGrid1.Col = 1
    MSFlexGrid1.TopRow = 1
    MSFlexGrid1.Visible = True         '再表示

    DB_Data_Disp_Syo = i
    Exit Function


DBSET_NASI:

    oRs.Close
    Set oRs = Nothing
 
    DB_Data_Disp_Syo = 0
    Exit Function


ErrorHandler:
    MsgBox "DB_Data_Disp_Syo()関数でエラーが発生しました。"


End Function



'-------------------------------------------------------------------------------
' csv read sample1
'-------------------------------------------------------------------------------
Public Sub sample1()

    Dim intNo As Integer        'ファイルNo
    Dim lngCount As Long        'データ数
    Dim strItem1() As String    '項目1用配列
    Dim strItem2() As String    '項目2用配列
    Dim strItem3() As String    '項目3用配列

    'csvファイルオープン
    intNo = FileSystem.FreeFile()
    Open "C:\tmp\sample.csv" For Input As #intNo

    'csvファイルの読み込み
    lngCount = 0
    Do Until EOF(intNo)

        '配列のリサイズ
        ReDim Preserve strItem1(lngCount) As String
        ReDim Preserve strItem2(lngCount) As String
        ReDim Preserve strItem3(lngCount) As String

        'データを各配列に読み込み
        Input #intNo, strItem1(lngCount), strItem2(lngCount), strItem3(lngCount)

        lngCount = lngCount + 1
    Loop

    'csvファイルクローズ
    Close #intNo

    '読み込んだ値を確認
    Dim i As Long
    For i = 0 To UBound(strItem1)
        Debug.Print strItem1(i), strItem2(i), strItem3(i)
    Next i

End Sub

'-------------------------------------------------------------------------------
' csv read sample2
'-------------------------------------------------------------------------------
Public Sub sample2()

    Dim intNo As Integer        'ファイルNo
    Dim lngCount As Long        'データ数
    Dim tmpData() As String     '一時保存用配列
    Dim strItem1() As String    '項目1用配列
    Dim strItem2() As String    '項目2用配列
    Dim strItem3() As String    '項目3用配列
    Dim FSO As New FileSystemObject
    Dim ts As TextStream

    'ファイルオープン
    Set ts = FSO.OpenTextFile("C:\tmp\sample.csv")

    'ファイル読み込み
    lngCount = 0
    With ts
        Do Until .AtEndOfStream

            '配列のリサイズ
            ReDim Preserve strItem1(lngCount) As String
            ReDim Preserve strItem2(lngCount) As String
            ReDim Preserve strItem3(lngCount) As String

            '一行データを読み込み、カンマ区切りで配列に変換
            tmpData = Split(.ReadLine, ",")

            '読み込んだデータの前後の「"」を削除して配列に格納
            strItem1(lngCount) = Mid(tmpData(0), 2, Len(tmpData(0)) - 2)
            strItem2(lngCount) = Mid(tmpData(1), 2, Len(tmpData(1)) - 2)
            strItem3(lngCount) = Mid(tmpData(2), 2, Len(tmpData(2)) - 2)

            lngCount = lngCount + 1
        Loop

    End With

    'ファイルクローズ
    ts.Close
    Set ts = Nothing

    '読み込んだ値を確認
    Dim i As Long
    For i = 0 To UBound(strItem1)
        Debug.Print strItem1(i), strItem2(i), strItem3(i)
    Next i

End Sub

