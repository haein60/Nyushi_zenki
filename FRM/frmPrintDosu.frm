VERSION 5.00
Begin VB.Form frmPrintDosu 
   Caption         =   "frmPrintDosu : 度数分布図印刷"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13035
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmPrintDosu.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   13035
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel出力"
      Height          =   495
      Left            =   1815
      TabIndex        =   32
      Top             =   5790
      Width           =   1420
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   0
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   30
      Text            =   "0"
      Top             =   1680
      Width           =   525
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   1
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   28
      Text            =   "100"
      Top             =   2160
      Width           =   525
   End
   Begin VB.ComboBox cmbAdmission 
      Height          =   360
      Index           =   2
      Left            =   7680
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   27
      Top             =   4800
      Width           =   3405
   End
   Begin VB.ComboBox cmbSex 
      Height          =   360
      Index           =   2
      Left            =   11280
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   26
      Top             =   4800
      Width           =   885
   End
   Begin VB.ComboBox cmbAdmission 
      Height          =   360
      Index           =   1
      Left            =   7680
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   25
      Top             =   4200
      Width           =   3405
   End
   Begin VB.ComboBox cmbSex 
      Height          =   360
      Index           =   1
      Left            =   11280
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   24
      Top             =   4200
      Width           =   885
   End
   Begin VB.ComboBox cmbAdmission 
      Height          =   360
      Index           =   0
      Left            =   7680
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   23
      Top             =   3600
      Width           =   3405
   End
   Begin VB.ComboBox cmbSex 
      Height          =   360
      Index           =   0
      Left            =   11280
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   21
      Top             =   3600
      Width           =   885
   End
   Begin VB.ComboBox cmbSub 
      Height          =   360
      Index           =   0
      Left            =   2880
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   18
      Top             =   1200
      Width           =   3645
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "出力"
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   5790
      Width           =   1095
   End
   Begin VB.ComboBox cmbTarget 
      Height          =   360
      Index           =   2
      Left            =   4440
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   16
      Top             =   4800
      Width           =   3045
   End
   Begin VB.ComboBox cmbTarget 
      Height          =   360
      Index           =   1
      Left            =   4440
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   15
      Top             =   4200
      Width           =   3045
   End
   Begin VB.ComboBox cmbTarget 
      Height          =   360
      Index           =   0
      Left            =   4440
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   14
      Top             =   3600
      Width           =   3045
   End
   Begin VB.TextBox txtCnt 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   2
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "2"
      Top             =   4800
      Width           =   525
   End
   Begin VB.TextBox txtCnt 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   1
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "2"
      Top             =   4200
      Width           =   525
   End
   Begin VB.TextBox txtCnt 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   0
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "2"
      Top             =   3600
      Width           =   525
   End
   Begin VB.TextBox txtMark 
      Alignment       =   2  '中央揃え
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   6
      Text            =   "×"
      Top             =   4800
      Width           =   525
   End
   Begin VB.TextBox txtMark 
      Alignment       =   2  '中央揃え
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Text            =   "△"
      Top             =   4200
      Width           =   525
   End
   Begin VB.TextBox txtMark 
      Alignment       =   2  '中央揃え
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   2
      Text            =   "○"
      Top             =   3600
      Width           =   525
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   2
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "5"
      Top             =   2640
      Width           =   525
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "最低点"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   31
      Top             =   1740
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "最高点"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   29
      Top             =   2220
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "現浪区分"
      Height          =   240
      Index           =   9
      Left            =   8760
      TabIndex        =   22
      Top             =   3255
      Width           =   960
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "男女"
      Height          =   255
      Index           =   8
      Left            =   11400
      TabIndex        =   20
      Top             =   3255
      Width           =   495
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "出力対象科目"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   19
      Top             =   1260
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "対象"
      Height          =   255
      Index           =   6
      Left            =   5565
      TabIndex        =   13
      Top             =   3255
      Width           =   1215
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "人数"
      Height          =   255
      Index           =   5
      Left            =   3750
      TabIndex        =   12
      Top             =   3255
      Width           =   495
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "マーク"
      Height          =   255
      Index           =   4
      Left            =   2790
      TabIndex        =   8
      Top             =   3255
      Width           =   735
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "出力対象３"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   4860
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "出力対象２"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   4260
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "出力対象１"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   3660
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "刻み幅"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2700
      Width           =   2535
   End
End
Attribute VB_Name = "frmPrintDosu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private prvuPrintTarget_()    As puPrintCategoryType
Private prvuPrintSub_()       As puPrintCategoryType
Private prvuPrintAdmission_() As puPrintCategoryType


'*******************************************************************************
'* 度数分布図のデータを格納するStructure文                                     *
'*-----------------------------------------------------------------------------*
'* 2021.12.07 add jhi                                                          *
'*******************************************************************************
Private Type DosuDataType

    fStartScore As Integer
    fEndScore   As Integer
    vScore      As String

    lCnt1       As Integer
    lCnt2       As Integer
    lCnt3       As Integer

    fMax1       As Integer
    fMax2       As Integer
    fMax3       As Integer

    fMin1       As Integer
    fMin2       As Integer
    fMin3       As Integer

    fAvg1       As Double
    fAvg2       As Double
    fAvg3       As Double

    fSd1       As Double
    fSd2       As Double
    fSd3       As Double

    fSum1       As Integer
    fSum2       As Integer
    fSum3       As Integer

    lRuiCnt1    As Integer
    lRuiCnt2    As Integer
    lRuiCnt3    As Integer

End Type

Dim Dosu() As DosuDataType

'*******************************************************************************
'* 度数分布図をExcelグラフ機能を使用し、出力する                               *
'*-----------------------------------------------------------------------------*
'* 2021.12.07 add jhi                                                          *
'*******************************************************************************
Private Sub cmdExcel_Click()

    Dim oXl         As Object 'Excel

    Dim FileNM      As String 'ファイル名
    Dim BookNM      As String 'ブック名
    Dim SheetNM     As String 'シート名

    Dim myFilename  As String
    Dim myfile      As String

    Dim strTemp     As String

    Dim i           As Integer
    Dim step_no     As Integer
    Dim rinf        As Long


    '--------------------------------------------------------------------------
    '度数分布データを配列に設定する関数
    '--------------------------------------------------------------------------
    Call PrintProc_JHI


    FileNM = App.Path & "\Template_DosuBun.xls"   'ファイル名

    Set oXl = CreateObject("Excel.Application")   'excel起動

    oXl.Workbooks.Open (FileNM)                   'ブックを開く
    BookNM = oXl.ActiveWorkbook.Name              'ブック名を取得
    SheetNM = oXl.ActiveSheet.Name                'シート名を取得



    'Titleを設定
    oXl.Workbooks(BookNM).Worksheets(SheetNM).Cells(1, 1).Value = cmbSub(0).Text & " 度数分布図" ''''Title設定

    For i = 0 To UBound(Dosu) - 1
        strTemp = Dosu(i).lCnt1 & "(" & Format(Dosu(i).lRuiCnt1, "###0") & ")"
        oXl.Workbooks(BookNM).Worksheets(SheetNM).Cells(i + 65, 3).Value = strTemp         'Y軸
        oXl.Workbooks(BookNM).Worksheets(SheetNM).Cells(i + 65, 4).Value = Dosu(i).vScore  '階級(X軸をセット)
        oXl.Workbooks(BookNM).Worksheets(SheetNM).Cells(i + 65, 5).Value = Dosu(i).lCnt1   '度数(グラフ)をセット
    Next i

    '***************************************************************************
    '* 最低点/最高点/平均点/標準偏差 をセットする                              *
    '***************************************************************************
    oXl.Workbooks(BookNM).Worksheets(SheetNM).Cells(45, 3).Value = Format(Dosu(0).fMin1, "##0.0") '最低点
    oXl.Workbooks(BookNM).Worksheets(SheetNM).Cells(46, 3).Value = Format(Dosu(0).fMax1, "##0.0") '最高点
    oXl.Workbooks(BookNM).Worksheets(SheetNM).Cells(47, 3).Value = Format(Dosu(0).fAvg1, "##0.0") '平均点
    oXl.Workbooks(BookNM).Worksheets(SheetNM).Cells(48, 3).Value = Format(Dosu(0).fSd1, "#0.0")   '標準偏差


    '出力ファイルの指定
    myFilename = "C:\Output.xls"

    'ファイルが存在するか調べる
    myfile = Dir$(myFilename)

    oXl.DisplayAlerts = False

    If Len(myfile) = 0 Then
''''    MsgBox "そのファイルは存在しません。saveします。"
        oXl.Workbooks(BookNM).SaveAs ("C:\Output.xls")    '保存
        oXl.Visible = True                                '表示

    Else

'        'ファイルがあればファイルの名前を同じ名前で変更します。
'        Name myFilename As myFilename
'
'        'ファイルが使用中であればエラーが発生します
'        If Err.Number Then
'            MsgBox "ファイルは使用中です。"
'
'           'エラーが発生した場合は Err オブジェクトをクリアします。
'            Err.Clear
'        Else
'            MsgBox "ファイルは使われていません。"
'            oXl.Workbooks(BookNM).SaveAs ("C:\Output.xls")    '保存
'            oXl.Visible = True                                '表示
'        End If



''''    On Error Resume Next
        On Error GoTo ErrProc

        rinf = MsgBox("この場所に 「Output.xls」 という名前のファイルが既にあります。置き換えますか？", vbInformation + vbYesNo + vbDefaultButton2)
        If rinf = vbYes Then

            step_no = 1

            oXl.Workbooks(BookNM).SaveAs ("C:\Output.xls")    '保存

''''        On Error GoTo 0                                   'goto 0の意味: エラールーチンを無効にします。実行時エラーが発生しても、エラー処理しません。
''''        On Error GoTo ErrProc
            oXl.Visible = True                                '表示
        Else
            step_no = 2
            oXl.Workbooks(BookNM).Close                       '閉じる
            On Error GoTo ErrProc
            oXl.Quit                                          '終了
            Set oXl = Nothing                                 '解放
        End If

    End If


    oXl.DisplayAlerts = True
    Exit Sub


ErrProc:

    If step_no = 1 Then
        MsgBox ("Output.xlsファイルが開いているので閉じてから再度操作を行ってください。")
        oXl.Workbooks(BookNM).Close                       '閉じる
        On Error GoTo 0
        oXl.Quit                                          '終了
        Set oXl = Nothing                                 '解放

    ElseIf (step_no = 2) Then
        ''''処理なし

    End If

End Sub

'*******************************************************************************
'* Excel 度数分布図のデータ格納処理                                            *
'*-----------------------------------------------------------------------------*
'* 2021.12.06 add JHI                                                          *
'*******************************************************************************
Private Sub PrintProc_JHI()

    On Error GoTo ErrProc

    Dim oRs As ADODB.Recordset
    Dim sSQL As String

    Dim icnt  As Integer

    Dim lStartScore As Long
    Dim lEndScore   As Long
    Dim dScoreScale As Double

    Dim lCnt1 As Long
    Dim lCnt2 As Long
    Dim lCnt3 As Long

    Dim sMark1 As String
    Dim sMark2 As String
    Dim sMark3 As String


    Dim iPosCnt As Integer


    lStartScore = txtPara(0).Text
    lEndScore = txtPara(1).Text
    dScoreScale = txtPara(2).Text

    lCnt1 = txtCnt(0).Text
    lCnt2 = txtCnt(1).Text
    lCnt3 = txtCnt(2).Text

    sMark1 = txtMark(0).Text
    sMark2 = txtMark(1).Text
    sMark3 = txtMark(2).Text

    If cmbTarget(1).Text = "" Then
        iPosCnt = 1
    ElseIf cmbTarget(2).Text = "" Then
        iPosCnt = 2
    Else
        iPosCnt = 3
    End If

    sSQL = "exec uspSTEGetDosuData '" & Trim(str(g_int_CurrentNendo))
    sSQL = sSQL & "','" & Trim(str(cmbSub(0).ItemData(cmbSub(0).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(lStartScore))
    sSQL = sSQL & "','" & Trim(str(lEndScore))
    sSQL = sSQL & "','" & Trim(str(dScoreScale))
    sSQL = sSQL & "','" & Trim(str(cmbTarget(0).ItemData(cmbTarget(0).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbAdmission(0).ItemData(cmbAdmission(0).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbSex(0).ItemData(cmbSex(0).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbTarget(1).ItemData(cmbTarget(1).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbAdmission(1).ItemData(cmbAdmission(1).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbSex(1).ItemData(cmbSex(1).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbTarget(2).ItemData(cmbTarget(2).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbAdmission(2).ItemData(cmbAdmission(2).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbSex(2).ItemData(cmbSex(2).ListIndex))) & "'"

    g_obj_Conn.CommandTimeout = 360
    Set oRs = g_obj_Conn.Execute(sSQL)

    icnt = 0
    ReDim Dosu(icnt) As DosuDataType

    Do Until oRs.EOF

        Dosu(icnt).fStartScore = oRs.Fields("fStartScore")
        Dosu(icnt).fEndScore = oRs.Fields("fEndScore")
        Dosu(icnt).vScore = Format(oRs.Fields("fStartScore"), "##0") & "～" & Format(oRs.Fields("fEndScore"), "##0")

        Dosu(icnt).lCnt1 = oRs.Fields("lCnt1")
        Dosu(icnt).lRuiCnt1 = oRs.Fields("lRuiCnt1")


        '***************************************************************************
        '* 最低点/最高点/平均点/標準偏差 をセットする                              *
        '***************************************************************************
        Dosu(icnt).fMin1 = oRs.Fields("fMin1")
        Dosu(icnt).fMax1 = oRs.Fields("fMax1")
        Dosu(icnt).fAvg1 = oRs.Fields("fAvg1")
        Dosu(icnt).fSd1 = oRs.Fields("fSd1")


''''    Debug.Print "icnt=" & icnt

        icnt = icnt + 1
        ReDim Preserve Dosu(icnt) As DosuDataType

        oRs.MoveNext

    Loop

    oRs.Close
    Set oRs = Nothing


''''MsgBox "icnt----->" & iCnt
''''MsgBox "LBound(Dosu)=" & CStr(LBound(Dosu)) & " UBound(Dosu)=" & CStr(UBound(Dosu))

    Exit Sub

ErrProc:

End Sub

'*******************************************************************************
'* 度数分布図をExcelグラフ機能を使用し、出力する                               *
'*-----------------------------------------------------------------------------*
'* 2021.12.07 add jhi                                                          *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler

    LoadResStrings Me
    Call g_void_SetFontProperties(Me)     ' set the font properties

    Me.Caption = "frmPrintDosu : 度数分布図印刷"    ''''LoadResString(2700)



    Call makeTarget

    Call f_void_PopulateCmbTarget
    Call f_void_PopulateCmbAdmission
    Call f_void_PopulateCmbSex
    Call f_void_PopulateCmbSub
    Call f_void_GetDefPosMark

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim i As Integer

    fMainForm.mnuTools.Enabled = False  ' disable tools menu

    For i = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next i

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub makeTarget()

    On Error GoTo ErrProc

    Dim sSQL     As String
    Dim oRs      As ADODB.Recordset
    Dim iLoopCnt As Integer

    Erase prvuPrintTarget_

    sSQL = "select "
    sSQL = sSQL & "  iExamineeCategoryID "
    sSQL = sSQL & ", vDispName "
    sSQL = sSQL & ", vCondition "
    sSQL = sSQL & " from tbSTEExamineeCategory "
    sSQL = sSQL & " where iDispOrder <> -1 "
    sSQL = sSQL & " order by iDispOrder "

    Set oRs = g_obj_Conn.Execute(sSQL)

    If oRs.EOF Then
        Set oRs = Nothing
        MsgBox "tbSTEExamineeCategoryマスタテーブルのデータ取得に失敗しました。"
        Exit Sub
    End If

    iLoopCnt = 0

    Do Until oRs.EOF

        ReDim Preserve prvuPrintTarget_(iLoopCnt)

        prvuPrintTarget_(iLoopCnt).iID = oRs.Fields(0)
        prvuPrintTarget_(iLoopCnt).sDispName = oRs.Fields(1)

        iLoopCnt = iLoopCnt + 1

        oRs.MoveNext

    Loop

    oRs.Close
    Set oRs = Nothing

    Exit Sub

ErrProc:

End Sub

Private Sub PrintHeader(poRs As ADODB.Recordset)

    Dim lStrHeight As Long
    Dim lStrWidth  As Long
    Dim x1 As Long, x2 As Long, x3 As Long, y1 As Long, y2 As Long
    Dim xBase As Long, yBase As Long
    Dim xDiff As Long, yDiff As Long
    Dim xMaxPos As Long, yMaxPos As Long
    Dim iLoopCnt As Integer, iLoopCnt2 As Integer

    Dim lScoreScale As Long
    Dim lCnt1 As Long
    Dim lCnt2 As Long
    Dim lCnt3 As Long

    Dim sTitle As String
    Dim sMark1 As String, sMark2 As String, sMark3 As String
    Dim iColCnt As Integer, iMaxCnt As Integer

    Dim svx1 As Long
    Dim svy1 As Long
    Dim svy2 As Long
    Dim svy3 As Long

    Printer.PaperSize = vbPRPSA3    '用紙をA3に
    Printer.Orientation = 2         '印刷向きを横に

    Printer.Font = "ＭＳゴシック"
    Printer.FontSize = 10

    lStrHeight = Printer.TextHeight("○")
    lStrWidth = Printer.TextWidth("○")

    xBase = 3000
    yBase = 2000
    xDiff = 0
    yDiff = lStrHeight / 2
    xMaxPos = 100 '本番は100
    yMaxPos = 40  '本番は40

    x1 = xBase
    y1 = yBase
    x2 = xBase + xDiff * xMaxPos + lStrWidth * xMaxPos
    y2 = yBase + yDiff * yMaxPos + lStrHeight * yMaxPos

    lScoreScale = txtPara(0).Text
    lCnt1 = txtCnt(0).Text
    lCnt2 = txtCnt(1).Text
    lCnt3 = txtCnt(2).Text
    sMark1 = txtMark(0).Text
    sMark2 = txtMark(1).Text
    sMark3 = txtMark(2).Text

    Printer.Line (x1, y1)-(x1, y2)
'    Printer.Line (x1, y1)-(x2, y1)

'ヘッダの出力
    sTitle = "度数分布図"
    Printer.FontSize = 16
    x1 = Printer.ScaleWidth / 2

    For iLoopCnt = 1 To Len(sTitle)
        x1 = x1 - Printer.TextHeight(Mid(sTitle, iLoopCnt, 1)) / 2
    Next

    Printer.CurrentX = x1
    Printer.CurrentY = 0
    Printer.Print sTitle
    Printer.FontSize = 10

    x1 = xBase - xDiff * 8 - lStrWidth * 3
    Printer.CurrentX = x1
    Printer.Print "科目：" & cmbSub(0).Text
'    Printer.Print "１ポイントあたりの人数：" & txtPara(0).Text
    Printer.CurrentX = x1
    Printer.Print sMark1 & "：" & cmbTarget(0).Text & "／" & cmbAdmission(0).Text & "／" & cmbSex(0).Text & "／" & txtCnt(0).Text & "名"
    Printer.CurrentX = x1

    If cmbTarget(1).Text <> "" Then
        Printer.Print sMark2 & "：" & cmbTarget(1).Text & "／" & cmbAdmission(1).Text & "／" & cmbSex(1).Text & "／" & txtCnt(1).Text & "名"
    Else
        Printer.Print ""
    End If

    Printer.CurrentX = x1

    If cmbTarget(2).Text <> "" Then
        Printer.Print sMark3 & "：" & cmbTarget(2).Text & "／" & cmbAdmission(2).Text & "／" & cmbSex(2).Text & "／" & txtCnt(2).Text & "名"
    Else
        Printer.Print ""
    End If

    y1 = yBase + yDiff * (yMaxPos + 1) + lStrHeight * (yMaxPos + 1)
    Printer.CurrentY = y1
    Printer.Print
    sTitle = "最低点"
    x1 = xBase - lStrWidth

    For iLoopCnt = 1 To Len(sTitle)
        x1 = x1 - Printer.TextHeight(Mid(sTitle, iLoopCnt, 1))
    Next

    Printer.CurrentX = x1
    Printer.Print sTitle

    sTitle = "最高点"
    x1 = xBase - lStrWidth

    For iLoopCnt = 1 To Len(sTitle)
        x1 = x1 - Printer.TextHeight(Mid(sTitle, iLoopCnt, 1))
    Next

    Printer.CurrentX = x1
    Printer.Print sTitle

    sTitle = "平均点"
    x1 = xBase - lStrWidth

    For iLoopCnt = 1 To Len(sTitle)
        x1 = x1 - Printer.TextHeight(Mid(sTitle, iLoopCnt, 1))
    Next

    Printer.CurrentX = x1
    Printer.Print sTitle

    sTitle = "標準偏差"
    x1 = xBase - lStrWidth

    For iLoopCnt = 1 To Len(sTitle)
        x1 = x1 - Printer.TextHeight(Mid(sTitle, iLoopCnt, 1))
    Next

    Printer.CurrentX = x1
    Printer.Print sTitle
    
    x1 = xBase
    x1 = xBase + lStrWidth * 2.5
    Printer.CurrentX = x1
    Printer.CurrentY = y1
    Printer.Print sMark1
    x1 = xBase + lStrWidth * 2
    Printer.CurrentX = x1
    Printer.Print Format(poRs.Fields("fmin1"), "##0.0")
    Printer.CurrentX = x1
    Printer.Print Format(poRs.Fields("fmax1"), "##0.0")
    Printer.CurrentX = x1
    Printer.Print Format(poRs.Fields("favg1"), "##0.0")
    Printer.CurrentX = x1
    Printer.Print Format(poRs.Fields("fsd1"), "#0.0")

    If cmbTarget(1).Text <> "" Then
        x1 = xBase + lStrWidth * 6.5
        Printer.CurrentX = x1
        Printer.CurrentY = y1
        Printer.Print sMark2
        x1 = xBase + lStrWidth * 6
        Printer.CurrentX = x1
        Printer.Print Format(poRs.Fields("fmin2"), "##0.0")
        Printer.CurrentX = x1
        Printer.Print Format(poRs.Fields("fmax2"), "##0.0")
        Printer.CurrentX = x1
        Printer.Print Format(poRs.Fields("favg2"), "##0.0")
        Printer.CurrentX = x1
        Printer.Print Format(poRs.Fields("fsd2"), "#0.0")
    End If

    If cmbTarget(2).Text <> "" Then
        x1 = xBase + lStrWidth * 10.5
        Printer.CurrentX = x1
        Printer.CurrentY = y1
        Printer.Print sMark3
        x1 = xBase + lStrWidth * 10
        Printer.CurrentX = x1
        Printer.Print Format(poRs.Fields("fmin3"), "##0.0")
        Printer.CurrentX = x1
        Printer.Print Format(poRs.Fields("fmax3"), "##0.0")
        Printer.CurrentX = x1
        Printer.Print Format(poRs.Fields("favg3"), "##0.0")
        Printer.CurrentX = x1
        Printer.Print Format(poRs.Fields("fsd3"), "#0.0")
    End If

End Sub

'*******************************************************************************
'* 実際 度数分布図を作成する関数                                               *
'*******************************************************************************
Private Sub PrintProc()

    On Error GoTo ErrProc

    Dim lStrHeight As Long
    Dim lStrWidth  As Long

    Dim x1 As Long
    Dim x2 As Long
    Dim x3 As Long

    Dim y1 As Long
    Dim y2 As Long

    Dim xBase As Long
    Dim yBase As Long

    Dim xDiff As Long
    Dim yDiff As Long

    Dim xMaxPos As Long
    Dim yMaxPos As Long

    Dim iLoopCnt  As Integer
    Dim iLoopCnt2 As Integer

    Dim dScoreScale As Double, lStartScore As Long, lEndScore As Long
    Dim lCnt1 As Long, lCnt2 As Long, lCnt3 As Long
    Dim sTitle As String
    Dim sMark1 As String, sMark2 As String, sMark3 As String
    Dim iColCnt As Integer, iMaxCnt As Integer

    Dim oRs As ADODB.Recordset
    Dim sSQL As String

    Dim iPosCnt As Integer


    lStartScore = txtPara(0).Text
    lEndScore = txtPara(1).Text
    dScoreScale = txtPara(2).Text

    lCnt1 = txtCnt(0).Text
    lCnt2 = txtCnt(1).Text
    lCnt3 = txtCnt(2).Text

    sMark1 = txtMark(0).Text
    sMark2 = txtMark(1).Text
    sMark3 = txtMark(2).Text

    If cmbTarget(1).Text = "" Then
        iPosCnt = 1
    ElseIf cmbTarget(2).Text = "" Then
        iPosCnt = 2
    Else
        iPosCnt = 3
    End If

    sSQL = "exec uspSTEGetDosuData '" & Trim(str(g_int_CurrentNendo))
    sSQL = sSQL & "','" & Trim(str(cmbSub(0).ItemData(cmbSub(0).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(lStartScore))
    sSQL = sSQL & "','" & Trim(str(lEndScore))
    sSQL = sSQL & "','" & Trim(str(dScoreScale))
    sSQL = sSQL & "','" & Trim(str(cmbTarget(0).ItemData(cmbTarget(0).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbAdmission(0).ItemData(cmbAdmission(0).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbSex(0).ItemData(cmbSex(0).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbTarget(1).ItemData(cmbTarget(1).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbAdmission(1).ItemData(cmbAdmission(1).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbSex(1).ItemData(cmbSex(1).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbTarget(2).ItemData(cmbTarget(2).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbAdmission(2).ItemData(cmbAdmission(2).ListIndex)))
    sSQL = sSQL & "','" & Trim(str(cmbSex(2).ItemData(cmbSex(2).ListIndex))) & "'"

    g_obj_Conn.CommandTimeout = 360
    Set oRs = g_obj_Conn.Execute(sSQL)

    Do Until oRs.EOF

        Call PrintHeader(oRs)

        lStrHeight = Printer.TextHeight("○")
        lStrWidth = Printer.TextWidth("○")

        xBase = 3000
        yBase = 2000
        xDiff = 0
        yDiff = lStrHeight / 2

        xMaxPos = 100 '本番は100
        yMaxPos = 40  '本番は40

        For iLoopCnt = 1 To yMaxPos

            If oRs.EOF Then Exit For

            'Ｙ軸見出し出力
            '人数（累計人数）1
            If iPosCnt = 3 Then
                x1 = xBase - lStrWidth * 14
                y1 = yBase + yDiff * (iLoopCnt - 1) + lStrHeight * (iLoopCnt)
            ElseIf iPosCnt = 2 Then
                x1 = xBase - lStrWidth * 14
                y1 = yBase + yDiff * (iLoopCnt - 1) + lStrHeight * (iLoopCnt)
            Else
                x1 = xBase - lStrWidth * 9
                y1 = yBase + yDiff * (iLoopCnt - 1) + lStrHeight * (iLoopCnt)
            End If
            
            Printer.CurrentY = y1
            Printer.CurrentX = x1
            Printer.Print Right(Space(2) & Format(oRs.Fields("lCnt1"), "##0"), 3) & "(" & Right(Space(3) & Format(oRs.Fields("lRuiCnt1"), "###0"), 4) & ")"

            '人数（累計人数）2
            If iPosCnt = 3 Then
                x1 = xBase - lStrWidth * 9
                y1 = yBase + yDiff * (iLoopCnt - 1) + lStrHeight * (iLoopCnt) - lStrHeight * 0.4
            Else
                x1 = xBase - lStrWidth * 9
                y1 = yBase + yDiff * (iLoopCnt - 1) + lStrHeight * (iLoopCnt)
            End If

            If iPosCnt > 1 Then
                Printer.CurrentY = y1
                Printer.CurrentX = x1
                Printer.Print Right(Space(2) & Format(oRs.Fields("lCnt2"), "##0"), 3) & "(" & Right(Space(3) & Format(oRs.Fields("lRuiCnt2"), "###0"), 4) & ")"
            End If

            '人数（累計人数）3
            If iPosCnt > 2 Then
                y1 = y1 + lStrHeight * 0.8
                Printer.CurrentY = y1
                Printer.CurrentX = x1
                Printer.Print Right(Space(2) & Format(oRs.Fields("lCnt3"), "##0"), 3) & "(" & Right(Space(3) & Format(oRs.Fields("lRuiCnt3"), "###0"), 4) & ")"
            End If

            '点数範囲(95～100,90～95...)
            x1 = xBase - lStrWidth * 4
            y1 = yBase + yDiff * (iLoopCnt - 1) + lStrHeight * (iLoopCnt)
            Printer.CurrentY = y1
            Printer.CurrentX = x1

            If Int(dScoreScale) <> dScoreScale Then
                Printer.Print Format(oRs.Fields(0), "0.0") & "～" & Format(oRs.Fields(1), "0.0")
            Else
                Printer.Print Right(Space(2) & Format(oRs.Fields(0), "##0"), 3) & "～" & Right(Space(2) & Format(oRs.Fields(1), "##0"), 3)
            End If

            'データの出力
            x1 = xBase
'            y1 = yBase + yDiff * (iLoopCnt - 1) + lStrHeight * (iLoopCnt)

            Printer.CurrentY = y1
            Printer.CurrentX = x1
            Printer.Print String(RoundUp(CDbl(oRs.Fields(2)) / lCnt1), sMark1)           ''''lCnt1

            If iPosCnt > 1 Then
                Printer.CurrentY = y1
                Printer.CurrentX = x1
                Printer.Print String(RoundUp(CDbl(oRs.Fields(3)) / lCnt2), sMark2)       ''''lCnt2
                If iPosCnt > 2 Then
                    Printer.CurrentY = y1
                    Printer.CurrentX = x1
                    Printer.Print String(RoundUp(CDbl(oRs.Fields(4)) / lCnt3), sMark3)   ''''lCnt3
                End If
            End If
            oRs.MoveNext
        Next

        Printer.EndDoc

    Loop

    oRs.Close
    Set oRs = Nothing

'最低点
'最高点
'平均点
'標準偏差の印刷
'標準偏差=SQRT(分散値/総得点）
'分散値=SUM((点数-平均点)^2)
'総得点=SUM(点数)

    Exit Sub

ErrProc:

On Error Resume Next

    Printer.KillDoc

On Error GoTo 0
On Error Resume Next

    oRs.Close

On Error GoTo 0
On Error Resume Next

    Set oRs = Nothing

On Error GoTo 0

End Sub

Private Sub cmbSub_Click(Index As Integer)

    Dim iLoopCnt As Integer
    Dim iDimIndex As Integer

    iDimIndex = -1

    For iLoopCnt = LBound(prvuPrintSub_) To UBound(prvuPrintSub_)
        If cmbSub(Index).ItemData(cmbSub(Index).ListIndex) = prvuPrintSub_(iLoopCnt).iID Then
            iDimIndex = iLoopCnt
        End If
    Next

    If iDimIndex >= 0 Then
        txtPara(0).Text = Trim(str(prvuPrintSub_(iDimIndex).dDefStartScore))
        txtPara(1).Text = Trim(str(prvuPrintSub_(iDimIndex).dDefEndScore))
    End If

End Sub

Private Sub cmdExec_Click()

    cmdExec.Enabled = False

    DoEvents

    Call PrintProc

    cmdExec.Enabled = True

End Sub


Private Sub f_void_GetDefPosMark()

    On Error GoTo ErrorHandler

    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    Dim l_int_Counter As Integer
    Dim iID As Long
    Dim vSubName As String


    ' select all subjects that come under the selected exam type
    l_str_Sql = "SELECT vMark FROM tbSTEMarks "
    l_str_Sql = l_str_Sql & " Where iMarkType  = 3 ORDER BY iMarkId "
    
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    ' add the subjects to combo box
    l_int_Counter = 0

    Do While Not l_obj_Rst.EOF
        l_int_Counter = l_int_Counter + 1
        txtMark(l_int_Counter).Text = l_obj_Rst.Fields(0).Value
        l_obj_Rst.MoveNext
    Loop
    
    ' release the object variables
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing

    Exit Sub

ErrorHandler:

End Sub

Private Sub f_void_PopulateCmbSub()

    On Error GoTo ErrorHandler

    Dim sSQL     As String                 ' SQL string
    Dim oRs      As New ADODB.Recordset    ' recordset object
    Dim iID      As Long
    Dim vSubName As String
    Dim iLoopCnt As Integer


    Erase prvuPrintSub_

'    cmbSub(0).AddItem ""
'    cmbSub(0).ItemData(cmbSub(0).NewIndex) = -1

    iLoopCnt = 0

    ' select all subjects that come under the selected exam type
    sSQL = "SELECT iTotalCategoryID,vDispName "
    sSQL = sSQL & ", isnull( fDefStartScore , 0 ) "
    sSQL = sSQL & ", isnull( fDefEndScore , 100 ) "
    sSQL = sSQL & ", isnull( fDefScaleScore , 5 ) "
    sSQL = sSQL & " FROM tbSTETotalCategory "
    sSQL = sSQL & " WHERE iDispOrder <> -1 "
    sSQL = sSQL & " ORDER BY iDispOrder "

    Set oRs = g_obj_Conn.Execute(sSQL)

    ' add the subjects to combo box
    Do While Not oRs.EOF

        ReDim Preserve prvuPrintSub_(iLoopCnt)

        prvuPrintSub_(iLoopCnt).iID = oRs.Fields(0)
        prvuPrintSub_(iLoopCnt).sDispName = oRs.Fields(1)
        prvuPrintSub_(iLoopCnt).dDefStartScore = oRs.Fields(2)
        prvuPrintSub_(iLoopCnt).dDefEndScore = oRs.Fields(3)
        prvuPrintSub_(iLoopCnt).dDefScaleScore = oRs.Fields(4)

        cmbSub(0).AddItem prvuPrintSub_(iLoopCnt).sDispName
        cmbSub(0).ItemData(cmbSub(0).NewIndex) = prvuPrintSub_(iLoopCnt).iID
        oRs.MoveNext

        iLoopCnt = iLoopCnt + 1

    Loop
    
    ' release the object variables
    oRs.Close
    Set oRs = Nothing

    If cmbSub.Count > 0 Then
        cmbSub(0).ListIndex = 0
    End If

    Exit Sub

ErrorHandler:

End Sub

Private Sub f_void_PopulateCmbTarget()

    On Error GoTo ErrorHandler

    Dim l_int_Counter As Integer
    Dim iID           As Long
    Dim vSubName      As String
    

'    cmbTarget(0).AddItem ""
'    cmbTarget(0).ItemData(cmbTarget(0).NewIndex) = -1
    cmbTarget(1).AddItem ""
    cmbTarget(1).ItemData(cmbTarget(1).NewIndex) = -1
    cmbTarget(2).AddItem ""
    cmbTarget(2).ItemData(cmbTarget(2).NewIndex) = -1

    ' add the subjects to combo box
    For l_int_Counter = LBound(prvuPrintTarget_) To UBound(prvuPrintTarget_)
        iID = prvuPrintTarget_(l_int_Counter).iID
        vSubName = prvuPrintTarget_(l_int_Counter).sDispName
        cmbTarget(0).AddItem vSubName
        cmbTarget(0).ItemData(cmbTarget(0).NewIndex) = iID
        cmbTarget(1).AddItem vSubName
        cmbTarget(1).ItemData(cmbTarget(1).NewIndex) = iID
        cmbTarget(2).AddItem vSubName
        cmbTarget(2).ItemData(cmbTarget(2).NewIndex) = iID
    Next

    If cmbTarget(0).ListCount > 0 Then
        cmbTarget(0).ListIndex = 0
    End If

    cmbTarget(1).ListIndex = 0
    cmbTarget(2).ListIndex = 0

    Exit Sub

ErrorHandler:

End Sub

Private Sub f_void_PopulateCmbAdmission()

    On Error GoTo ErrorHandler

    Dim sSQL     As String                 ' SQL string
    Dim oRs      As New ADODB.Recordset    ' recordset object
    Dim iLoopCnt As Integer
    Dim iID      As Long
    Dim vName    As String


    Erase prvuPrintAdmission_

    iLoopCnt = 0

    sSQL = "select "
    sSQL = sSQL & "  iAdmissionCategoryID "
    sSQL = sSQL & ", vDispName "
    sSQL = sSQL & " from tbSTEAdmissionCategory "
    sSQL = sSQL & " where iDispOrder > 0 "
    sSQL = sSQL & " order by iDispOrder "

    Set oRs = g_obj_Conn.Execute(sSQL)

    ' add the subjects to combo box
    Do Until oRs.EOF
        ReDim Preserve prvuPrintAdmission_(iLoopCnt)
        prvuPrintAdmission_(iLoopCnt).iID = oRs.Fields(0)
        prvuPrintAdmission_(iLoopCnt).sDispName = oRs.Fields(1)
        cmbAdmission(0).AddItem prvuPrintAdmission_(iLoopCnt).sDispName
        cmbAdmission(0).ItemData(cmbAdmission(0).NewIndex) = prvuPrintAdmission_(iLoopCnt).iID
        cmbAdmission(1).AddItem prvuPrintAdmission_(iLoopCnt).sDispName
        cmbAdmission(1).ItemData(cmbAdmission(1).NewIndex) = prvuPrintAdmission_(iLoopCnt).iID
        cmbAdmission(2).AddItem prvuPrintAdmission_(iLoopCnt).sDispName
        cmbAdmission(2).ItemData(cmbAdmission(2).NewIndex) = prvuPrintAdmission_(iLoopCnt).iID
        oRs.MoveNext
        iLoopCnt = iLoopCnt + 1
    Loop

    ' release the object variables
    oRs.Close
    Set oRs = Nothing

    If cmbAdmission(0).ListCount > 0 Then
        cmbAdmission(0).ListIndex = 0
    End If
    If cmbAdmission(1).ListCount > 0 Then
        cmbAdmission(1).ListIndex = 0
    End If
    If cmbAdmission(2).ListCount > 0 Then
        cmbAdmission(2).ListIndex = 0
    End If

    Exit Sub

ErrorHandler:

End Sub

Private Sub f_void_PopulateCmbSex()

    cmbSex(0).AddItem "全員"
    cmbSex(0).ItemData(cmbSex(0).NewIndex) = -1
    cmbSex(1).AddItem "全員"
    cmbSex(1).ItemData(cmbSex(1).NewIndex) = -1
    cmbSex(2).AddItem "全員"
    cmbSex(2).ItemData(cmbSex(2).NewIndex) = -1

    cmbSex(0).AddItem "男"
    cmbSex(0).ItemData(cmbSex(0).NewIndex) = 0
    cmbSex(1).AddItem "男"
    cmbSex(1).ItemData(cmbSex(1).NewIndex) = 0
    cmbSex(2).AddItem "男"
    cmbSex(2).ItemData(cmbSex(2).NewIndex) = 0

    cmbSex(0).AddItem "女"
    cmbSex(0).ItemData(cmbSex(0).NewIndex) = 1
    cmbSex(1).AddItem "女"
    cmbSex(1).ItemData(cmbSex(1).NewIndex) = 1
    cmbSex(2).AddItem "女"
    cmbSex(2).ItemData(cmbSex(2).NewIndex) = 1

    cmbSex(0).ListIndex = 0
    cmbSex(1).ListIndex = 0
    cmbSex(2).ListIndex = 0

End Sub
