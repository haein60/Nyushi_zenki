Attribute VB_Name = "basExcelReport"
Option Explicit

' ShellExecute
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

' エラー番号定義
Public Const USERERR = vbObjectError + 1000

' プログラムカウンタ変数定義
Private g_nPc As Long
' 変数配列
Private g_asParam() As String
' Dialog
Public g_bInput As Boolean

' Excel オブジェクト変数定義
Private g_objExcel As Excel.Application
Private g_objWorkbook As Workbook
Private g_objSheetCmd As Worksheet
Private g_objSheetData As Worksheet

'add,2007/11/09,S-------
' Page 用変数
Private g_avPageId() As Variant
Private g_nPageNo As Long
Private g_nPageSetupPc As Long
'add,2007/11/09,E-------

' ADO Connection オブジェクト変数定義
Private g_objCnn As ADODB.Connection

Public Const gsUserPwd = "ComPwd"
'//////////////////////////////////////////////////////////////////////////////////////////////////
'// ExcelReportMain
Public Sub ExcelReportMain(ByVal sTemplateFile As String, ByRef sOutputFile As String, Optional ByVal sPassword As String = "")
    ' Init
    g_nPc = 0
    Erase g_asParam
    ReDim g_asParam(1, 0)
    
    ' エラーハンドル登録
    On Error GoTo ERROR_HANDLE

    ' Excelオブジェクト作成
    Set g_objExcel = CreateObject("Excel.Application")
    g_objExcel.Visible = False
    g_objExcel.DisplayAlerts = False
    ' テンプレートファイル読込
    g_objExcel.Workbooks.Open FileName:=sTemplateFile, ReadOnly:=True, Password:=sPassword
    ' ワークブックオブジェクト取得
    Set g_objWorkbook = g_objExcel.ActiveWorkbook

    ' Commandシートオブジェクト取得
    On Error Resume Next
    Set g_objSheetCmd = g_objWorkbook.Sheets("Command")
    If Err.Number Then
        Err.Number = USERERR
        Err.Description = "シート ""Command"" が有りません。"
        GoTo ERROR_HANDLE
    End If
    If Err.Number Then GoTo ERROR_HANDLE
    On Error GoTo ERROR_HANDLE
    
    ' コマンド処理
    Call CommandProcess
    
    ' Command Sheetを除去
    g_objSheetCmd.Delete
    ' 同名のファイルを除去
Dim sWkFileName As String
Dim sWk As String
Dim iRetryCnt As Integer
    iRetryCnt = 0
    sWkFileName = sOutputFile & ".xls"
    sWk = Dir(sWkFileName)
    If sWk <> "" Then
        On Error GoTo KillErr
        Kill sOutputFile & ".xls"
        GoTo KillOK
KillErr:
'オープンしていてだめ
        iRetryCnt = iRetryCnt + 1
        sWkFileName = sOutputFile & Trim(Str(iRetryCnt)) & ".xls"
        sWk = Dir(sWkFileName)
        If sWk <> "" Then
            On Error GoTo KillErr
            Kill sWkFileName
        End If
KillOK:

    End If

    On Error GoTo ERROR_HANDLE
    ' 保存
'    g_objWorkbook.SaveAs sWkFileName, , sPassword
    g_objWorkbook.SaveAs sWkFileName, xlExcel7, ""
    sOutputFile = sWkFileName
    ' 終了
'    g_objWorkbook.Close True, sWkFileName, False
    g_objWorkbook.Close

    ' Excel オブジェクトを除去
    Set g_objSheetData = Nothing
    Set g_objSheetCmd = Nothing
    Set g_objWorkbook = Nothing
    g_objExcel.Quit
    Set g_objExcel = Nothing
    Exit Sub

ERROR_HANDLE:
    ' AdoClose
    AdoClose
    ' Excelオブジェクト除去
    Set g_objSheetData = Nothing
    Set g_objSheetCmd = Nothing
    Set g_objWorkbook = Nothing
    If Not (g_objExcel Is Nothing) Then
        g_objExcel.Quit
    End If
    Set g_objExcel = Nothing
    ' エラー情報に行番号を付加
    If g_nPc Then
        Err.Description = Err.Description & "  Line=" & CStr(g_nPc)
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// AdoConnect
Public Sub AdoConnect(ByVal sDsn As String, ByVal sUserId As String, ByVal sPassword As String)
    Dim sParam As String
    
    ' ADO Connection オブジェクトを除去する
    If Not (g_objCnn Is Nothing) Then
        If g_objCnn.State And adStateOpen Then g_objCnn.Close
        Set g_objCnn = Nothing
    End If
    
    ' ADO Connection オブジェクトを作成する
    Set g_objCnn = New ADODB.Connection
    ' パラメータを設定する
    sParam = "dsn=" & sDsn & ";uid=" & sUserId & ";pwd=" & sPassword
    g_objCnn.ConnectionString = sParam
    g_objCnn.ConnectionTimeout = 30
    g_objCnn.CommandTimeout = 120
    ' オープンする
    g_objCnn.Open
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// AdoClose
Public Sub AdoClose()
    ' ADO Connection Close
    If Not (g_objCnn Is Nothing) Then
        If g_objCnn.State And adStateOpen Then g_objCnn.Close
        Set g_objCnn = Nothing
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// SetParam
Public Sub SetParam(ByVal sVar As String, ByVal sValue As String)
    Dim i As Long, nNum As Long
    ' Make VarName
    sVar = Trim$(sVar)
    sVar = "#" & UCase$(sVar) & "#"
    ' 変数を探す
    nNum = UBound(g_asParam, 2)
    For i = 0 To nNum - 1
        If g_asParam(0, i) = sVar Then Exit For
    Next
    If i = nNum Then
        ' 新規に変数を登録する
        ReDim Preserve g_asParam(1, nNum + 1)
        g_asParam(0, i) = sVar
    End If
    g_asParam(1, i) = sValue
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// ReplaceParam
Public Function ReplaceParam(ByVal sValue As String) As String
    Dim i As Long, nParamNum As Long
    Dim sCmd As String, sCmdPos As String, sCell As String
    Dim nRow As Long, nCol As Long

    ' 変数の数を求める
    nParamNum = UBound(g_asParam, 2)
    ' 変数を展開する
    For i = 0 To nParamNum - 1
        sValue = Replace(sValue, g_asParam(0, i), g_asParam(1, i), 1, -1, vbTextCompare)
    Next
    
    ' Cell
    Do While InStr(1, sValue, "#CELL(", vbTextCompare) <> 0
        sCmd = Right$(sValue, Len(sValue) - InStr(1, sValue, "#CELL(", vbTextCompare))
        sCmd = Left$(sCmd, InStr(sCmd, "#") - 1)
        sCmdPos = Mid$(sCmd, InStr(sCmd, "("), InStr(sCmd, ")") - InStr(sCmd, "(") + 1)
        GetPos sCmdPos, nRow, nCol
        sCell = g_objSheetData.Cells(nCol, nRow).Value
        sValue = Replace(sValue, "#" & sCmd & "#", sCell, 1, 1, vbTextCompare)
    Loop
    
    ' 変数が残っていないか調べる
    ' If InStr(sValue, "#") Then
    '     Err.Raise USERERR, , "指定された変数は値が指定されていません。"
    ' End If
    ReplaceParam = sValue
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// GetPos
Private Sub GetPos(ByVal sPos As String, nRow As Long, nCol As Long)
    Dim i As Long
    Dim sChr As String
    
    ' Del ()
    sPos = Trim(sPos)
    sPos = Mid$(sPos, 2, Len(sPos) - 2)
    If InStr(sPos, ",") Then
        ' 数式指定
        nRow = CLng(Left$(sPos, InStr(sPos, ",") - 1))
        nCol = CLng(Right$(sPos, Len(sPos) - InStr(sPos, ",")))
    Else
        ' 文字式指定
        i = 1: nRow = 0: nCol = 0
        Do While Mid$(sPos, i, 1) <> ""
            sChr = Mid$(sPos, i, 1)
            Select Case sChr
            Case "a" To "z":
                nRow = nRow * 26 + (Asc(sChr) - Asc("a")) + 1
            Case "A" To "Z":
                nRow = nRow * 26 + (Asc(sChr) - Asc("A")) + 1
            Case "0" To "9":
                Exit Do
            Case Else
                Err.Raise USERERR, , "コマンドの指定が間違っています。"
            End Select
            i = i + 1
        Loop
        Do While Mid$(sPos, i, 1) <> ""
            sChr = Mid$(sPos, i, 1)
            Select Case sChr
            Case "0" To "9":
                nCol = nCol * 10 + (Asc(sChr) - Asc("0"))
            Case Else
                Err.Raise USERERR, , "コマンドの指定が間違っています。"
            End Select
            i = i + 1
        Loop
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// ExecSql
Private Sub ExecSql(ByVal sSql As String, ByVal nRow As Long, ByVal nCol As Long)
    Dim i As Long
    Dim sValue As String
    Dim objRst As ADODB.Recordset


    ' ADO RecordSet オブジェクトオープン
    Set objRst = New ADODB.Recordset

    objRst.Open sSql, g_objCnn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' 取得したテーブルの値をセルに書き込む
    If Not (objRst.BOF Or objRst.EOF) Then
        objRst.MoveFirst
        Do While Not objRst.EOF
            For i = 0 To objRst.Fields.Count - 1
                g_objSheetData.Cells(nCol, nRow + i).Value = objRst.Fields(i).Value
            Next
            nCol = nCol + 1
            objRst.MoveNext
        Loop
    End If
    
    ' ADO RecordSet オブジェクト除去
    objRst.Close
    Set objRst = Nothing
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// CommandProcess
Private Sub CommandProcess()
    Dim i As Long, nRow As Long, nCol As Long
    Dim sCmd As String, sValue As String, sVar As String
    Dim sCmdType As String, sCmdPos As String
    Dim sDsn As String, sUserId As String, sPassword As String
    Dim nComboItemCnt As Long, aComboItem() As String, aComboItemId() As Long
    Dim sActiveSheet As String
    Dim bMandatory As Boolean
    
    g_nPc = 0
    Do While True
        ' CountUp g_nPc
        g_nPc = g_nPc + 1
        ' コマンド取得
        sCmd = Trim$(UCase$(g_objSheetCmd.Cells(g_nPc, 1).Value))
        ' End
        If sCmd = "" Then Exit Do
        ' コメント
        If Mid$(sCmd, 1, 1) = "[" Then sCmd = "Nothing"
        
        ' 座標付コマンド
        If InStr(sCmd, "(") Then
            ' Check ")"
            If Mid$(sCmd, Len(sCmd), 1) <> ")" Then
                Err.Raise USERERR, , "コマンドの指定が間違っています。"
            End If
            ' コマンド取得
            sCmdType = Trim$(Left$(sCmd, InStr(sCmd, "(") - 1))
            ' 座標取得
            sCmdPos = Mid$(sCmd, InStr(sCmd, "("), InStr(sCmd, ")") - InStr(sCmd, "(") + 1)
            GetPos sCmdPos, nRow, nCol
            If nRow = 0 Or nCol = 0 Then
                Err.Raise USERERR, , "座標の指定が正しくありません。"
            End If
            ' 引数展開
            sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
            ' Check SheetData
            If g_objSheetData Is Nothing Then
                Err.Raise USERERR, , "書込先シートが指定されていません。"
            End If
            ' Execute Cmd
            Select Case sCmdType
            Case "SQL"
                Call ExecSql(sValue, nRow, nCol)
            Case "CELL"
                g_objSheetData.Cells(nCol, nRow).Value = sValue
            Case Else
                Err.Raise USERERR, , "コマンドの指定が間違っています。"
            End Select
            sCmd = "Nothing"
        End If
        
        ' 単一コマンド
        Select Case sCmd
        Case "LET"
            sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 3).Value)
            SetParam g_objSheetCmd.Cells(g_nPc, 2).Value, sValue
        Case "SHEET"
            sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
            ' データシート指定
            If Not (g_objSheetData Is Nothing) Then
                Set g_objSheetData = Nothing
            End If
            On Error Resume Next
            Set g_objSheetData = g_objWorkbook.Sheets(sValue)
            If Err.Number Then
                On Error GoTo 0
                Err.Raise USERERR, , "シート " & sValue & " が有りません。"
            End If
            On Error GoTo 0
         'add,2007/11/09,S----------
         Case "PAGENUMBER"
            PageNumber
            PageCheck
         Case "ADDSHEET"
            AddSheet
        Case "DELSHEET"
            DelSheet
         Case "PAGESQL"
            PageSql
            PageCheck
        Case "NEXTPAGE"
            g_nPageNo = g_nPageNo + 1
            SetParam "PageNo", CStr(g_nPageNo + 1)
            ' ページ数をチェック
            If g_nPageNo <> UBound(g_avPageId) Then
                SetParam "PageID", g_avPageId(g_nPageNo)
                g_nPc = g_nPageSetupPc ' Goto PageSetup
            End If
            
         'add,2007/11/09,E----------
        Case "ODBC"
            sDsn = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
            sUserId = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 3).Value)
            sPassword = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 4).Value)
            ' AdoConnect
            AdoConnect sDsn, sUserId, sPassword
        Case "MSGBOX"
            sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
            MsgBox sValue, vbOKOnly, "ExcelReport"
            DoEvents
        Case "DIALOGINIT"
            Unload dlgExcelReportInput
            Load dlgExcelReportInput
        Case "TITLE"
            sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
            dlgExcelReportInput.SetTitle sValue
        Case "LABEL"
            sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
            bMandatory = (ReplaceParam(g_objSheetCmd.Cells(g_nPc, 3).Value) = 1)
            dlgExcelReportInput.AddLabel sValue, bMandatory
        Case "TEXT"
            If g_objSheetCmd.Cells(g_nPc, 4).Value = "<SQL>" Then
                sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 5).Value)
                Call GetDafaultDataFromDB(sValue, sValue)
            Else
                sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 4).Value)
            End If
            bMandatory = (ReplaceParam(g_objSheetCmd.Cells(g_nPc, 3).Value) = 1)
            dlgExcelReportInput.AddText g_objSheetCmd.Cells(g_nPc, 2).Value, sValue, bMandatory
        Case "TEXTPASSWORD"
            If g_objSheetCmd.Cells(g_nPc, 4).Value = "<SQL>" Then
                sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 5).Value)
                Call GetDafaultDataFromDB(sValue, sValue)
            Else
                sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 4).Value)
            End If
            bMandatory = (ReplaceParam(g_objSheetCmd.Cells(g_nPc, 3).Value) = 1)
            dlgExcelReportInput.AddTextPassword g_objSheetCmd.Cells(g_nPc, 2).Value, sValue, bMandatory
        Case "DATE"
            If g_objSheetCmd.Cells(g_nPc, 4).Value = "<SQL>" Then
                sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 5).Value)
                Call GetDafaultDataFromDB(sValue, sValue)
            Else
                sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 4).Value)
            End If
            If sValue <> "" And IsDate(sValue) Then
                sValue = Format$(CDate(sValue), "yyyy/mm/dd")
            Else
                sValue = Format$(Now, "yyyy/mm/dd")
            End If
            bMandatory = (ReplaceParam(g_objSheetCmd.Cells(g_nPc, 3).Value) = 1)
            dlgExcelReportInput.AddDate g_objSheetCmd.Cells(g_nPc, 2).Value, sValue, bMandatory
        Case "COMBOBOX"
            nComboItemCnt = 0
            Select Case g_objSheetCmd.Cells(g_nPc, 4).Value
            Case "<SQL>"
                sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 5).Value)
                If sValue = "" Then GoTo EndSelectCOMBOBOX
                Call MakeCmbListFromDB(sValue, aComboItem, aComboItemId)
                If g_objSheetCmd.Cells(g_nPc, 6).Value = 1 Then
                    nComboItemCnt = UBound(aComboItem) + 1
                    ReDim Preserve aComboItem(nComboItemCnt)
                    ReDim Preserve aComboItemId(nComboItemCnt)
                    aComboItem(nComboItemCnt) = ""
                    aComboItemId(nComboItemCnt) = -1
                End If
'            Case "<SQL2>"
'                sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 4).Value)
'                If sValue = "" Then GoTo EndSelectCOMBOBOX
'                Call MakeCmbListFromDB(sValue, aComboItem, aComboItemId)
'                If g_objSheetCmd.Cells(g_nPc, 5).Value = 1 Then
'                    nComboItemCnt = UBound(aComboItem) + 1
'                    ReDim Preserve aComboItem(nComboItemCnt)
'                    ReDim Preserve aComboItemId(nComboItemCnt)
'                    aComboItem(nComboItemCnt) = ""
'                    aComboItemId(nComboItemCnt) = -1
'                End If
            Case Else
                For i = 4 To 100
                    sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, i).Value)
                    If sValue = "" Then Exit For
                    ReDim Preserve aComboItem(nComboItemCnt)
                    ReDim Preserve aComboItemId(nComboItemCnt)
                    aComboItem(nComboItemCnt) = sValue
                    aComboItemId(nComboItemCnt) = -1
                    nComboItemCnt = nComboItemCnt + 1
                Next
            End Select
            bMandatory = (ReplaceParam(g_objSheetCmd.Cells(g_nPc, 3).Value) = 1)
            dlgExcelReportInput.AddComboBox g_objSheetCmd.Cells(g_nPc, 2).Value, aComboItem, aComboItemId, bMandatory
            Erase aComboItem
EndSelectCOMBOBOX:
        Case "DIALOGSHOW"
            dlgExcelReportInput.AdjustSize
            dlgExcelReportInput.Show vbModal
            If g_bInput = False Then Err.Raise USERERR, , "ユーザーによる入力中断"
            DoEvents
        Case "EXECUTEMACRO"
            sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
            Call g_objExcel.Run(sValue)
        Case "ACTIVESHEET"
            sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
            sActiveSheet = sValue
        Case "Nothing"
        Case Else
            Err.Raise USERERR, , "コマンドの指定が間違っています。"
        End Select
    Loop

    ' ADO Connection オブジェクトを除去する
    AdoClose

    If sActiveSheet <> "" Then
    'ActiveSheet指定有り時
        On Error Resume Next
        g_objWorkbook.Sheets(sActiveSheet).Activate
        g_objWorkbook.Sheets(sActiveSheet).Range("A1").Select
        If Err.Number Then
            On Error GoTo 0
            Err.Raise USERERR, , "シート " & sValue & " が有りません。"
            GoTo DefaultSet
        End If
        On Error GoTo 0

    Else

    ' セル"A1"を選択状態にする
DefaultSet:
        On Error Resume Next
        If Not (g_objSheetData Is Nothing) Then
            g_objSheetData.Activate
            g_objSheetData.Range("A1").Select
            Set g_objSheetData = Nothing
        End If
    End If

End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////
'// DelSheet
Private Sub DelSheet()
    Dim sValue As String
    Dim objDelSheet As Object
    
    ' シート選択
    sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
    On Error Resume Next
    Set objDelSheet = g_objWorkbook.Sheets(sValue)
    If Err.Number Then
        On Error GoTo 0
        Err.Raise USERERR, , "シート " & sValue & " が有りません。"
    End If
    On Error GoTo 0
    
    ' 現在のシートを開放する
    If Not (g_objSheetData Is Nothing) Then
        If g_objSheetData.Index = objDelSheet.Index Then Set g_objSheetData = Nothing
    End If
    ' シート削除
    objDelSheet.Delete
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// AddSheet
Private Sub AddSheet()
    Dim sValue As String
    
    ' 現在のシートを開放する
    If Not (g_objSheetData Is Nothing) Then Set g_objSheetData = Nothing
    ' シート追加
    sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
    If sValue = "" Then
        g_objWorkbook.Sheets.Add After:=g_objWorkbook.Sheets(g_objWorkbook.Sheets.Count)
        Set g_objSheetData = g_objWorkbook.Sheets(g_objWorkbook.Sheets.Count)
    Else
        ' シート選択
        On Error Resume Next
        Set g_objSheetData = g_objWorkbook.Sheets(sValue)
        If Err.Number Then
            On Error GoTo 0
            Err.Raise USERERR, , "シート " & sValue & " が有りません。"
        End If
        On Error GoTo 0
        g_objSheetData.Copy After:=g_objWorkbook.Sheets(g_objWorkbook.Sheets.Count)
        Set g_objSheetData = g_objWorkbook.Sheets(g_objWorkbook.Sheets.Count)
    End If
    ' シート名の設定
    sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 3).Value)
    If sValue <> "" Then g_objSheetData.Name = sValue
    ' アクティブシートに指定
    g_objSheetData.Activate
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// PageNumber
Private Sub PageNumber()
    Dim i As Long, nRow As Long
    Dim nStart As Long, nEnd As Long
    
    ' Clear g_avPageId
    Erase g_avPageId
    ReDim g_avPageId(0)

    ' 数字を列挙する
    nStart = CLng(ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value))
    nEnd = CLng(ReplaceParam(g_objSheetCmd.Cells(g_nPc, 3).Value))
    nRow = 0
    For i = nStart To nEnd
       ReDim Preserve g_avPageId(nRow + 1)
       g_avPageId(nRow) = CStr(i)
       nRow = nRow + 1
    Next
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// PageList
Private Sub PageList()
    Dim i As Long, nRow As Long
    Dim sValue As String
    
    ' Clear g_avPageId
    Erase g_avPageId
    ReDim g_avPageId(0)

    ' セルを列挙する
    nRow = 0
    For i = 2 To 100
       sValue = ReplaceParam(g_objSheetCmd.Cells(g_nPc, i).Value)
       If sValue = "" Then Exit For
       ReDim Preserve g_avPageId(nRow + 1)
       g_avPageId(nRow) = sValue
       nRow = nRow + 1
    Next
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// PageSql
Private Sub PageSql()
    Dim nRow As Long
    Dim sSql As String
    Dim objRst As ADODB.Recordset
    
    sSql = ReplaceParam(g_objSheetCmd.Cells(g_nPc, 2).Value)
    
    ' ADO RecordSet オブジェクトオープン
    Set objRst = New ADODB.Recordset
    objRst.Open sSql, g_objCnn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Clear g_avPageId
    Erase g_avPageId
    ReDim g_avPageId(0)

    ' 取得したテーブルの最初のフィールド値を保存する
    nRow = 0
    If Not (objRst.BOF Or objRst.EOF) Then
        objRst.MoveFirst
        Do While Not objRst.EOF
            ReDim Preserve g_avPageId(nRow + 1)
            g_avPageId(nRow) = objRst.Fields(0).Value
            nRow = nRow + 1
            objRst.MoveNext
        Loop
    End If
    
    ' ADO RecordSet オブジェクト除去
    objRst.Close
    Set objRst = Nothing
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// PageCheck
Private Sub PageCheck()
    ' Init PageInfo
    g_nPageNo = 0
    SetParam "PageNo", CStr(g_nPageNo)
    SetParam "PageID", ""

    ' ページ数をチェック
    If UBound(g_avPageId) = 0 Then
        ' Goto NextPage
        g_nPc = g_nPc + 1
        Do While g_objSheetCmd.Cells(g_nPc, 1).Value <> "" _
            And UCase$(g_objSheetCmd.Cells(g_nPc, 1).Value) <> "NEXTPAGE"
            g_nPc = g_nPc + 1
        Loop
    Else
        ' Init PageInfo
        g_nPageNo = 0
        SetParam "PageNo", CStr(g_nPageNo + 1)
        SetParam "PageID", g_avPageId(g_nPageNo)
        g_nPageSetupPc = g_nPc
    End If
End Sub

Private Sub MakeCmbListFromDB(psSQL As String, paComboItem() As String, paComboItemId() As Long)

Dim oRs As ADODB.Recordset
Dim nComboItemCnt As Long
Dim sValue As String

    Set oRs = g_objCnn.Execute(psSQL)
    nComboItemCnt = 0

    Do Until oRs.EOF
        If IsNull(oRs.Fields(0)) Then
            sValue = ""
        Else
            sValue = oRs.Fields(0)
        End If
        ReDim Preserve paComboItem(nComboItemCnt)
        ReDim Preserve paComboItemId(nComboItemCnt)
        paComboItem(nComboItemCnt) = sValue
        If oRs.Fields.Count = 2 Then
            If IsNull(oRs.Fields(1)) Then
                sValue = "-1"
            Else
                sValue = oRs.Fields(1)
            End If
        Else
            sValue = "-1"
        End If
        paComboItemId(nComboItemCnt) = sValue
        nComboItemCnt = nComboItemCnt + 1
NextData:
        oRs.MoveNext
    Loop

    oRs.Close
    Set oRs = Nothing

End Sub

Private Sub GetDafaultDataFromDB(psSQL As String, paData As String)

Dim oRs As ADODB.Recordset
Dim nComboItemCnt As Long
Dim sValue As String

    Set oRs = g_objCnn.Execute(psSQL)

    If Not oRs.EOF Then
        If IsNull(oRs.Fields(0)) Then
            paData = ""
        Else
            paData = oRs.Fields(0)
        End If
        oRs.Close
    Else
        paData = ""
    End If

    Set oRs = Nothing

End Sub
