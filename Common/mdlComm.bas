Attribute VB_Name = "mdlComm"
'*******************************************************************************
'* 共通関数群                                                                  *
'* 作成日 : 2021.12.10                                                         *
'* 作成日 : jyon hein                                                          *
'*******************************************************************************

Option Explicit

Private Const GWL_STYLE                  As Long = (-16)
Private Const TVS_HASLINES               As Long = 2
Private Const TV_FIRST                   As Long = &H1100
Private Const TVM_SETBKCOLOR             As Long = (TV_FIRST + 29)

Private Declare Function SendMessage Lib "user32" _
                         Alias "SendMessageA" _
                         (ByVal hwnd As Long, _
                          ByVal wMsg As Long, _
                          ByVal wParam As Long, _
                          lParam As Any) As Long

Private Declare Function GetWindowLong Lib "user32" _
                         Alias "GetWindowLongA" _
                         (ByVal hwnd As Long, _
                          ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
                         Alias "SetWindowLongA" _
                         (ByVal hwnd As Long, _
                          ByVal nIndex As Long, _
                          ByVal dwNewLong As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32" _
                         (ByVal clr As OLE_COLOR, _
                          ByVal hPal As Long, _
                          dwRGB As Long) As Long


'*******************************************************************************
'* 数値変換して再度文字列に変換して0を削除する方法                             *
'* Val関数でDouble型に変換し､それをCStr関数で文字列に戻す関数です｡             *
'* 引数に変換したい文字列を渡して使います｡                                     *
'*******************************************************************************
Public Function DelZero(s As String) As String
    
    Dim ret As String
    
    ret = CStr(Val(s))
    
    DelZero = ret
    
End Function

'*******************************************************************************
'* 引数１：文字列                                                              *
'* 引数２：削除文字数                                                          *
'* 戻り値：削除後の文字列                                                      *
'*******************************************************************************
Public Function CutLeft(s As String, i As Long) As Variant
    
    Dim iLen    As Long     '文字列長
    
    
    '文字列ではない場合
'    If VarType(s) <> vbString Then
'        Exit Function
'    End If
    
    iLen = Len(s)
    
    ' 文字列長より指定文字数が大きい場合
    If iLen < i Then
        Exit Function
    End If
    
    
    ' 指定文字数を削除して返す
    If (Mid(s, 1, 1) = "0") Then
        CutLeft = CVar(Right(s, iLen - i))
    Else
        CutLeft = CVar(s)
    End If
    
End Function

'===============================================================================
' 指定の文字数になるまで先頭を文字で埋めます。
'
' @Param    stTarget    処理対象となる文字列。
' @Param    iLength     文字の長さ。
' @Param    [chOne]     埋める文字。
' @Return               先頭を指定の文字で iLength の長さまで埋められた文字列。
'===============================================================================
Public Function PadLeft(stTarget, iLength, chOne)
   
   Do While (Len(stTarget) < iLength)
       stTarget = chOne & stTarget
   Loop

   PadLeft = Right(stTarget, iLength)

End Function

'===============================================================================
'半角・全角が混在するような場合は、下記の関数を使用する
'2021.12.16 add jhi
'===============================================================================
Public Function fPadLeft(ByVal myData As String, ByVal CutLen As Long, ByVal CutStr As String) As String

    '文字を右寄せし、指定した文字列の文字数になるまで左側に指定した文字(0 や " " 等)を埋め込みます。
    Dim tmp As String

    tmp = StrConv(RightB$(StrConv(String$(CutLen, CutStr) & myData, vbFromUnicode), CutLen), vbUnicode)
    fPadLeft = tmp

End Function


'*******************************************************************************
'* debug用 log関数                                                             *
'* (簡易版:parameterが無条件string 1個のみ)                                    *
'*******************************************************************************
Public Sub log(ByVal str As String)

    ''''system date and timeを取得する
    Dim strDateTime    As String
    Dim sYM            As String
    Dim fName          As String
    
    Dim FileNumber     As Integer

    
    strDateTime = Format(Now, "yyyy/MM/dd HH:mm:ss") & " "
    
    'システム日付より年月を取得する
    sYM = Format(Now, "yyyymm")
    'Debug.Print sYM

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    
    'ファイルをAppendモードで開きます。
    fName = App.Path & "\log_" & sYM & ".txt"
    Open fName For Append As #FileNumber
    
    Print #FileNumber, strDateTime & str
    
    Close #FileNumber

End Sub

'*******************************************************************************
'* チェック結果を出力する csv関数                                              *
'*******************************************************************************
Public Sub logcsv(title_flag As Integer, ByVal str As String)

    'system date and timeを取得する
    Dim strDateTime    As String
    Dim sYM            As String
    Dim sYMD           As String
    Dim fName          As String
    
    Dim FileNumber     As Integer

    
    strDateTime = Format(Now, "yyyy/MM/dd HH:mm:ss") & ","
    
    'システム日付より年月を取得する
    sYM = Format(Now, "yyyymm")
    sYMD = Format(Now, "yyyymmdd")
''''gYMD = sYMD
    
    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    
    'ファイルをAppendモードで開きます。
    fName = ThisWorkbook.Path & "\log_" & sYMD & ".csv"
    Open fName For Append As #FileNumber
    
    'title_flagより、先頭にstrDateTimeを省略するか、出力する
    If (title_flag = 1) Then
        Print #FileNumber, str
    Else
        Print #FileNumber, strDateTime & str
    End If
    
    Close #FileNumber

End Sub

'*******************************************************************************
'* チェック結果を出力する csv関数                                              *
'*******************************************************************************
Public Sub logcsv_2(fn As String, title_flag As Integer, ByVal str As String)

    'system date and timeを取得する
    Dim strDateTime    As String
    Dim sYM            As String
    Dim sYMD           As String
    Dim fName          As String

    Dim FileNumber     As Integer

    
    strDateTime = Format(Now, "yyyy/MM/dd HH:mm:ss") & ","

    'システム日付より年月を取得する
    sYM = Format(Now, "yyyymm")
    sYMD = Format(Now, "yyyymmdd")
''''gYMD = sYMD
    
    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    
    'ファイルをAppendモードで開きます。
    fName = ThisWorkbook.Path & "\log" & fn & "_" & sYMD & ".csv"
    Open fName For Append As #FileNumber
    
    'title_flagより、先頭にstrDateTimeを省略するか、出力する
    If (title_flag = 1) Then
        Print #FileNumber, str
    Else
        Print #FileNumber, strDateTime & str
    End If
    
    Close #FileNumber

End Sub

'*******************************************************************************
'* 結果ファイルを削除して初期化をする                                          *
'*******************************************************************************
Public Sub Del_Csvfile(dummy As String)
    
    Dim FSO      As Object
    
    Dim sYMD     As String
    Dim fName    As String
    
    
    On Error Resume Next
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    
    'システム日付より年月を取得する
    sYMD = Format(Now, "yyyymmdd")
        
    fName = ThisWorkbook.Path & "\log_" & sYMD & ".csv"
    FSO.DeleteFile fName
    
    Set FSO = Nothing


End Sub
'*******************************************************************************
'* 結果ファイルを削除して初期化をする                                          *
'*******************************************************************************
Public Sub Del_Csvfile_2(fn As String)
    
    Dim FSO      As Object
    
    Dim sYMD     As String
    Dim fName    As String
    
    
    On Error Resume Next
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    
    'システム日付より年月を取得する
    sYMD = Format(Now, "yyyymmdd")
        
    fName = ThisWorkbook.Path & "\log" & fn & "_" & sYMD & ".csv"
    FSO.DeleteFile fName
    
    Set FSO = Nothing


End Sub

'*******************************************************************************
'* 選択したメニューからシステム定数Table tbSTESystemProfileにフェーズflagを    *
'* セットする                                                                  *
'*-----------------------------------------------------------------------------*
'* 2021.12.09 add jhi                                                          *
'*******************************************************************************
Public Sub Phase_FlagSet(phno As Integer)

    On Error GoTo ErrorHandler

    Dim l_obj_Rst      As New ADODB.Recordset
    Dim sSQL           As String
    Dim rinf           As Integer
    


    '-----------------------------------------------------------------------
    ' tbSTESystemProfile tableにフェーズflagをセットする
    '-----------------------------------------------------------------------
    sSQL = ""
    sSQL = "update tbSTESystemProfile set iCurrentPhase=" & phno & " where iActiveFlag=1"
    g_obj_Conn.Execute (sSQL)
        
    '
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
    End If

    Exit Sub


ErrorHandler:
    MsgBox Err.Description

End Sub

'*******************************************************************************
'* myMsgBox                                                                    *
'*-----------------------------------------------------------------------------*
'* 2021.12.09 add jhi                                                          *
'*******************************************************************************
Public Function myMsgBox(sMsg As String, sTit) As Long

    Dim rinf As Long

    rinf = MsgBox(sMsg, vbOKCancel, sTit)
    
    myMsgBox = rinf

End Function

'*******************************************************************************
'* Treeview Background change                                                  *
'*-----------------------------------------------------------------------------*
'* 2021.12.09 add jhi                                                          *
'*******************************************************************************
Public Sub SetTVBackColor(pobjTV As TreeView, plngBackColor As Long)
 
    Dim lngTVHwnd   As Long
    Dim lngStyle    As Long
    Dim objTVNode   As Node
    

    lngTVHwnd = pobjTV.hwnd
    
    ' Change the background
    Call SendMessage(lngTVHwnd, TVM_SETBKCOLOR, 0, ByVal plngBackColor)
    
    ' Set the backcolor of the nodes ...
    For Each objTVNode In pobjTV.Nodes
        objTVNode.BackColor = plngBackColor
    Next
 
    ' Reset the treeview style so the tree lines appear properly ...
    lngStyle = GetWindowLong(lngTVHwnd, GWL_STYLE)
    
    ' If the treeview has lines, temporarily remove them so the back
    ' repaints to the selected colour, then restore ...
    If lngStyle And TVS_HASLINES Then
       Call SetWindowLong(lngTVHwnd, GWL_STYLE, lngStyle Xor TVS_HASLINES)
       Call SetWindowLong(lngTVHwnd, GWL_STYLE, lngStyle)
    End If

    
End Sub


Public Function StrNullCut(psInStr As String) As String

    Dim lPos As Long


    lPos = InStr(1, psInStr, vbNullChar)

    If lPos > 0 Then
        StrNullCut = Left$(psInStr, lPos - 1)
    Else
        StrNullCut = psInStr
    End If

End Function

'*******************************************************************************
'【機能】    ファイルを別の場所へコピーします。
'CopyFile メソッド
'[参照設定]
'Microsoft Scripting Runtime (scrrun.dll)
'2022.02.08 add jhi
'*******************************************************************************
Public Sub fCopy(strSrcName As String, strDestName As String)

    On Error GoTo ErrorHandler

    'FileSystemObjectインスタンスを生成
    Dim FSO As Object


    Set FSO = CreateObject("Scripting.FileSystemObject")

    'ファイルをコピー
    FSO.CopyFile strSrcName, strDestName, True '上書き許可の場合

    'オブジェクトの解放
    Set FSO = Nothing
    Exit Sub


ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"

End Sub

