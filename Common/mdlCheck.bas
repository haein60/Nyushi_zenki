Attribute VB_Name = "mdlCheck"
Option Explicit

Public Function gf_IntCheck(psNum As String) As Boolean

Dim iWk As Integer

On Error GoTo ErrProc

gf_IntCheck = False

    iWk = CInt(psNum)

gf_IntCheck = True

Exit Function

ErrProc:

End Function

Public Function gf_LongCheck(psNum As String) As Boolean

Dim lWk As Long

On Error GoTo ErrProc

gf_LongCheck = False

    lWk = CLng(psNum)

gf_LongCheck = True

Exit Function

ErrProc:

End Function

Public Function gf_DblCheck(psNum As String) As Boolean

Dim lWk As Long

On Error GoTo ErrProc

gf_DblCheck = False

    lWk = CDbl(psNum)

gf_DblCheck = True

Exit Function

ErrProc:

End Function
Public Function gf_FileCheck(psFile As String) As Boolean

Dim lWk As Long
Dim sWk As String
Dim sCkStr As String

On Error GoTo ErrProc

gf_FileCheck = False

    lWk = InStrRev(psFile, "\")

    If lWk = 0 Then Exit Function
    If lWk = Len(psFile) Then Exit Function

    sWk = StrConv(Mid(psFile, lWk + 1), vbUpperCase)

    sCkStr = Dir(Left(psFile, lWk), vbNormal)

    Do Until sCkStr = ""
        If StrConv(sCkStr, vbUpperCase) = sWk Then
            gf_FileCheck = True
            Exit Function
        End If
        sCkStr = Dir
    Loop

Exit Function

ErrProc:

End Function

'----------------------------------------------------
' 入力制限処理(0～9)
'
'----------------------------------------------------
Public Sub NumericOnly(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' 入力制限処理(0～9 & Period)
'
'----------------------------------------------------
Public Sub NumericPeriod(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case "."
            If InStr(F.ActiveControl, ".") = 0 Then
                Exit Sub            '--- .(ピリオド)は１度だけ入力可
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' 入力制限処理(0～9 & Period & '-')
'
'----------------------------------------------------
Public Sub NumericPeriodMinus(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case "."
            If InStr(F.ActiveControl, ".") = 0 Then
                Exit Sub            '--- .(ピリオド)は１度だけ入力可
            End If
        Case "-"
            If Len(F.ActiveControl.Text) = 0 _
            Or F.ActiveControl.SelLength = Len(F.ActiveControl.Text) Then
                Exit Sub            '--- -(マイナス)は１度だけ、先頭のみ入力可
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' 入力制限処理(0～9 & '-')
'
'----------------------------------------------------
Public Sub NumericMinus(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
                Exit Sub            '--- 0～9は入力可
        Case "-"
            If InStr(F.ActiveControl, "-") = 0 Then
                If Len(F.ActiveControl.Text) = 0 _
                Or F.ActiveControl.SelLength = Len(F.ActiveControl.Text) Then
                    Exit Sub            '--- -(マイナス)は入力可
                End If
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

'--------------------------------------------------
' フォーカス取得追跡処理
'
'--------------------------------------------------
Public Sub GotFocusTracking(pForm As Form, pControl As Control)
', pInitIMEMode() As Integer
    Dim wIndex   As Integer
    
    wIndex = -1: On Error Resume Next: wIndex = pControl.index: On Error GoTo 0
    Call AllControlGotFocus(pForm, pControl, wIndex)
'    Call SetIMEMode(pForm, pInitIMEMode())
    
End Sub

'--------------------------------------------------
' フォーカス喪失追跡処理
'
'--------------------------------------------------
Public Sub LostFocusTracking(pForm As Form, pControl As Control)

    Dim wIndex   As Integer
    
    wIndex = -1: On Error Resume Next: wIndex = pControl.index: On Error GoTo 0
    Call AllControlLostFocus(pForm, pControl, wIndex)
    
End Sub


'--------------------------------------------------
' 全コントロールフォーカス取得時処理
'
'--------------------------------------------------
Private Sub AllControlGotFocus(pForm As Form, pGotControl As Control, pGotIndex As Integer)

    Dim strLostControl As String
    Dim strGotControl  As String
    
    strGotControl = pGotControl.Name
    
    On Error Resume Next
    If TypeOf pGotControl Is CommandButton Or _
       TypeOf pGotControl Is OptionButton Or _
       TypeOf pGotControl Is CheckBox Then
        '--- コマンドボタン/オプションボタン/チェックボックスの場合は
        '    フォントのボールド設定を On/Off する
        If pGotIndex = -1 Then
            pForm.Controls(strGotControl).Font.Bold = True
        Else
            pForm.Controls(strGotControl)(pGotIndex).Font.Bold = True
        End If
    ElseIf TypeOf pGotControl Is TextBox Or _
           TypeOf pGotControl Is ListBox Or _
           TypeOf pGotControl Is ComboBox Then
        '--- テキストボックス/リストボックス/コンボボックスの場合は
        '    バックカラーを変更する
'        If pGotIndex = -1 Then
'            If pForm.Controls(strGotControl).Locked = False Then
'                pForm.Controls(strGotControl).BackColor = gGotFocusBackColor
'                Call FTCOverWriteMode(pForm)
'            End If
'        Else
'            If pForm.Controls(strGotControl)(pGotIndex).Locked = False Then
'                pForm.Controls(strGotControl)(pGotIndex).BackColor = gGotFocusBackColor
'                Call FTCOverWriteMode(pForm)
'            End If
'        End If
    End If
    On Error GoTo 0

End Sub

'--------------------------------------------------
' 全コントロールフォーカス喪失時処理
'
'--------------------------------------------------
Private Sub AllControlLostFocus(pForm As Form, pLostControl As Control, pLostIndex As Integer)

    Dim strLostControl As String
    Dim strGotControl  As String
    
    strLostControl = pLostControl.Name
    
    On Error Resume Next
    If TypeOf pLostControl Is CommandButton Or _
       TypeOf pLostControl Is OptionButton Or _
       TypeOf pLostControl Is CheckBox Then
        '--- コマンドボタン/オプションボタン/チェックボックスの場合は
        '    フォントのボールド設定を On/Off する
        If pLostIndex = -1 Then
            pForm.Controls(strLostControl).Font.Bold = False
        Else
            pForm.Controls(strLostControl)(pLostIndex).Font.Bold = False
        End If
    ElseIf TypeOf pLostControl Is TextBox Or _
           TypeOf pLostControl Is ListBox Or _
           TypeOf pLostControl Is ComboBox Then
        '--- テキストボックス/リストボックス/コンボボックスの場合は
        '    バックカラーを変更する
'        If pLostIndex = -1 Then
'            pForm.Controls(strLostControl).BackColor = gLostFocusBackColor
'        Else
'            pForm.Controls(strLostControl)(pLostIndex).BackColor = gLostFocusBackColor
'        End If
    End If
    On Error GoTo 0
    
End Sub

'-----------------------------------------------------
' クリティカルセッション開始処理
'
'-----------------------------------------------------
Public Function BeginCriticalSession(pSessionKey As String) As Integer

    Dim fp As Integer

    fp = FreeFile

    On Error Resume Next
    
    '--- OSのファイルロックを利用してロックを行う
    Open pSessionKey & ".LCK" For Output Lock Write As #fp
    
    
    Close #fp
    Open pSessionKey & ".LCK" For Output Lock Write As #fp


    Do Until Err = 0
        Err = 0
        Close #fp
        DoEvents
        Open pSessionKey & ".LCK" For Output Lock Write As #fp
    Loop
    Print #fp, App.EXEName
    BeginCriticalSession = fp
    
End Function

'-----------------------------------------------------
' クリティカルセッション終了処理
'
'-----------------------------------------------------
Public Sub EndCriticalSession(pFp As Integer)

    Close #pFp
    
End Sub

'--------------------------------------------------------
'   関数名  : GetFileExistence
'   用途    : ファイルが存在するかどうか調べる
'   引数    : strPathName ファイル・ディレトクリ(パス)名
'   戻り値  : True ファイルは存在する
'             False ファイルは存在しない
'--------------------------------------------------------
Public Function GetFileExistence(strPathName As String) As Boolean
    
    '引数のサイズを格納/ファイル番号を格納
    Dim lngPNameSize As Long

    'エラーを無効にしておく
    On Error Resume Next

    If strPathName = "" Then
        '引数のファイル名・パス名がセットされていない
        'Nullをセットして
        GetFileExistence = ""
        '関数を抜ける
        Exit Function
    End If

    'パス名の最後にディレクトリ記号がある場合は削除
    If Right(strPathName, 1) = "\" Then
        
        'パスのサイズ-1を格納
        lngPNameSize = Len(strPathName) - 1
        '最後の一文字を取り除く
        strPathName = Left(strPathName, lngPNameSize)
    
    End If
    
    'ファイルを開いて、エラーかどうか確かめる
    '現在使用可能なファイル番号を割り振る
    lngPNameSize = FreeFile
    
    'では、開く
    Open strPathName For Input As lngPNameSize
    
    'エラー番号を調べる。0は「ファイルがあった」
    If Err = 0 Then
        '「ファイルがありました」をセット
        GetFileExistence = True
    Else
        '「ファイルは、なかったよ」をセット
        GetFileExistence = False
    End If
    
    Close lngPNameSize
    
    'エラー値を初期化
    Err = 0

End Function

'----------------------------------------------------
' 入力制限処理(0～9 & /)
'
'----------------------------------------------------
Public Sub gf_ChkDayInput(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case "/"
'            If InStr(F.ActiveControl, "/") = 0 Then
                Exit Sub            '--- /(ピリオド)は入力可
'            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' 入力制限処理(0～9 & :)
'
'----------------------------------------------------
Public Sub gf_ChkTimeInput(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case ":"
'            If InStr(F.ActiveControl, ":") = 0 Then
                Exit Sub            '--- :(ピリオド)は入力可
'            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' 入力制限処理(0～9 & / & :)
'
'----------------------------------------------------
Public Sub gf_ChkDateInput(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case "/"
'            If InStr(F.ActiveControl, "/") = 0 Then
                Exit Sub            '--- /(ピリオド)は入力可
'            End If
        Case ":"
'            If InStr(F.ActiveControl, ":") = 0 Then
                Exit Sub            '--- :(ピリオド)は入力可
'            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

Public Function gfNullChkStr(ByVal psInStr As Variant) As String

    If IsNull(psInStr) Then
        gfNullChkStr = ""
    Else
        gfNullChkStr = psInStr
    End If

End Function

Public Function gfNullChkStrTrim(ByVal psInStr As Variant) As String

    If IsNull(psInStr) Then
        gfNullChkStrTrim = ""
    Else
        gfNullChkStrTrim = Trim(psInStr)
    End If

End Function

'変換不可は０を戻す？
Public Function gfNullZeroChkInt(ByVal psInStr As Variant) As String

    If IsNull(psInStr) Then
        gfNullZeroChkInt = ""
    Else
        gfNullZeroChkInt = Trim(psInStr)
    End If

End Function

'----------------------------------------------------
' 入力制限処理(0～9 & Period )
'
'----------------------------------------------------
Public Sub NumericPeriodVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case "."
            If InStr(ovsfGrd.EditText, ".") = 0 Then
                Exit Sub            '--- .(ピリオド)は１度だけ入力可
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' 入力制限処理(0～9 & '-')
'
'----------------------------------------------------
Public Sub NumericMinusVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case "-"
            If Len(ovsfGrd.EditText) = 0 _
            Or ovsfGrd.EditSelLength = Len(ovsfGrd.EditText) Then
                Exit Sub            '--- -(マイナス)は１度だけ、先頭のみ入力可
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' 入力制限処理(0～9 & Period & '-')
'
'----------------------------------------------------
Public Sub NumericPeriodMinusVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case "."
            If InStr(ovsfGrd.EditText, ".") = 0 Then
                Exit Sub            '--- .(ピリオド)は１度だけ入力可
            End If
        Case "-"
            If MinusCheckVsfGrd(ovsfGrd, pKeyAscii) Then Exit Sub
'            If Len(ovsfGrd.EditText) = 0 _
'            Or ovsfGrd.EditSelLength = Len(ovsfGrd.EditText) Then
'                Exit Sub            '--- -(マイナス)は１度だけ、先頭のみ入力可
'            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0

End Sub

Private Function MinusCheckVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer) As Boolean
    MinusCheckVsfGrd = False
    If Len(ovsfGrd.EditText) = 0 _
    Or ovsfGrd.EditSelLength = Len(ovsfGrd.EditText) Then
        MinusCheckVsfGrd = True
        Exit Function            '--- -(マイナス)は１度だけ、先頭のみ入力可
    End If
End Function
