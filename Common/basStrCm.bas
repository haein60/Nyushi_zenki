Attribute VB_Name = "basStrCommon"
Option Explicit

Public Function fncNullChkStr(ByVal psInStr As Variant) As String

    If IsNull(psInStr) Then
        fncNullChkStr = ""
    Else
        fncNullChkStr = psInStr
    End If

End Function

Public Function fncAnsiLenB(ByVal psInStr As String) As String

    fncAnsiLenB = LenB(StrConv(psInStr, vbFromUnicode))

End Function

Public Function fncAnsiMidB(ByVal psInStr As String, plStart As Long, plLen As Long) As String

    fncAnsiMidB = StrConv(MidB(StrConv(psInStr, vbFromUnicode), plStart, plLen), vbUnicode)

End Function

Public Function fncAnsiLeftB(ByVal psInStr As String, plLen As Long) As String

    fncAnsiLeftB = fncAnsiMidB(psInStr, 1, plLen)

End Function

Public Function fncAnsiRightB(ByVal psInStr As String, plLen As Long) As String

    fncAnsiRightB = StrReverse(fncAnsiMidB(StrReverse(psInStr), 1, plLen))

End Function

Public Function fncSpaceRPad(psInStr As String, plLen As Long) As String

Dim lLen As Long

    lLen = LenB(StrConv(psInStr, vbFromUnicode))

    If lLen < plLen Then
        fncSpaceRPad = psInStr & Space(plLen - lLen)
    Else
        fncSpaceRPad = psInStr
    End If

End Function

Public Function fncSpaceLPad(psInStr As String, plLen As Long) As String

Dim lLen As Long

    lLen = LenB(StrConv(psInStr, vbFromUnicode))

    If lLen < plLen Then
        fncSpaceLPad = Space(plLen - lLen) & psInStr
    Else
        fncSpaceLPad = psInStr
    End If

End Function

Public Function fncStrAllSelect(psObj As Object)

    psObj.SelStart = 0
    psObj.SelLength = psObj.MaxLength

End Function
'-----------------------------------------------------
' ゼロ前置編集
'
'-----------------------------------------------------
Public Function Zero(pVal As Variant, pN As Integer) As Variant

    If IsNumeric(pVal) Then
        Zero = Right(String(pN, "0") & Format(pVal), pN)
    Else
        Zero = ""
    End If

End Function

'-----------------------------------------------------
' 文字列分割処理
'
'-----------------------------------------------------
Public Function SplitString(pSrc As String, pLength As Integer) As String

    Dim AnsiWrkStr0 As String
    Dim AnsiWrkStr1 As String
    Dim AnsiWrkStr2 As String
    Dim AnsiSpace   As String
    Dim KanjiSplit  As Boolean
    
    '--- パラメータエラー
    If pLength <= 1 Then
        SplitString = ""
        pSrc = ""
        Exit Function
    End If
    
    '--- AnsiCodeの " " を求める
    AnsiSpace = StrConv(" ", vbFromUnicode)
    
    '--- パラメータを Ansi変換する
    AnsiWrkStr0 = StrConv(pSrc, vbFromUnicode)
    
    '--- 短い場合はスペースで埋めておく
    If LenB(AnsiWrkStr0) < pLength Then
        AnsiWrkStr0 = LeftB(AnsiWrkStr0 & StrConv(Space(pLength), vbFromUnicode), pLength)
    End If
    
    '--- きりが良く分割できるかを判断する
    If AnsiIsKanjiSplit(AnsiWrkStr0, pLength) Then
        '--- 漢字の前半で切れる場合は、1つ前で分割
        AnsiWrkStr1 = LeftB(AnsiWrkStr0, pLength - 1) & AnsiSpace
        AnsiWrkStr2 = MidB(AnsiWrkStr0, pLength)
    Else
        '--- きりの良い場合はそのまま分割
        AnsiWrkStr1 = LeftB(AnsiWrkStr0, pLength)
        AnsiWrkStr2 = MidB(AnsiWrkStr0, pLength + 1)
    End If
    
    '--- 分割した文字列を返り値とする
    SplitString = StrConv(AnsiWrkStr1, vbUnicode)
    
    '--- 分割された残りを戻す
    pSrc = StrConv(AnsiWrkStr2, vbUnicode)
    
End Function

'-----------------------------------------------------
' ShiftJis漢字 1Byte目判定
'
'-----------------------------------------------------
Public Function AnsiIsKanji(pAnsiChar As String) As Boolean

    Dim wAnsiCharCode As Integer
    
    wAnsiCharCode = AscB(pAnsiChar)
    
    If &H80 <= wAnsiCharCode And wAnsiCharCode <= &H9F Or _
       &HE0 <= wAnsiCharCode And wAnsiCharCode <= &HFE Then
        AnsiIsKanji = True
    Else
        AnsiIsKanji = False
    End If
    
End Function

'-----------------------------------------------------
' 漢字途中分割判定
'
'-----------------------------------------------------
Public Function AnsiIsKanjiSplit(pSrc As String, pLength As Integer) As Boolean

    Dim i As Integer
    
    i = 0
    Do Until i >= pLength
        i = i + 1
        If AnsiIsKanji(MidB(pSrc, i, 1)) Then
            i = i + 1
        End If
    Loop
    If i = pLength Then
        AnsiIsKanjiSplit = False
    Else
        AnsiIsKanjiSplit = True
    End If
    
End Function

