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
' �[���O�u�ҏW
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
' �����񕪊�����
'
'-----------------------------------------------------
Public Function SplitString(pSrc As String, pLength As Integer) As String

    Dim AnsiWrkStr0 As String
    Dim AnsiWrkStr1 As String
    Dim AnsiWrkStr2 As String
    Dim AnsiSpace   As String
    Dim KanjiSplit  As Boolean
    
    '--- �p�����[�^�G���[
    If pLength <= 1 Then
        SplitString = ""
        pSrc = ""
        Exit Function
    End If
    
    '--- AnsiCode�� " " �����߂�
    AnsiSpace = StrConv(" ", vbFromUnicode)
    
    '--- �p�����[�^�� Ansi�ϊ�����
    AnsiWrkStr0 = StrConv(pSrc, vbFromUnicode)
    
    '--- �Z���ꍇ�̓X�y�[�X�Ŗ��߂Ă���
    If LenB(AnsiWrkStr0) < pLength Then
        AnsiWrkStr0 = LeftB(AnsiWrkStr0 & StrConv(Space(pLength), vbFromUnicode), pLength)
    End If
    
    '--- ���肪�ǂ������ł��邩�𔻒f����
    If AnsiIsKanjiSplit(AnsiWrkStr0, pLength) Then
        '--- �����̑O���Ő؂��ꍇ�́A1�O�ŕ���
        AnsiWrkStr1 = LeftB(AnsiWrkStr0, pLength - 1) & AnsiSpace
        AnsiWrkStr2 = MidB(AnsiWrkStr0, pLength)
    Else
        '--- ����̗ǂ��ꍇ�͂��̂܂ܕ���
        AnsiWrkStr1 = LeftB(AnsiWrkStr0, pLength)
        AnsiWrkStr2 = MidB(AnsiWrkStr0, pLength + 1)
    End If
    
    '--- ���������������Ԃ�l�Ƃ���
    SplitString = StrConv(AnsiWrkStr1, vbUnicode)
    
    '--- �������ꂽ�c���߂�
    pSrc = StrConv(AnsiWrkStr2, vbUnicode)
    
End Function

'-----------------------------------------------------
' ShiftJis���� 1Byte�ڔ���
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
' �����r����������
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

