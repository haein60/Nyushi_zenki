Attribute VB_Name = "mdlGao"
Option Explicit

Private prvoGao As Object

Public Sub gsOpenGao()

    '使用準備
    Set objGao = CreateObject("GaoEncode.GaoeAPI")

End Sub

Public Sub gsCloseGao()

    '開放
    Set objGao = Nothing

End Sub
' Encode関数
    'piAlgorithm:アルゴリズム
    '   0=CSAT-128
    '   1=Blowfish
    '   2=TripleDES
    'plDivide:分割サイズ
    '   0=分割なし
    '   その他=そのサイズで分割[kbyte]
Public Function gfEncodeFile(piAlgorithm As Integer, plDivide As Integer) As Boolean
    'Encode
    
    Dim bOK As Boolean

    'piAlgorithm:アルゴリズム
    '0=CSAT-128
    '1=Blowfish
    '2=TripleDES

    objGao.Algorithm = piAlgorithm

    '分割サイズ
    objGao.DivideHi = 0
    objGao.DivideLo = objGao.DivideHi
    
    'piCompression:圧縮
    '   0=しない
    '   1=deflate
    '   2=cab
    objGao.Compression = cmbCompression.ListIndex

    'Disguise:偽装
    '   0=.gao
    '   1=.bmp
    '   2=.exe
    '   3=.jpg
    '   4=.lzh
    '   5=.Mid
    '   6=.wav
    '   7=外部ファイル
    '   8=外部フォルダ

    objGao.Disguise = cmbDisguise.ListIndex
    objGao.DisguiseEx = txtDisguiseEx.Text
    
    '情報隠蔽
    objGao.CryptoList = chkCryptoList.Value
    
    'Encodeするファイルの追加
    objGao.ClearTarget
    objGao.AddTarget txtTarget_Encode.Text
    
    'Encode　逝ってらっさい
    bOK = objGao.EncodeFile(txtPass.Text, cmbMode.ListIndex, txtFolder_Encode.Text, txtOutName_Encode.Text)

    gfEncodeFile = bOK
        
End Function
' Encode関数
    'plAlgorithm:アルゴリズム
    '   0=CSAT-128
    '   1=Blowfish
    '   2=TripleDES
Public Function gfEncodeStr(psStr As String, psPass As String, plAlgorithm As Long) As String

Dim sRet As String

    'piAlgorithm:アルゴリズム
    '0=CSAT-128
    '1=Blowfish
    '2=TripleDES
    objGao.Algorithm = plAlgorithm

    'Encode　逝ってらっさい
    sRet = objGao.EncodeStr(psStr, psPass, plAlgorithm)

    gfEncodeStr = sRet
        
End Function
