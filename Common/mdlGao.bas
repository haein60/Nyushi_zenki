Attribute VB_Name = "mdlGao"
Option Explicit

Private prvoGao As Object

Public Sub gsOpenGao()

    '�g�p����
    Set objGao = CreateObject("GaoEncode.GaoeAPI")

End Sub

Public Sub gsCloseGao()

    '�J��
    Set objGao = Nothing

End Sub
' Encode�֐�
    'piAlgorithm:�A���S���Y��
    '   0=CSAT-128
    '   1=Blowfish
    '   2=TripleDES
    'plDivide:�����T�C�Y
    '   0=�����Ȃ�
    '   ���̑�=���̃T�C�Y�ŕ���[kbyte]
Public Function gfEncodeFile(piAlgorithm As Integer, plDivide As Integer) As Boolean
    'Encode
    
    Dim bOK As Boolean

    'piAlgorithm:�A���S���Y��
    '0=CSAT-128
    '1=Blowfish
    '2=TripleDES

    objGao.Algorithm = piAlgorithm

    '�����T�C�Y
    objGao.DivideHi = 0
    objGao.DivideLo = objGao.DivideHi
    
    'piCompression:���k
    '   0=���Ȃ�
    '   1=deflate
    '   2=cab
    objGao.Compression = cmbCompression.ListIndex

    'Disguise:�U��
    '   0=.gao
    '   1=.bmp
    '   2=.exe
    '   3=.jpg
    '   4=.lzh
    '   5=.Mid
    '   6=.wav
    '   7=�O���t�@�C��
    '   8=�O���t�H���_

    objGao.Disguise = cmbDisguise.ListIndex
    objGao.DisguiseEx = txtDisguiseEx.Text
    
    '���B��
    objGao.CryptoList = chkCryptoList.Value
    
    'Encode����t�@�C���̒ǉ�
    objGao.ClearTarget
    objGao.AddTarget txtTarget_Encode.Text
    
    'Encode�@�����Ă������
    bOK = objGao.EncodeFile(txtPass.Text, cmbMode.ListIndex, txtFolder_Encode.Text, txtOutName_Encode.Text)

    gfEncodeFile = bOK
        
End Function
' Encode�֐�
    'plAlgorithm:�A���S���Y��
    '   0=CSAT-128
    '   1=Blowfish
    '   2=TripleDES
Public Function gfEncodeStr(psStr As String, psPass As String, plAlgorithm As Long) As String

Dim sRet As String

    'piAlgorithm:�A���S���Y��
    '0=CSAT-128
    '1=Blowfish
    '2=TripleDES
    objGao.Algorithm = plAlgorithm

    'Encode�@�����Ă������
    sRet = objGao.EncodeStr(psStr, psPass, plAlgorithm)

    gfEncodeStr = sRet
        
End Function
