'�}�C�h�L�������g���̑S�t�@�C����Encode����

Dim shell,target,gao,pass,name,ret

Set shell = WScript.CreateObject("WScript.Shell")
Set gao = WScript.CreateObject("GaoEncode.GaoeAPI")

'MyDocument�̏ꏊ���Q�b�g
target = shell.SpecialFolders("MyDocuments")

'Encode�̃I�v�V�����ݒ�
gao.Algorithm   = 0         'CAST - 128
gao.DivideHi    = 0         '��������
gao.Compression = 1         'deflate���k
gao.Disguise    = 0         '�U������
gao.CryptoList  = 1         '���B������

pass = "GaoEncode"          '�����L�[
name = "SampleEnc(MyDoc)"   '�ۑ���

'�Í�������t�@�C����o�^
gao.AddTarget target

'SampleEnc(MyDoc).gao���폜����
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
temp = target & "\SampleEncode\" & name & ".gao"
IF fso.FileExists(temp) Then
fso.DeleteFile(temp)
End IF


'�Í�������
ret = gao.EncodeFile(pass,0,target & "\SampleEncode",name)

'���s�����Ƃ�����MsgBox���o��
If ret = 0 Then
    MsgBox "MyDocument��Encode�Ɏ��s���܂���"
End If
