'D&D���ꂽ�t�@�C��(�t�H���_)��Encode����

Dim shell,target,gao,pass,name,args,I,ret

Set shell = WScript.CreateObject("WScript.Shell")
Set gao = WScript.CreateObject("GaoEncode.GaoeAPI")

'MyDocument�̏ꏊ���Q�b�g
target = shell.SpecialFolders("MyDocuments")

'Encode�̃I�v�V�����ݒ�
gao.Algorithm   = 1         'Blowfish
gao.DivideHi    = 250       '�����ő�250KB
gao.DivideLo    = 200       '�����ŏ�200KB
gao.Compression = 2         'CAB���k
gao.Disguise    = 3         '����������Jpeg�U��
gao.CryptoList  = 0         '���B�����Ȃ�

pass = "GaoEncode"          '�����L�[
name = "SampleEnc"          '�ۑ���

'�Í�������t�@�C����o�^(D&D���ꂽ�t�@�C���j
set args = WScript.Arguments
IF args.Count = 0 Then
    WScript.Quit
End IF
For I = 0 To args.Count - 1
   gao.AddTarget args(I)
Next

'�Í�������
ret = gao.EncodeFile(pass,0,target & "\SampleEncode",name)

'���s�����Ƃ�����MsgBox���o��
If ret = 0 Then
    MsgBox "MyDocument��Encode�Ɏ��s���܂���"
End If
