'D&D���ꂽ�t�@�C����Decode����
Dim shell,target,gao,pass,args,ret

Set shell = WScript.CreateObject("WScript.Shell")
Set gao = WScript.CreateObject("GaoEncode.GaoeAPI")

'MyDocument�̏ꏊ���Q�b�g
target = shell.SpecialFolders("MyDocuments")

pass = "GaoEncode"          '�����L�[

'�t�@�C����D&D���ꂽ���ǂ����`�F�L
set args = WScript.Arguments
IF args.Count = 0 Then
    WScript.Quit
End IF

'��������
ret = gao.DecodeFile(args(0),pass,0,target & "\SampleDecode")

'���s�����Ƃ�����MsgBox���o��
If ret = 0 Then
    MsgBox "MyDocument��Encode�Ɏ��s���܂���"
End If