'�N���b�v�{�[�h���̕������閧��Decode����

Set gao = WScript.CreateObject("GaoEncode.GaoeAPI")

'Encode�̃I�v�V�����ݒ�
gao.Algorithm  = 0         'CAST - 128

'�g�p����閧�����w��
pass = "Sample|"	'�����g���Ƃ��� Sample|Hoge| �������Ȋ���

'�N���b�v�{�[�h���̕�����擾
clipdata = gao.GetClipStr()

'��������
result = gao.DecodeStr(clipdata,pass,2)

If result = "" Then
  MsgBox "�����Ɏ��s���܂���"
  WScript.Quit
End If

'�N���b�v�{�[�h�ɓ��e���R�s�[
gao.SetClipStr result

MsgBox "�N���b�v�{�[�h�̕�����̕����͐���ɏI�����܂���" & vbCrLf & "����������������N���b�v�{�[�h�ɓ\��t���܂���"
