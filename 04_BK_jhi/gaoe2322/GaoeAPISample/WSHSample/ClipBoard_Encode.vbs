'�N���b�v�{�[�h���̕���������J����Encode����

Set gao = WScript.CreateObject("GaoEncode.GaoeAPI")

'Encode�̃I�v�V�����ݒ�
gao.Algorithm  = 0         'CAST - 128

'�g�p������J�����w��
pass = "Sample|"	'�����g���Ƃ��� Sample|Hoge| �������Ȋ���

'�N���b�v�{�[�h���̕�����擾
clipdata = gao.GetClipStr()

'�Í�������
result = gao.EncodeStr(clipdata,pass,2)

If result = "" Then
  MsgBox "�Í����Ɏ��s���܂���"
  WScript.Quit
End If

'�N���b�v�{�[�h�ɓ��e���R�s�[
gao.SetClipStr result

MsgBox "�N���b�v�{�[�h�̕�����̈Í����͐���ɏI�����܂���" & vbCrLf & "�Í���������������N���b�v�{�[�h�ɓ\��t���܂���"