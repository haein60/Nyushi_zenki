'マイドキュメント中の全ファイルをEncodeする

Dim shell,target,gao,pass,name,ret

Set shell = WScript.CreateObject("WScript.Shell")
Set gao = WScript.CreateObject("GaoEncode.GaoeAPI")

'MyDocumentの場所をゲット
target = shell.SpecialFolders("MyDocuments")

'Encodeのオプション設定
gao.Algorithm   = 0         'CAST - 128
gao.DivideHi    = 0         '分割無し
gao.Compression = 1         'deflate圧縮
gao.Disguise    = 0         '偽装無し
gao.CryptoList  = 1         '情報隠蔽する

pass = "GaoEncode"          '解除キー
name = "SampleEnc(MyDoc)"   '保存名

'暗号化するファイルを登録
gao.AddTarget target

'SampleEnc(MyDoc).gaoを削除する
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
temp = target & "\SampleEncode\" & name & ".gao"
IF fso.FileExists(temp) Then
fso.DeleteFile(temp)
End IF


'暗号化する
ret = gao.EncodeFile(pass,0,target & "\SampleEncode",name)

'失敗したときだけMsgBoxを出す
If ret = 0 Then
    MsgBox "MyDocumentのEncodeに失敗しました"
End If
