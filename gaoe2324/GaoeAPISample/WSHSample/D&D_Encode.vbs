'D&Dされたファイル(フォルダ)をEncodeする

Dim shell,target,gao,pass,name,args,I,ret

Set shell = WScript.CreateObject("WScript.Shell")
Set gao = WScript.CreateObject("GaoEncode.GaoeAPI")

'MyDocumentの場所をゲット
target = shell.SpecialFolders("MyDocuments")

'Encodeのオプション設定
gao.Algorithm   = 1         'Blowfish
gao.DivideHi    = 250       '分割最大250KB
gao.DivideLo    = 200       '分割最小200KB
gao.Compression = 2         'CAB圧縮
gao.Disguise    = 3         'がおぇ内蔵Jpeg偽装
gao.CryptoList  = 0         '情報隠蔽しない

pass = "GaoEncode"          '解除キー
name = "SampleEnc"          '保存名

'暗号化するファイルを登録(D&Dされたファイル）
set args = WScript.Arguments
IF args.Count = 0 Then
    WScript.Quit
End IF
For I = 0 To args.Count - 1
   gao.AddTarget args(I)
Next

'暗号化する
ret = gao.EncodeFile(pass,0,target & "\SampleEncode",name)

'失敗したときだけMsgBoxを出す
If ret = 0 Then
    MsgBox "MyDocumentのEncodeに失敗しました"
End If
