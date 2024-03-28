'D&DされたファイルをDecodeする
Dim shell,target,gao,pass,args,ret

Set shell = WScript.CreateObject("WScript.Shell")
Set gao = WScript.CreateObject("GaoEncode.GaoeAPI")

'MyDocumentの場所をゲット
target = shell.SpecialFolders("MyDocuments")

pass = "GaoEncode"          '解除キー

'ファイルがD&Dされたかどうかチェキ
set args = WScript.Arguments
IF args.Count = 0 Then
    WScript.Quit
End IF

'復号する
ret = gao.DecodeFile(args(0),pass,0,target & "\SampleDecode")

'失敗したときだけMsgBoxを出す
If ret = 0 Then
    MsgBox "MyDocumentのEncodeに失敗しました"
End If