'クリップボード内の文字列を秘密鍵Decodeする

Set gao = WScript.CreateObject("GaoEncode.GaoeAPI")

'Encodeのオプション設定
gao.Algorithm  = 0         'CAST - 128

'使用する秘密鍵を指定
pass = "Sample|"	'複数使うときは Sample|Hoge| 見たいな感じ

'クリップボード内の文字列取得
clipdata = gao.GetClipStr()

'復号する
result = gao.DecodeStr(clipdata,pass,2)

If result = "" Then
  MsgBox "復号に失敗しました"
  WScript.Quit
End If

'クリップボードに内容をコピー
gao.SetClipStr result

MsgBox "クリップボードの文字列の復号は正常に終了しました" & vbCrLf & "復号した文字列をクリップボードに貼り付けました"
