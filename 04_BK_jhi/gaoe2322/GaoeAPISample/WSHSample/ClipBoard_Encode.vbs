'クリップボード内の文字列を公開鍵でEncodeする

Set gao = WScript.CreateObject("GaoEncode.GaoeAPI")

'Encodeのオプション設定
gao.Algorithm  = 0         'CAST - 128

'使用する公開鍵を指定
pass = "Sample|"	'複数使うときは Sample|Hoge| 見たいな感じ

'クリップボード内の文字列取得
clipdata = gao.GetClipStr()

'暗号化する
result = gao.EncodeStr(clipdata,pass,2)

If result = "" Then
  MsgBox "暗号化に失敗しました"
  WScript.Quit
End If

'クリップボードに内容をコピー
gao.SetClipStr result

MsgBox "クリップボードの文字列の暗号化は正常に終了しました" & vbCrLf & "暗号化した文字列をクリップボードに貼り付けました"