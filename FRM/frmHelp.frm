VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ヘルプ"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Label lblHelp 
      BackStyle       =   0  '透明
      Caption         =   "Label1"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim sSTR As String

    sSTR = "帳票設定についての注意事項" & vbCrLf
    sSTR = sSTR & "" & vbCrLf
    sSTR = sSTR & "１．合計などは左側科目を対象となる科目について選択状態の上、ボタンクリックにより" & vbCrLf
    sSTR = sSTR & "　　追加してください。選択した科目の合計等となります。" & vbCrLf
    sSTR = sSTR & "" & vbCrLf
    sSTR = sSTR & "２．合計等は登録後、真中のリストを選択すると対象の科目を右側のリストに表示します。" & vbCrLf
    sSTR = sSTR & "" & vbCrLf
    sSTR = sSTR & "３．合計などは８番目以降でなければ集計されません。" & vbCrLf
    sSTR = sSTR & "" & vbCrLf
    sSTR = sSTR & "４．細かい部分については取扱説明書を参照してください。" & vbCrLf

    lblHelp.Caption = sSTR

End Sub
