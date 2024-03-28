VERSION 5.00
Begin VB.Form GaoeSample 
   Caption         =   "GaoeAPIサンプル"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdVersion 
      Caption         =   "GetVersion"
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtDisguiseEx 
      Height          =   270
      Left            =   2400
      TabIndex        =   14
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CheckBox chkCryptoList 
      Caption         =   "情報隠蔽"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
   End
   Begin VB.ComboBox cmbDisguise 
      Height          =   300
      ItemData        =   "GaoeSample.frx":0000
      Left            =   840
      List            =   "GaoeSample.frx":001F
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ComboBox cmbCompression 
      Height          =   300
      ItemData        =   "GaoeSample.frx":0069
      Left            =   3600
      List            =   "GaoeSample.frx":0076
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cmbDivide 
      Height          =   300
      ItemData        =   "GaoeSample.frx":0094
      Left            =   2160
      List            =   "GaoeSample.frx":00A4
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.ComboBox cmbAlgorithm 
      Height          =   300
      ItemData        =   "GaoeSample.frx":00CA
      Left            =   840
      List            =   "GaoeSample.frx":00D7
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decode"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txtPass 
      Height          =   270
      Left            =   3000
      TabIndex        =   8
      Top             =   2880
      Width           =   3135
   End
   Begin VB.ComboBox cmbMode 
      Height          =   300
      ItemData        =   "GaoeSample.frx":00FA
      Left            =   1800
      List            =   "GaoeSample.frx":0107
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtTarget_Decode 
      Height          =   270
      Left            =   3480
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox txtFolder_Decode 
      Height          =   270
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtOutName_Encode 
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox txtFolder_Encode 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtTarget_Encode 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "解除キー"
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Decode対象となるファイル"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "出力先フォルダ"
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "出力ファイル名（空欄-ランダム）"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "出力先フォルダ"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Encode対象となるファイル/フォルダ"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "GaoeSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objGao As Object    'GaoeAPI
'プロジェクト→参照設定　で　GaoEncode.tblを設定しておけば
'objGao As GaoeAPI とすることができます

Private Sub Form_Load()
    
    cmbMode.ListIndex = 0
    cmbAlgorithm.ListIndex = 0
    cmbDivide.ListIndex = 0
    cmbCompression.ListIndex = 0
    cmbDisguise.ListIndex = 0
    
    '使用準備
    Set objGao = CreateObject("GaoEncode.GaoeAPI")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '開放
    Set objGao = Nothing
    
End Sub

Private Sub cmdVersion_Click()
    'バージョン
    
    MsgBox objGao.GetVersion()

End Sub

Private Sub cmdEncode_Click()
    'Encode
    
    Dim bOK As Boolean
    
    'アルゴリズム
    objGao.Algorithm = cmbAlgorithm.ListIndex
    
    '分割サイズ
    Select Case cmbDivide.ListIndex
        Case 0 '分割無し
            objGao.DivideHi = 0
        Case 1 '300KB
            objGao.DivideHi = 300
        Case 2 '700KB
            objGao.DivideHi = 700
        Case 3 '1380KB
            objGao.DivideHi = 1380
    End Select
    objGao.DivideLo = objGao.DivideHi
    
    '圧縮
    objGao.Compression = cmbCompression.ListIndex
    
    '偽装
    objGao.Disguise = cmbDisguise.ListIndex
    objGao.DisguiseEx = txtDisguiseEx.Text
    
    '情報隠蔽
    objGao.CryptoList = chkCryptoList.Value
    
    'Encodeするファイルの追加
    objGao.ClearTarget
    objGao.AddTarget txtTarget_Encode.Text
    
    'Encode　逝ってらっさい
    bOK = objGao.EncodeFile(txtPass.Text, cmbMode.ListIndex, txtFolder_Encode.Text, txtOutName_Encode.Text)
    
    'どうよ？　どうなのよ？
    If bOK Then
        MsgBox "Encode成功"
    Else
        MsgBox "Encode失敗"
    End If
        
End Sub

Private Sub cmdDecode_Click()
    'Deocde
    
    Dim bOK As Boolean
    
    'Decode 逝ってらっさい
    bOK = objGao.DecodeFile(txtTarget_Decode.Text, txtPass.Text, cmbMode.ListIndex, txtFolder_Decode.Text)
    
    'どうよ？　どうなのよ
    If bOK Then
        MsgBox "Decode成功"
    Else
        MsgBox "Decode失敗"
    End If
    
End Sub

