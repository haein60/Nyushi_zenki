VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "出力選択"
   ClientHeight    =   3480
   ClientLeft      =   2850
   ClientTop       =   1590
   ClientWidth     =   5490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5490
   Begin VB.CommandButton cmdCnf 
      Caption         =   "キャンセル"
      Height          =   525
      Index           =   1
      Left            =   3360
      TabIndex        =   12
      Top             =   2730
      Width           =   1905
   End
   Begin VB.CommandButton cmdCnf 
      Caption         =   "出力"
      Height          =   525
      Index           =   0
      Left            =   210
      TabIndex        =   11
      Top             =   2730
      Width           =   1905
   End
   Begin VB.CheckBox chkSemi 
      Caption         =   "セミナマスタ"
      Height          =   405
      Left            =   300
      TabIndex        =   1
      Top             =   780
      Width           =   1785
   End
   Begin VB.CheckBox chkCamp 
      Caption         =   "キャンペーンマスタ"
      Height          =   405
      Left            =   300
      TabIndex        =   0
      Top             =   180
      Width           =   2865
   End
   Begin VB.Frame fraSemi 
      Height          =   1575
      Left            =   210
      TabIndex        =   2
      Top             =   810
      Width           =   5055
      Begin VB.TextBox txtSemiTo 
         Height          =   360
         IMEMode         =   3  'ｵﾌ固定
         Left            =   3780
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "99999999"
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtSemiFrom 
         Height          =   360
         IMEMode         =   3  'ｵﾌ固定
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "00000000"
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtChiikiTo 
         Height          =   360
         IMEMode         =   3  'ｵﾌ固定
         Left            =   3780
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "99"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtChiikiFrom 
         Height          =   360
         IMEMode         =   3  'ｵﾌ固定
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblNyoro 
         Caption         =   "～"
         Height          =   255
         Index           =   1
         Left            =   3420
         TabIndex        =   10
         Top             =   1020
         Width           =   255
      End
      Begin VB.Label lblSemi 
         Caption         =   "セミナ受付番号"
         Height          =   225
         Left            =   330
         TabIndex        =   7
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label lblNyoro 
         Caption         =   "～"
         Height          =   255
         Index           =   0
         Left            =   3420
         TabIndex        =   6
         Top             =   540
         Width           =   255
      End
      Begin VB.Label lblChiiki 
         Caption         =   "地域コード"
         Height          =   225
         Left            =   330
         TabIndex        =   3
         Top             =   540
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private prvbChange As Boolean
Private prvbCampOut As Boolean
Private prvbSemiOut As Boolean
Private prvsChiikiFrom As String
Private prvsChiikiTo As String
Private prvsSemiFrom As String
Private prvsSemiTo As String

Public Function pfGetPrintObj(pbCampOut As Boolean, pbSemiOut As Boolean, psChiikiFrom As String, psChiikiTo As String, psSemiFrom As String, psSemiTo As String) As Boolean

    prvbChange = False
    prvbCampOut = pbCampOut
    prvbSemiOut = pbSemiOut
    prvsChiikiFrom = psChiikiFrom
    prvsChiikiTo = psChiikiTo
    prvsSemiFrom = psSemiFrom
    prvsSemiTo = psSemiTo

    Me.Show vbModal

    pbCampOut = prvbCampOut
    pbSemiOut = prvbSemiOut
    psChiikiFrom = prvsChiikiFrom
    psChiikiTo = prvsChiikiTo
    psSemiFrom = prvsSemiFrom
    psSemiTo = prvsSemiTo

    pfGetPrintObj = prvbChange

End Function

Private Sub chkSemi_Click()

Dim bChk As Boolean

    bChk = (chkSemi.Value = 1)

    txtChiikiFrom.Enabled = bChk
    txtChiikiTo.Enabled = bChk
    txtSemiFrom.Enabled = bChk
    txtSemiTo.Enabled = bChk

End Sub

Private Sub cmdCnf_Click(Index As Integer)

    If Index = 0 Then
        If Not (txtChiikiFrom.Text Like "[0-9][0-9]") Then
            MsgBox "地域コードの左側の入力が不正です。0から9までの数値２桁で入力してください。", vbOKOnly Or vbExclamation, "入力エラー"
            txtChiikiFrom.SetFocus
            Exit Sub
        End If
        If Not (txtChiikiTo.Text Like "[0-9][0-9]") Then
            MsgBox "地域コードの右側の入力が不正です。0から9までの数値２桁で入力してください。", vbOKOnly Or vbExclamation, "入力エラー"
            txtChiikiTo.SetFocus
            Exit Sub
        End If
        If txtChiikiFrom.Text > txtChiikiTo.Text Then
            MsgBox "地域コードの左側より右側が大きくなっています。右側のほうが大きくなるように入力してください。", vbOKOnly Or vbExclamation, "入力エラー"
            txtChiikiTo.SetFocus
            Exit Sub
        End If
        If Not (txtSemiFrom.Text Like "[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]") Then
            MsgBox "セミナ受付番号の左側の入力が不正です。0から9までの数値８桁で入力してください。", vbOKOnly Or vbExclamation, "入力エラー"
            txtSemiFrom.SetFocus
            Exit Sub
        End If
        If Not (txtSemiTo.Text Like "[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]") Then
            MsgBox "セミナ受付番号の左側の入力が不正です。0から9までの数値８桁で入力してください。", vbOKOnly Or vbExclamation, "入力エラー"
            txtSemiTo.SetFocus
            Exit Sub
        End If
        If txtSemiFrom.Text > txtSemiTo.Text Then
            MsgBox "セミナ受付番号の左側より右側が大きくなっています。右側のほうが大きくなるように入力してください。", vbOKOnly Or vbExclamation, "入力エラー"
            txtChiikiTo.SetFocus
            Exit Sub
        End If
        prvbChange = True
        prvbCampOut = (chkCamp.Value = 1)
        prvbSemiOut = (chkSemi.Value = 1)
        prvsChiikiFrom = txtChiikiFrom.Text
        prvsChiikiTo = txtChiikiTo.Text
        prvsSemiFrom = txtSemiFrom.Text
        prvsSemiTo = txtSemiTo.Text
    End If

    Unload Me

End Sub

Private Sub Form_Load()

    chkCamp.Value = IIf(prvbCampOut, 1, 0)
    chkSemi.Value = IIf(prvbSemiOut, 1, 0)
    txtChiikiFrom.Enabled = prvbSemiOut
    txtChiikiTo.Enabled = prvbSemiOut
    txtSemiFrom.Enabled = prvbSemiOut
    txtSemiTo.Enabled = prvbSemiOut
    txtChiikiFrom.Text = prvsChiikiFrom
    txtChiikiTo.Text = prvsChiikiTo
    txtSemiFrom.Text = prvsSemiFrom
    txtSemiTo.Text = prvsSemiTo

End Sub

Private Sub txtChiikiFrom_GotFocus()

    txtChiikiFrom.SelStart = 0
    txtChiikiFrom.SelLength = txtChiikiFrom.MaxLength

End Sub

Private Sub txtChiikiFrom_LostFocus()

    '数字ではないときエラー
'    If Trim(txtChiikiFrom) <> "" Then
'        If Not fncIs数字(txtChiikiFrom) Then
'            Call fncビープ(1)
'            txtChiikiFrom.SetFocus
'            Exit Sub
'        End If
'    End If

    If IsNumeric(txtChiikiFrom.Text) Then
        txtChiikiFrom.Text = Format(CLng(txtChiikiFrom.Text), "00")
    End If

End Sub

Private Sub txtChiikiTo_GotFocus()

    txtChiikiTo.SelStart = 0
    txtChiikiTo.SelLength = txtChiikiTo.MaxLength

End Sub

Private Sub txtChiikiTo_LostFocus()
    
    '数字ではないときエラー
'    If Trim(txtChiikiTo) <> "" Then
'        If Not fncIs数字(txtChiikiTo) Then
'            Call fncビープ(1)
'            txtChiikiTo.SetFocus
'            Exit Sub
'        End If
'    End If

    If IsNumeric(txtChiikiTo.Text) Then
        txtChiikiTo.Text = Format(CLng(txtChiikiTo.Text), "00")
    End If

End Sub

Private Sub txtSemiFrom_GotFocus()

    txtSemiFrom.SelStart = 0
    txtSemiFrom.SelLength = txtSemiFrom.MaxLength

End Sub

Private Sub txtSemiFrom_LostFocus()
    
    '数字ではないときエラー
'    If Trim(txtSemiFrom) <> "" Then
'        If Not fncIs数字(txtSemiFrom) Then
'            Call fncビープ(1)
'            txtSemiFrom.SetFocus
'            Exit Sub
'        End If
'    End If

    If IsNumeric(txtSemiFrom.Text) Then
        txtSemiFrom.Text = Format(CLng(txtSemiFrom.Text), "00000000")
    End If

End Sub

Private Sub txtSemiTo_GotFocus()

    txtSemiTo.SelStart = 0
    txtSemiTo.SelLength = txtSemiTo.MaxLength

End Sub

Private Sub txtSemiTo_LostFocus()

    '数字ではないときエラー
'    If Trim(txtSemiTo) <> "" Then
'        If Not fncIs数字(txtSemiTo) Then
'            Call fncビープ(1)
'            txtSemiTo.SetFocus
'            Exit Sub
'        End If
'    End If

    If IsNumeric(txtSemiTo.Text) Then
        txtSemiTo.Text = Format(CLng(txtSemiTo.Text), "00000000")
    End If

End Sub

