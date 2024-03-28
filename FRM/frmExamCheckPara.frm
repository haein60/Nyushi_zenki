VERSION 5.00
Begin VB.Form frmExamCheckPara 
   AutoRedraw      =   -1  'True
   Caption         =   "frmExamCheckPara : "
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmExamCheckPara.frx":0000
   ScaleHeight     =   5520
   ScaleWidth      =   12480
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdImput 
      Caption         =   "Web出願 CSVインポート"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   2415
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.CommandButton cmdDisp 
      Caption         =   "表示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3630
      TabIndex        =   6
      Top             =   2415
      Width           =   1600
   End
   Begin VB.TextBox txtJukenNumberTo 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   400
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1710
      MaxLength       =   4
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "[iNendo]"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtJukenNumberFrom 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   400
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1695
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "[iNendo]"
      Top             =   1905
      Width           =   1215
   End
   Begin VB.TextBox txtNendo 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   400
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1695
      MaxLength       =   4
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "[iNendo]"
      Top             =   1305
      Width           =   1215
   End
   Begin VB.Label lblGuidance 
      BackStyle       =   0  '透明
      Caption         =   "この機能は、frmBrowseに移したためここでは見えなくする。"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   5535
      TabIndex        =   8
      Top             =   2940
      Visible         =   0   'False
      Width           =   5340
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "件数"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   510
      TabIndex        =   5
      Tag             =   "1804"
      Top             =   2595
      Width           =   945
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "受験番号"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   510
      TabIndex        =   3
      Tag             =   "1804"
      Top             =   1995
      Width           =   1110
   End
   Begin VB.Label lblNendo 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "年度"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   510
      TabIndex        =   1
      Tag             =   "1804"
      Top             =   1395
      Width           =   825
   End
End
Attribute VB_Name = "frmExamCheckPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
'* 1.2 受験者データの編集                                                      *
'* Form Load                                                                   *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrProc

    Dim oRs                As ADODB.Recordset
    Dim sSQL               As String
    Dim l_int_CurrentPhase As Integer


    Me.Caption = "frmExamCheckPara : 受験者データの編集"

    txtNendo.Text = g_int_CurrentNendo

    ' get the current phase
    sSQL = "SELECT iCurrentPhase FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag=1" & _
        " AND iCurrentPhase IS NOT NULL"

    Set oRs = g_obj_Conn.Execute(sSQL)

    If Not oRs.EOF Then
        l_int_CurrentPhase = oRs.Fields("iCurrentPhase").Value
        oRs.Close
        Set oRs = Nothing
    Else
        l_int_CurrentPhase = 1
        Set oRs = Nothing
    End If

    If l_int_CurrentPhase = 0 Then
        txtJukenNumberTo.Text = "50"
    Else
        txtJukenNumberTo.Text = "1"
    End If

'    txtJukenNumberFrom.SetFocus



    ''''----------------------------------------------------------------------------------------
    ''''受験番号検索が1回OK->2回目NG->3回目OK->4回目NG の繰返しNGパターンの現象を修正
    ''''原因:frmExamineeCheckフォームが閉じられないなく裏に表示されていてそれを表示する模様なので
    ''''その閉じられないフォームがあれば強制的閉じるようにする
    ''''2023.01.24 add jhi
    ''''----------------------------------------------------------------------------------------
    Dim myObject As Object
    For Each myObject In Forms

        If myObject.Name = "frmExamineeCheck" Then
            ''''MsgBox myObject.Name
            Unload myObject
            Set myObject = Nothing
        End If

    Next

    Exit Sub

ErrProc:

End Sub

'*******************************************************************************
'* 1.2 受験者データの編集                                                      *
'* Form Activate                                                               *
'*******************************************************************************
Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim Index As Integer

    fMainForm.mnuTools.Enabled = False  ' disable tools menu

    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'*******************************************************************************
'* 1.2 受験者データの編集                                                      *
'* 【表示】ボタン処理                                                          *
'*******************************************************************************
Private Sub cmdDisp_Click()

    If Trim(txtNendo) = "" Then
        MsgBox "年度を入力してください。"
        Exit Sub
    End If

    gsExamCheckNendo = Trim(txtNendo)

    If Trim(txtJukenNumberFrom) = "" Then
        MsgBox "開始受験番号を入力してください。"
        Exit Sub
    End If

    gsExamIDFrom = txtJukenNumberFrom

    If Trim(txtJukenNumberTo) = "" Then
        MsgBox "件数を入力してください。"
        Exit Sub
    End If

    gsExamIDTo = txtJukenNumberTo

    frmExamineeCheck.Show
    frmExamineeCheck.ZOrder 0 'オブジェクトを Z オーダーの最前面(=0)に配置します。

    Unload Me

End Sub

'*******************************************************************************
'* 1.2 受験者データの編集                                                      *
'* 【インポート】ボタン処理                                                    *
'*******************************************************************************
Private Sub cmdImput_Click()

    'frmExamineeImport.Visible = True
    frmExamineeImport.txtNendo.Text = Me.txtNendo.Text
    frmExamineeImport.Show 1

End Sub

Private Sub txtJukenNumberFrom_KeyPress(KeyAscii As Integer)

    Call NumericOnly(Me, KeyAscii)

End Sub

Private Sub txtJukenNumberTo_KeyPress(KeyAscii As Integer)
    Call NumericOnly(Me, KeyAscii)
End Sub

Private Sub txtNendo_KeyPress(KeyAscii As Integer)
    Call NumericOnly(Me, KeyAscii)
End Sub
