VERSION 5.00
Begin VB.Form dlgHaitenWari 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "dlgHaitenWari : 配点割合"
   ClientHeight    =   3000
   ClientLeft      =   3480
   ClientTop       =   2040
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "dlgHaitenWari.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboSubject 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   2400
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   1200
      Width           =   1650
   End
   Begin VB.TextBox txtWariai 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1041
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1740
      Width           =   1650
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "更　新"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1740
      Width           =   1350
   End
   Begin VB.TextBox txtNendo 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   1
      Top             =   720
      Width           =   1650
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "割合"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1815
      Width           =   1815
   End
   Begin VB.Label lblErrorDetails 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   330
      TabIndex        =   4
      Top             =   2415
      Width           =   6855
   End
   Begin VB.Label lblZipCodeId 
      BackStyle       =   0  '透明
      Caption         =   "科目"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label lblZipCode 
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
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   780
      Width           =   1815
   End
End
Attribute VB_Name = "dlgHaitenWari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSearch_Click()
    On Error GoTo ERR_HANDLE

    Dim RS         As ADODB.Recordset
    Dim SQL        As String
    Dim strNendo   As String
    Dim strSubject As String
    Dim strWariai  As String
    Dim blnCommit  As Boolean
    
    '0.開始処理
    strNendo = Me.txtNendo.Text
    strSubject = Me.cboSubject.Text
    strWariai = Me.txtWariai.Text
    
    '1.画面データのチェック
    If Len(strNendo) <= 0 Then
         MsgBox "年度を入力してください。"
         Exit Sub
    End If
    If Len(strSubject) <= 0 Then
        MsgBox "科目を入力してください。"
        Exit Sub
    End If
    If Len(strWariai) <= 0 Then
        MsgBox "割合を入力してください。"
        Exit Sub
    End If
    
    strNendo = Trim(strNendo)
    strSubject = Trim(strSubject)
    strWariai = Trim(strWariai)
    
    If Len(strNendo) <> 4 Or (Not IsNumeric(strNendo)) Or strNendo > 2100 Or strNendo < 1970 Then
         MsgBox "年度を正しく入力してください。"
         Exit Sub
    End If
    
    If Len(strWariai) <= 0 Or (Not IsNumeric(strWariai)) Then
        MsgBox "割合を正しく入力してください。"
        Exit Sub
    End If
    
    If MsgBox("配点割合を実行します。よろしいですか？", vbOKCancel + vbInformation) = vbCancel Then
        Exit Sub
    End If
     
    g_obj_Conn.BeginTrans
    SQL = "Update tbSTELocks set iLocks = 1 where vTarget = 'tbSTEScoreProfile' "
    Call g_obj_Conn.Execute(SQL)
    blnCommit = True
    
    '2.画面のデータより、DBの成績をbackUp,
    'tbSTEScoreProfile_bak,tbSTEScoreDetail_bak

    SQL = ""
    SQL = SQL & "INSERT INTO tbSTEScoreProfile_bak("
    SQL = SQL & " iScoreProfileId"
    SQL = SQL & " ,iSubjectProfileId"
    SQL = SQL & " ,iExamineeProfileId"
    SQL = SQL & " ,fRawScore"
    SQL = SQL & " ,fChoseiScore"
    SQL = SQL & " ,iAbsentFlag"
    SQL = SQL & " ,dtCreate"
    SQL = SQL & " ,dtUpdate"
    SQL = SQL & " ,dtBakDate)"
    SQL = SQL & " SELECT  "
    SQL = SQL & "  sp.iScoreProfileId"
    SQL = SQL & " ,sp.iSubjectProfileId"
    SQL = SQL & " ,sp.iExamineeProfileId"
    SQL = SQL & " ,sp.fRawScore"
    SQL = SQL & " ,sp.fChoseiScore"
    SQL = SQL & " ,sp.iAbsentFlag"
    SQL = SQL & " ,sp.dtCreate,sp.dtUpdate,GETDATE()"
    SQL = SQL & " FROM tbSTEScoreProfile sp"
    SQL = SQL & " ,tbSTESubjectProfile sub,tbSTEExamineeProfile ep "
    SQL = SQL & " WHERE sp.iExamineeProfileId =ep.iExamineeProfileId "
    SQL = SQL & " AND sp.iSubjectProfileId =sub.iSubjectProfileId"
    SQL = SQL & " AND ep.iNendo=" & strNendo
    SQL = SQL & " AND sub.vSubjectName='" & strSubject & "'"
    SQL = SQL & " AND sub.iExamType=2"
    g_obj_Conn.Execute (SQL)
    
    SQL = ""
    SQL = SQL & "INSERT INTO tbSTEScoreDetail_bak("
    SQL = SQL & " iScoreDetailId"
    SQL = SQL & " ,iScoreProfileId"
    SQL = SQL & " ,siSerialNo"
    SQL = SQL & " ,fDetailScore"
    SQL = SQL & " ,dtCreate"
    SQL = SQL & " ,dtUpdate"
    SQL = SQL & " ,dtBakDate)"
    SQL = SQL & " SELECT  "
    SQL = SQL & "  sd.iScoreDetailId"
    SQL = SQL & " ,sd.iScoreProfileId"
    SQL = SQL & " ,sd.siSerialNo"
    SQL = SQL & " ,sd.fDetailScore"
    SQL = SQL & " ,sd.dtCreate,sd.dtUpdate,GETDATE()"
    SQL = SQL & " FROM tbSTEScoreDetail sd ,tbSTEScoreProfile sp"
    SQL = SQL & " ,tbSTESubjectProfile sub,tbSTEExamineeProfile ep "
    SQL = SQL & " WHERE sp.iExamineeProfileId =ep.iExamineeProfileId "
    SQL = SQL & " AND sp.iSubjectProfileId =sub.iSubjectProfileId"
    SQL = SQL & " AND sd.iScoreProfileId =sp.iScoreProfileId"
    SQL = SQL & " AND ep.iNendo=" & strNendo
    SQL = SQL & " AND sub.vSubjectName='" & strSubject & "'"
    SQL = SQL & " AND sub.iExamType=2"
    g_obj_Conn.Execute (SQL)
    
    '3.配点割合
    SQL = ""
    SQL = SQL & " UPDATE tbSTEScoreProfile "
    SQL = SQL & " SET  fRawScore=sp.fRawScore * " & strWariai
    SQL = SQL & " ,dtUpdate=GETDATE()"
    SQL = SQL & " FROM   tbSTEScoreProfile sp,tbSTESubjectProfile sub, tbSTEExamineeProfile ep "
    SQL = SQL & " WHERE sp.iExamineeProfileId =ep.iExamineeProfileId "
    SQL = SQL & " AND sp.iSubjectProfileId =sub.iSubjectProfileId"
    SQL = SQL & " AND ep.iNendo=" & strNendo
    SQL = SQL & " AND sub.vSubjectName='" & strSubject & "'"
    SQL = SQL & " AND sub.iExamType=2"
    g_obj_Conn.Execute (SQL)
    
    SQL = ""
    SQL = SQL & " UPDATE tbSTEScoreDetail "
    SQL = SQL & " SET  fDetailScore=sd.fDetailScore * " & strWariai
    SQL = SQL & " ,dtUpdate=GETDATE()"
    SQL = SQL & " FROM tbSTEScoreDetail sd ,tbSTEScoreProfile sp"
    SQL = SQL & " ,tbSTESubjectProfile sub,tbSTEExamineeProfile ep "
    SQL = SQL & " WHERE sp.iExamineeProfileId =ep.iExamineeProfileId "
    SQL = SQL & " AND sp.iSubjectProfileId =sub.iSubjectProfileId"
    SQL = SQL & " AND sd.iScoreProfileId =sp.iScoreProfileId"
    SQL = SQL & " AND ep.iNendo=" & strNendo
    SQL = SQL & " AND sub.vSubjectName='" & strSubject & "'"
    SQL = SQL & " AND sub.iExamType=2"
    g_obj_Conn.Execute (SQL)
    
    '4.終了処理
    SQL = "Update tbSTELocks set iLocks = 0 where vTarget = 'tbSTEScoreProfile' "
    Call g_obj_Conn.Execute(SQL)
    
    g_obj_Conn.CommitTrans
'    g_obj_Conn.RollbackTrans

    blnCommit = False
    MsgBox "配点割合をしました。"

    Exit Sub


ERR_HANDLE:
    If blnCommit = True Then
        g_obj_Conn.RollbackTrans
    End If
    MsgBox Err.Description

End Sub

Private Sub Form_Load()

    On Error GoTo ERR_HANDLE

    Dim oRs  As ADODB.Recordset
    Dim sSQL As String


    txtNendo.Text = g_int_CurrentNendo

    sSQL = "SELECT iSubjectProfileId,vSubjectName "
    sSQL = sSQL & " FROM tbSTESubjectProfile"
    sSQL = sSQL & " WHERE iExamType = 2"
    sSQL = sSQL & " ORDER BY iSubjectProfileId"

    Set oRs = g_obj_Conn.Execute(sSQL)

    Do While Not oRs.EOF
        cboSubject.AddItem oRs("vSubjectName")
        cboSubject.ItemData(cboSubject.NewIndex) = oRs("iSubjectProfileId")
        oRs.MoveNext
    Loop

    oRs.Close

    Exit Sub

ERR_HANDLE:
    MsgBox Err.Description

End Sub
