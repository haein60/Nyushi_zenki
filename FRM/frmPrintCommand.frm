VERSION 5.00
Begin VB.Form frmPrintCommand 
   Caption         =   "frmPrintCommand : 印刷指示"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12720
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmPrintCommand.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   12720
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdPrint 
      Caption         =   "印刷"
      Height          =   615
      Left            =   6600
      TabIndex        =   31
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrintCommand 
      Height          =   360
      Left            =   960
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
   End
   Begin VB.ComboBox cmbPara 
      Height          =   360
      Index           =   9
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   28
      Top             =   7440
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.ComboBox cmbPara 
      Height          =   360
      Index           =   8
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   25
      Top             =   6840
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.ComboBox cmbPara 
      Height          =   360
      Index           =   7
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.ComboBox cmbPara 
      Height          =   360
      Index           =   6
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.ComboBox cmbPara 
      Height          =   360
      Index           =   5
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.ComboBox cmbPara 
      Height          =   360
      Index           =   4
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.ComboBox cmbPara 
      Height          =   360
      Index           =   3
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.ComboBox cmbPara 
      Height          =   360
      Index           =   2
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.ComboBox cmbPara 
      Height          =   360
      Index           =   1
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.ComboBox cmbPara 
      Height          =   360
      Index           =   0
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   9
      Left            =   3840
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   7440
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   8
      Left            =   3840
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   7
      Left            =   3840
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   6240
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   6
      Left            =   3840
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   5640
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   5
      Left            =   3840
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   5040
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   4
      Left            =   3840
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.TextBox txtPara 
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "12345678901234567890"
      Height          =   255
      Index           =   9
      Left            =   960
      TabIndex        =   30
      Top             =   7500
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "12345678901234567890"
      Height          =   255
      Index           =   8
      Left            =   960
      TabIndex        =   27
      Top             =   6900
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "12345678901234567890"
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   24
      Top             =   6300
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "12345678901234567890"
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   21
      Top             =   5700
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "12345678901234567890"
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   18
      Top             =   5100
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "12345678901234567890"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   15
      Top             =   4500
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "12345678901234567890"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   12
      Top             =   3900
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "12345678901234567890"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   9
      Top             =   3300
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "12345678901234567890"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   2700
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  '透明
      Caption         =   "12345678901234567890"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   2100
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmPrintCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type uParaData_Type
    sID         As String
    sDispData   As String
End Type

Private Type uParaList_Type
    lPrintCommandPrameterID     As Long
    sPrintCommandPrameterName   As String
    sDispName                   As String
    iDataType                   As Integer
    iInputType                  As Integer
    sDefaultValue               As String
    sDataFormat                 As String
    siMandatory                 As Integer
    siPrintParameter            As Integer
    iDataSource                 As Integer
    sODBCDataSourceName         As String
    sODBCUserName               As String
    sODBCPassword               As String
    sODBCSQL                    As String
    sODBCSQLID                  As String
    sODBCSQLDESC                As String
    uParaData_()                As uParaData_Type
End Type

Private Type uPrintCommand_Type
    lPrintCommandProfileID  As Long
    sPrintCommandName       As String
    sPrintControlFileName   As String
    sTitle                  As String
    lPrinterID              As Long
End Type

Private prvuParaList_() As uParaList_Type
Private prvuPrintCommand_() As uPrintCommand_Type

Private prvStartCmd As String

Private Sub Form_Load()

    Dim iRtn As Integer
    Dim sCommand As String
    Dim sErrMsg As String

    Dim i     As Long
    Dim llCnt As Long
    Dim luCnt As Long

    Me.Caption = "frmPrintCommand : 印刷指示"

    sCommand = prvStartCmd

    iRtn = lfMakePrintCommandList(sCommand, prvuPrintCommand_, sErrMsg)

    If iRtn <> 0 Then
        MsgBox sErrMsg, vbOKOnly, "印刷コマンドリスト取得失敗"
        Unload Me
        Exit Sub
    End If

    'コンボボックスにリスト展開
    llCnt = LBound(prvuPrintCommand_)
    luCnt = UBound(prvuPrintCommand_)

    For i = llCnt To luCnt
        Me.cmbPrintCommand.AddItem prvuPrintCommand_(i).sPrintCommandName
    Next

    Me.cmbPrintCommand.ListIndex = 0

'    fMainForm.mnuPrint.Enabled = True
'    fMainForm.Toolbar1.Buttons("Print").Enabled = True

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim i As Integer

    fMainForm.mnuTools.Enabled = False  ' disable tools menu

    For i = 1 To fMainForm.Toolbar1.Buttons.Count
        fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next i

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Function lfMakePrintCommandList(psGamenID As String, puPrintCommand_() As uPrintCommand_Type, psErrMsg As String) As Integer

    On Error GoTo ErrProc

    Dim oRs As ADODB.Recordset
    Dim sSQL As String
    Dim oCn As New ADODB.Connection
    Dim lRtn As Long
    Dim lCnt As Long


    lfMakePrintCommandList = -99

    sSQL = ""
    sSQL = sSQL & " SELECT "
    sSQL = sSQL & "        pcd.iPrintCommandProfileID "
    sSQL = sSQL & "      , pcp.vPrintCommandName  "
    sSQL = sSQL & "      , pcp.vPrintControlFileName  "
    sSQL = sSQL & "      , isNull( pcp.vTitle , pcp.vPrintCommandName ) vTitle  "
    sSQL = sSQL & "      , pcd.iPrinterID  "
    sSQL = sSQL & " FROM   tbSTRPrintCommandByDisp as pcd "
    sSQL = sSQL & "   inner join tbSTRPrintCommandProfile as pcp on pcp.iPrintCommandProfileID = pcd.iPrintCommandProfileID "
    sSQL = sSQL & "   ,    tbSTESystemProfile as sp "

'画面に表示するメニューはシステムのフェーズにより切り替える
'    sSQL = sSQL & " WHERE  pcd.iDispID = " & psGamenID

    sSQL = sSQL & " WHERE  pcd.iDispID = sp.iCurrentPhase "
    sSQL = sSQL & " AND    sp.iActiveFlag = 1 "
    sSQL = sSQL & " AND pcp.siUserLevel <= " & Trim(str(glUserLevel))
    sSQL = sSQL & " ORDER BY pcd.iDispOrder "

    lRtn = mdlADODB.pf_ExecSQL(g_obj_Conn, oRs, sSQL, psErrMsg, False, "etc\StMari.INI")

    If lRtn <> 0 Then
        lfMakePrintCommandList = lRtn
        If lRtn = -1 Then
            psErrMsg = "データがありませんでした。"
        End If
        Exit Function
    End If

    lCnt = 0

    Erase puPrintCommand_
    Do Until oRs.EOF
        ReDim Preserve puPrintCommand_(lCnt)
        puPrintCommand_(lCnt).lPrintCommandProfileID = oRs.Fields("iPrintCommandProfileID")
        puPrintCommand_(lCnt).sPrintCommandName = oRs.Fields("vPrintCommandName")
        puPrintCommand_(lCnt).sPrintControlFileName = oRs.Fields("vPrintControlFileName")
        puPrintCommand_(lCnt).sTitle = oRs.Fields("vTitle")
        puPrintCommand_(lCnt).lPrinterID = oRs.Fields("iPrinterID")
        oRs.MoveNext
        lCnt = lCnt + 1
    Loop

    oRs.Close
    Set oRs = Nothing
'    oCn.Close

    lfMakePrintCommandList = 0

Exit Function

ErrProc:

End Function

Private Function lfMakeParaList(plPrintCommandProfileID As Long, puParaList_() As uParaList_Type, psErrMsg As String) As Integer

Dim sSQL As String
Dim oCn As New ADODB.Connection
Dim oRs As ADODB.Recordset
Dim oRs2 As ADODB.Recordset
Dim lRtn As Long
Dim lCnt As Long
Dim lCnt2 As Long

Dim sIDColName As String
Dim sDataColName As String

On Error GoTo ErrProc

    lfMakeParaList = -99

    sSQL = ""
    sSQL = sSQL & " SELECT "
    sSQL = sSQL & "        pcpr.iPrintCommandPrameterID "
    sSQL = sSQL & "      , pcpr.vPrintCommandPrameterName  "
    sSQL = sSQL & "      , pcpr.vDispName  "
    sSQL = sSQL & "      , pcpr.iDataType  "
    sSQL = sSQL & "      , pcpr.iInputType  "
    sSQL = sSQL & "      , pcpr.vDefaultValue  "
    sSQL = sSQL & "      , pcpr.vDataFormat  "
    sSQL = sSQL & "      , pcpr.iDataSource  "
    sSQL = sSQL & "      , pcpr.vODBCDataSourceName  "
    sSQL = sSQL & "      , pcpr.vODBCUserName  "
    sSQL = sSQL & "      , pcpr.vODBCPassword  "
    sSQL = sSQL & "      , pcpr.vODBCSQL  "
    sSQL = sSQL & "      , pcpr.vODBCSQLID  "
    sSQL = sSQL & "      , pcpr.vODBCSQLDESC  "
    sSQL = sSQL & "      , isnull( pcpr.siMandatory , 0 ) as siMandatory  "
    sSQL = sSQL & "      , isnull( pcpr.siPrintParameter , 0 ) as siPrintParameter  "
    sSQL = sSQL & " FROM   tbSTRPrintCommandPrameter as pcpr "
    sSQL = sSQL & " WHERE  pcpr.iPrintCommandProfileID = " & str(plPrintCommandProfileID)
    sSQL = sSQL & " ORDER BY pcpr.iDispOrder "

    lRtn = mdlADODB.pf_ExecSQL(g_obj_Conn, oRs, sSQL, psErrMsg, False, "etc\StMari.INI")

    If lRtn <> 0 Then
        lfMakeParaList = lRtn
        If lRtn = -1 Then
            psErrMsg = "データがありませんでした。"
        End If
        Exit Function
    End If

    lCnt = 0

    Erase puParaList_
    Do Until oRs.EOF
        ReDim Preserve puParaList_(lCnt)
        puParaList_(lCnt).lPrintCommandPrameterID = oRs.Fields("iPrintCommandPrameterID")
        puParaList_(lCnt).sPrintCommandPrameterName = gfNullChkStrTrim(oRs.Fields("vPrintCommandPrameterName"))
        puParaList_(lCnt).sDispName = gfNullChkStrTrim(oRs.Fields("vDispName"))
        puParaList_(lCnt).iDataType = gfNullZeroChkInt(oRs.Fields("iDataType"))
        puParaList_(lCnt).iInputType = gfNullZeroChkInt(oRs.Fields("iInputType"))
        puParaList_(lCnt).sDefaultValue = gfNullChkStrTrim(oRs.Fields("vDefaultValue"))
        puParaList_(lCnt).sDataFormat = gfNullChkStrTrim(oRs.Fields("vDataFormat"))
        puParaList_(lCnt).siMandatory = gfNullZeroChkInt(oRs.Fields("siMandatory"))
        puParaList_(lCnt).siPrintParameter = gfNullZeroChkInt(oRs.Fields("siPrintParameter"))
        puParaList_(lCnt).iDataSource = gfNullZeroChkInt(oRs.Fields("iDataSource"))
        puParaList_(lCnt).sODBCDataSourceName = gfNullChkStrTrim(oRs.Fields("vODBCDataSourceName"))
        puParaList_(lCnt).sODBCUserName = gfNullChkStrTrim(oRs.Fields("vODBCUserName"))
        puParaList_(lCnt).sODBCPassword = gfNullChkStrTrim(oRs.Fields("vODBCPassword"))
        puParaList_(lCnt).sODBCSQL = gfNullChkStrTrim(oRs.Fields("vODBCSQL"))
        puParaList_(lCnt).sODBCSQLID = gfNullChkStrTrim(oRs.Fields("vODBCSQLID"))
        puParaList_(lCnt).sODBCSQLDESC = gfNullChkStrTrim(oRs.Fields("vODBCSQLDESC"))
        puParaList_(lCnt).sDataFormat = gfNullChkStrTrim(oRs.Fields("vDataFormat"))
        If puParaList_(lCnt).iInputType = "1" Then '入力がリスト選択のとき
            puParaList_(lCnt).iDataSource = gfNullZeroChkInt(oRs.Fields("iDataSource"))
            puParaList_(lCnt).sODBCDataSourceName = gfNullChkStrTrim(oRs.Fields("vODBCDataSourceName"))
            puParaList_(lCnt).sODBCUserName = gfNullChkStrTrim(oRs.Fields("vODBCUserName"))
            puParaList_(lCnt).sODBCPassword = gfNullChkStrTrim(oRs.Fields("vODBCPassword"))
            If puParaList_(lCnt).iDataSource = 0 Then '固定リストから取得
                sSQL = ""
                sSQL = sSQL & "SELECT "
                sSQL = sSQL & "   vDataID"
                sSQL = sSQL & " , vDispData"
                sSQL = sSQL & "  FROM tbSTRPrintCommandParaData "
                sSQL = sSQL & " WHERE iPrintCommandPrameterID = " & puParaList_(lCnt).lPrintCommandPrameterID
                sSQL = sSQL & " ORDER BY iDispOrder"
                sIDColName = "vDataID"
                sDataColName = "vDispData"
            Else '固有のＳＱＬにより取得
                sSQL = puParaList_(lCnt).sODBCSQL
                sIDColName = puParaList_(lCnt).sODBCSQLID
                sDataColName = puParaList_(lCnt).sODBCSQLDESC
            End If
            lRtn = mdlADODB.pf_ExecSQL(g_obj_Conn, oRs2, sSQL, psErrMsg, False, "etc\StMari.INI")
            If lRtn <> 0 Then
                lfMakeParaList = lRtn
                If lRtn = -1 Then
                    psErrMsg = "データがありませんでした。"
                End If
                Exit Function
            End If
            lCnt2 = 0
            Erase puParaList_(lCnt).uParaData_
            Do Until oRs2.EOF
                ReDim Preserve puParaList_(lCnt).uParaData_(lCnt2)
                puParaList_(lCnt).uParaData_(lCnt2).sID = gfNullChkStrTrim(oRs2.Fields(sIDColName))
                puParaList_(lCnt).uParaData_(lCnt2).sDispData = gfNullChkStrTrim(oRs2.Fields(sDataColName))
                oRs2.MoveNext
                lCnt2 = lCnt2 + 1
            Loop
        End If
        oRs.MoveNext
        lCnt = lCnt + 1
    Loop

    oRs.Close
    Set oRs = Nothing
'    oCn.Close

    lfMakeParaList = 0

Exit Function

ErrProc:

End Function

Private Function chkMandatory() As Long

Dim lCnt As Long
Dim llCnt As Long
Dim luCnt As Long
Dim sWk As String

On Error GoTo ErrProc

    chkMandatory = -99

    llCnt = LBound(prvuParaList_)
    luCnt = UBound(prvuParaList_)
    For lCnt = llCnt To luCnt
        If prvuParaList_(lCnt).siMandatory = 1 Then
            If prvuParaList_(lCnt).iInputType = 1 Then
                'リストに登録
                If cmbPara(lCnt).ListIndex < 0 Then
                    chkMandatory = lCnt
                    Exit Function
                End If
            Else
                sWk = Trim(txtPara(lCnt).Text)
                
                If sWk = "" Then
                    chkMandatory = lCnt
                    Exit Function
                End If
                Select Case prvuParaList_(lCnt).iDataType
                Case 1  'Int
                    If Not gf_IntCheck(sWk) Then
                        chkMandatory = lCnt
                        Exit Function
                    End If
                Case 2  'Double
                    If Not gf_DblCheck(sWk) Then
                        chkMandatory = lCnt
                        Exit Function
                    End If
                Case 3  'datetime
                    If Not IsDate(sWk) Then
                        chkMandatory = lCnt
                        Exit Function
                    End If
                End Select
            End If
        End If
    Next

chkMandatory = -1

Exit Function

ErrProc:

End Function

Private Sub cmbPrintCommand_Click()

    Dim lIndex As Long
    
    Dim iRtn As Integer
    Dim sErrMsg As String
    
    Dim lCnt As Long
    Dim llCnt As Long
    Dim luCnt As Long
    Dim lCnt2 As Long
    Dim lLCnt2 As Long
    Dim lUCnt2 As Long

    For lCnt = 0 To 9
        Me.cmbPara(lCnt).Clear
        Me.cmbPara(lCnt).Visible = False
        Me.txtPara(lCnt).Visible = False
        Me.lblPara(lCnt).Visible = False
    Next

    'パラメータの入力を促す
    lIndex = Me.cmbPrintCommand.ListIndex

    iRtn = lfMakeParaList(prvuPrintCommand_(lIndex).lPrintCommandProfileID, prvuParaList_, sErrMsg)

    If iRtn <> 0 Then
        If iRtn <> -1 Then
            MsgBox sErrMsg, vbOKOnly, "印刷コマンドパラメータ取得失敗"
        End If
        Exit Sub
    End If

    '画面表示
    llCnt = LBound(prvuParaList_)
    luCnt = UBound(prvuParaList_)

    For lCnt = llCnt To luCnt
        Me.lblPara(lCnt).Visible = True
        Me.lblPara(lCnt).Caption = prvuParaList_(lCnt).sDispName & IIf(prvuParaList_(lCnt).siMandatory = 1, " (*)", "")
        If prvuParaList_(lCnt).iInputType = 1 Then
            Me.cmbPara(lCnt).Visible = True
            'リストに登録
            lLCnt2 = LBound(prvuParaList_(lCnt).uParaData_)
            lUCnt2 = UBound(prvuParaList_(lCnt).uParaData_)
            For lCnt2 = lLCnt2 To lUCnt2
                Me.cmbPara(lCnt).AddItem prvuParaList_(lCnt).uParaData_(lCnt2).sDispData
            Next
            Me.cmbPara(lCnt).ListIndex = 0
        Else
            Me.txtPara(lCnt).Visible = True
            If prvuParaList_(lCnt).sDefaultValue = "<%NENDO%>" Then
                Me.txtPara(lCnt).Text = Trim(str(g_int_CurrentNendo))
            Else
                Me.txtPara(lCnt).Text = prvuParaList_(lCnt).sDefaultValue
            End If
        End If
    Next

End Sub

Public Sub cmdPrint_Click()

    Dim oCn As New ADODB.Connection
    Dim oRs As ADODB.Recordset
    
    Dim sSQL As String
    Dim sParam As String
    Dim sErrMsg As String
    Dim lRtn As Long
    
    Dim lCnt As Long
    Dim llCnt As Long
    Dim luCnt As Long
    Dim sTitle As String
    Dim sSubTitle As String

    lRtn = chkMandatory

    If lRtn >= 0 Then
        MsgBox "入力項目にエラーがあります。" & vbCrLf & vbCrLf & lblPara((lRtn)).Caption & vbCrLf & vbCrLf & "必須項目は必ず指定してください。", vbOKOnly, "入力エラー"
        If prvuParaList_(lRtn).iInputType = 1 Then
            If cmbPara(lRtn).Enabled Then
                cmbPara(lRtn).SetFocus
            End If
        Else
            If txtPara(lRtn).Enabled Then
                txtPara(lRtn).SetFocus
            End If
        End If
        Exit Sub
    End If

    If vbNo = MsgBox("印刷要求を実行します。よろしいですか？", vbYesNo, "印刷要求確認") Then Exit Sub

    sSQL = ""
    sSQL = sSQL & "INSERT INTO tbSTRReportData ( "
    sSQL = sSQL & "   iReportId "
    sSQL = sSQL & " , iModuleReportId "
    sSQL = sSQL & " , iPrinterId "
    sSQL = sSQL & " , vParameterString "
    sSQL = sSQL & " ) "
    sSQL = sSQL & "select isnull( max(iReportId) , 0 ) + 1 "
    sSQL = sSQL & " , " & prvuPrintCommand_(Me.cmbPrintCommand.ListIndex).sPrintControlFileName
    sSQL = sSQL & " , " & prvuPrintCommand_(Me.cmbPrintCommand.ListIndex).lPrinterID

    llCnt = LBound(prvuParaList_)
    luCnt = UBound(prvuParaList_)
    sParam = ""

    sTitle = prvuPrintCommand_(Me.cmbPrintCommand.ListIndex).sTitle
    sSubTitle = ""

    For lCnt = llCnt To luCnt

Dim slPara As String

        slPara = prvuParaList_(lCnt).sPrintCommandPrameterName
        If prvuParaList_(lCnt).iInputType = 1 Then

            'リストに登録
            Select Case prvuParaList_(lCnt).iDataType
            Case 0  'String
                slPara = slPara & "=""" & prvuParaList_(lCnt).uParaData_(Me.cmbPara(lCnt).ListIndex).sID & """"
            Case 1  'Int
                If prvuParaList_(lCnt).siPrintParameter = 1 Then
'                If slPara = "vOrderField" Or slPara = "iSubType" Or slPara = "iSubjectProfileId_Disp" Or slPara = "iSecondExamDayFlag" Then
                    If sSubTitle <> "" Then
                        sSubTitle = sSubTitle & ":" & prvuParaList_(lCnt).uParaData_(Me.cmbPara(lCnt).ListIndex).sDispData
                    Else
                        sSubTitle = prvuParaList_(lCnt).uParaData_(Me.cmbPara(lCnt).ListIndex).sDispData
                    End If
                End If
                slPara = slPara & "=" & prvuParaList_(lCnt).uParaData_(Me.cmbPara(lCnt).ListIndex).sID
'                If slPara = "vOrderField" And prvuParaList_(lCnt).uParaData_(Me.cmbPara(lCnt).ListIndex).sID = "iJukenNumber" Then
'                    slPara = slPara & "=" & "1"
'                Else
'                    slPara = slPara & "=" & prvuParaList_(lCnt).uParaData_(Me.cmbPara(lCnt).ListIndex).sID
'                End If
            Case 2  'Double
                slPara = slPara & "=" & prvuParaList_(lCnt).uParaData_(Me.cmbPara(lCnt).ListIndex).sID
            Case 3  'datetime
                slPara = slPara & "=""" & prvuParaList_(lCnt).uParaData_(Me.cmbPara(lCnt).ListIndex).sID & """"
            End Select
        Else
            Select Case prvuParaList_(lCnt).iDataType
            Case 0  'String
                slPara = slPara & "=""" & Me.txtPara(lCnt).Text & """"
            Case 1  'Int
                slPara = slPara & "=" & Me.txtPara(lCnt).Text
            Case 2  'Double
                slPara = slPara & "=" & Me.txtPara(lCnt).Text
            Case 3  'datetime
                slPara = slPara & "=""" & Me.txtPara(lCnt).Text & """"
            End Select
        End If
        sParam = sParam & ";" & slPara
    Next

    If sSubTitle <> "" Then sTitle = sTitle & "(" & sSubTitle & ")"
    sParam = sParam & ";vTitle=""" & sTitle & """"
    sSQL = sSQL & " , '" & sParam & "' "
    sSQL = sSQL & " FROM tbSTRReportData "

    lRtn = mdlADODB.pf_ExecSQL_NoRtn(g_obj_Conn, oRs, sSQL, sErrMsg, False, "etc\StMari.INI")

    If lRtn <> 0 Then
        If lRtn = -1 Then
            sErrMsg = "データ登録に失敗しました。"
        End If
        MsgBox sErrMsg, vbOKOnly, "実行エラー"
        Exit Sub
    End If

    MsgBox "登録が完了しました。", vbOKOnly, "登録完了"

End Sub

Private Sub txtPara_GotFocus(Index As Integer)

    txtPara(Index).SelStart = 0
    txtPara(Index).SelLength = Len(txtPara(Index).Text)

End Sub
