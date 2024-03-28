VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H80000004&
   Caption         =   "入試システム"
   ClientHeight    =   8535
   ClientLeft      =   2280
   ClientTop       =   1740
   ClientWidth     =   12495
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0000
   Tag             =   "1905"
   WindowState     =   2  '最大化
   Begin VB.PictureBox pctExplorer 
      Align           =   3  '左揃え
      Height          =   8115
      Left            =   0
      ScaleHeight     =   8055
      ScaleWidth      =   3645
      TabIndex        =   1
      Top             =   420
      Width           =   3705
      Begin MSComctlLib.TreeView tvwMenu 
         Height          =   7815
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   13785
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '上揃え
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Clear"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4035
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":415D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F43
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuExamKubun 
      Caption         =   "試験区分"
      Begin VB.Menu mnuExamZenki 
         Caption         =   "前期試験"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "メニュー"
      Begin VB.Menu mnuApplyPhase 
         Caption         =   "願書受付フェーズ"
         Begin VB.Menu mnuOCR 
            Caption         =   "Web出願データ取込"
         End
         Begin VB.Menu mnuMaintainExamineeData 
            Caption         =   "受験者データの編集"
         End
         Begin VB.Menu mnuExamineeCheck 
            Caption         =   "受験者情報メンテナンス"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFixData1 
            Caption         =   "データ確定"
         End
      End
      Begin VB.Menu mnu1stExam 
         Caption         =   "一次試験"
         Begin VB.Menu mnuRoomAllocation 
            Caption         =   "会場入力"
         End
         Begin VB.Menu mnuInputAbsenteeRecord 
            Caption         =   "欠席者入力"
         End
         Begin VB.Menu mnuInputRawScore 
            Caption         =   "素点入力"
         End
         Begin VB.Menu mnuInputChooseiScore2 
            Caption         =   "条件別調整点入力"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuInputChooseiScore 
            Caption         =   "科目別調整点入力"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuInputPassedPersonData 
            Caption         =   "合格者入力"
         End
         Begin VB.Menu mnuPreparationDay 
            Caption         =   "試験日振分"
         End
         Begin VB.Menu mnuManualAllocation 
            Caption         =   "試験日変更"
         End
         Begin VB.Menu mnuFixData2 
            Caption         =   "データ確定"
         End
         Begin VB.Menu mnuMaintainExamineeData2 
            Caption         =   "受験者データの編集"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu2ndExam 
         Caption         =   "ニ次試験"
         Begin VB.Menu mnuInputAbsenteeRecord2 
            Caption         =   "欠席者入力"
         End
         Begin VB.Menu mnuTeacherRoomMapInterview 
            Caption         =   "面接委員登録"
         End
         Begin VB.Menu mnuPreparationRoom 
            Caption         =   "面接グループ振分"
         End
         Begin VB.Menu mnuManualAllocationGrp 
            Caption         =   "面接グループ変更"
         End
         Begin VB.Menu mnuTeacherRoomMapReport 
            Caption         =   "小論文採点委員登録"
         End
         Begin VB.Menu mnuPreparationReport 
            Caption         =   "小論文振分"
         End
         Begin VB.Menu mnuImport_Syoronbun 
            Caption         =   "素点入力(小論文)_import"
         End
         Begin VB.Menu mnuInputRawScoreI 
            Caption         =   "素点入力(小論文)"
         End
         Begin VB.Menu mnuImport_Mensetu 
            Caption         =   "素点入力(面接)_import"
         End
         Begin VB.Menu mnuInputRawScore2 
            Caption         =   "素点入力(面接)"
         End
         Begin VB.Menu mnuInputPassedPersonData2 
            Caption         =   "合格者入力"
         End
         Begin VB.Menu mnuWaitList2 
            Caption         =   "補欠者入力"
         End
         Begin VB.Menu mnuHoketusyaJuni 
            Caption         =   "補欠者順位"
         End
         Begin VB.Menu mnuFixData3 
            Caption         =   "データ確定"
         End
         Begin VB.Menu mnuAdjustScoreM 
            Caption         =   "調整点入力(面接)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAdjustScoreS 
            Caption         =   "調整点入力(小論文)"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuEnterRefuse 
         Caption         =   "入学手続き処理"
         Begin VB.Menu mnuUpliftment 
            Caption         =   "補欠者合格者繰上げ処理"
         End
         Begin VB.Menu mnuRefuseOffer 
            Caption         =   "辞退"
         End
         Begin VB.Menu mnuFixData4 
            Caption         =   "データ確定"
         End
      End
      Begin VB.Menu mnuMaster 
         Caption         =   "マスターメンテナンス"
         Begin VB.Menu mnuRoomProfile 
            Caption         =   "会場・面接グループ"
         End
         Begin VB.Menu mnuInterviewerProfile 
            Caption         =   "採点者プロファイル"
         End
         Begin VB.Menu mnuInterviewGroupProfile 
            Caption         =   "所属プロフィール"
         End
         Begin VB.Menu mnuSystemData 
            Caption         =   "入試年度設定"
         End
      End
      Begin VB.Menu mnuPrintMenu 
         Caption         =   "印刷"
         Begin VB.Menu mnuPrintCommand 
            Caption         =   "印刷指示"
         End
         Begin VB.Menu mnuExcelReport 
            Caption         =   "Excel帳票"
         End
         Begin VB.Menu mnuPrintDosu 
            Caption         =   "度数分布図印刷"
         End
      End
      Begin VB.Menu mnuTransfer 
         Caption         =   "受験データCSV出力"
         Begin VB.Menu mnuOutputCSV 
            Caption         =   "受験生＋素点情報"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "ツール"
      Visible         =   0   'False
      Begin VB.Menu mnuToolsSearch 
         Caption         =   "レコードを表示"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuToolsSave 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuToolsDelete 
         Caption         =   "削除"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuToolsCancel 
         Caption         =   "キャンセル"
      End
      Begin VB.Menu mnuToolsNew 
         Caption         =   "新規"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuToolsQuery 
         Caption         =   "クエリ"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTreeMenu 
      Caption         =   "ツリーメニュー"
      Begin VB.Menu mnuShowTree 
         Caption         =   "メニュー表示"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "印刷"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "ウインドウ"
      Visible         =   0   'False
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuTileHorizontally 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile Vertically"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ヘルプ"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuExit 
      Caption         =   "終了"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*************************************************************************************************
'Form Name      :   frmMain
'Author         :   Dileep Cherian
'Created On     :
'Description    :   This form is the MDI form for the module.
'Reference      :   FunctionalSpecs Of MasterMaintenance.doc ver 1.0
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History - Mahesh Deshpande    -   05/04/2002
'Caption of master maintenance forms should display the mode in which they are at any time
'ie; Edit, Query or New Mode
'**************************************************************************************************

''''Public f_int_CurrentPhase  As Integer       'modNyushi.basに移動 2021.12.28 del jhi

Public frmChooseiSuisen2   As Form              ' choosei score for second phase
Public frmIntwrRoomMapInt  As Form              ' Teacher-Room Mapping for interview
Public frmIntwrRoomMapRpt  As Form              ' Teacher-Room Mapping for Report
Public frmRawScoreInt      As Form              ' Raw score for interview
Public frmRawScoreRpt      As Form              ' Raw score for report
Public frmChooseiGrace     As Form              ' choosei score for 1st phase
Public frmChooseiHyotei    As Form              ' choosei score for Hyotei
Public frmAbsenteeRecord   As Form              ' absentee record
Public frmPassedPersonData As Form              ' passed person data
Public frmWaitingList      As Form              ' waiting list
Public frmUpliftment       As Form              ' upliftment from waiting list
Public frmRefuseOffer      As Form              ' enter/refuse offer

Private Const prvsProfileName As String = "Passcheck"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

''''2021.12.28 del jhi globalに宣言する
''''Private Type prvuMenues_Type
''''    oMnuObj As Object
''''    sTVKey As String
''''    sIniKey As String
''''    sCaption As String
''''    lParent As Long
''''    bVisible As Boolean
''''End Type
''''
''''Dim uMenues_() As prvuMenues_Type



Private Sub MDIForm_Load()
    
    On Error GoTo ErrorHandler

    'For toolbar and image list  used in procedure initToolbar
    Dim l_bln_Conn            As Boolean                       ' to check the status of database connection
    Dim l_str_sqlCurrentPhase As String                        ' to get the curretn phase
    Dim l_obj_rsCurrentPhase  As New ADODB.Recordset           ' to get the curretn phase
    Dim l_obj_rsNendo         As New ADODB.Recordset           ' to get the current year




    gbExamCheckNewShow = True

    ' get the current year into global variable
    l_obj_rsNendo.Open "SELECT iNendo FROM tbSTESystemProfile WHERE iActiveFlag=1", g_obj_Conn

    If Not l_obj_rsNendo.EOF Then
        g_int_CurrentNendo = l_obj_rsNendo.Fields("iNendo").Value
    Else
        ' no active year set in the system profile table - so end the apllication
        g_int_CurrentNendo = 0
''''    MsgBox LoadResString(2481), vbCritical, LoadResString(1905) ''''2022.01.29 del jhi
        MsgBox "アプリケーションの初期化に失敗しました。しばらく後、再実行してください。", vbCritical, gTit
        Call mnuExit_Click
    End If

    l_obj_rsNendo.Close
    Set l_obj_rsNendo = Nothing

''''2021.12.22 del jhi
''''Call g_void_SetFontProperties(Me)       ' set the font properties


    Call InitToolbar                        'initilize the toolbar

    ' get the current phase
    l_str_sqlCurrentPhase = "SELECT iCurrentPhase FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag=1" & _
        " AND iCurrentPhase IS NOT NULL"

    l_obj_rsCurrentPhase.Open l_str_sqlCurrentPhase, g_obj_Conn

    If Not l_obj_rsCurrentPhase.EOF Then
         f_int_CurrentPhase = l_obj_rsCurrentPhase.Fields("iCurrentPhase").Value
    Else
        ' exit if failed to initialize
''''    MsgBox LoadResString(2481), vbCritical, LoadResString(1905) ''''2021.12.08 del jhi

''''2021.12.08 add jhi
        MsgBox "アプリケーションの初期化に失敗しました。しばらく後、再実行してください。", vbCritical, gTit

        Call mnuExit_Click
    End If

    l_obj_rsCurrentPhase.Close
    Set l_obj_rsCurrentPhase = Nothing

'
    Call SetPhaseMenu(CLng(f_int_CurrentPhase))

    LoadResStrings Me
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)


''''----------------------------------------------------------------------------
''''Init_TreeView_New()に統合 S
''''2021.12.28 del jhi
''''----------------------------------------------------------------------------

''''treeviewに背景色を設定
''''Call SetTVBackColor(tvwMenu, RGB(230, 230, 250)) 'Lavender
''''Call SetTVBackColor(tvwMenu, RGB(240, 255, 240)) 'Honeydew
''''Call SetTVBackColor(tvwMenu, RGB(220, 220, 220)) 'Gainsboro
'   Call SetTVBackColor(tvwMenu, RGB(255, 250, 250)) 'Snow


    '---------------------------------------------------------------------------
    'Menuを全部広げる
    '2021.12.22 add jhi
    '---------------------------------------------------------------------------
'    Dim objNode As Node
'    For Each objNode In tvwMenu.Nodes
''''''    If (objNode Is TreeView1.SelectedItem) Then
'            objNode.Expanded = True
''''''    End If
'    Next


''''----------------------------------------------------------------------------
''''Init_TreeView_New()に統合 E
''''2021.12.28 add
''''----------------------------------------------------------------------------

''''----------------------------------------------------------------------------
''''条件付きコンパイル引数の設定 2022.02.01 add jhi
''''----------------------------------------------------------------------------
#If zengo_kubun = 1 Then
    mnuExamZenki.Caption = "前期試験"
#Else
    mnuExamZenki.Caption = "後期試験"
#End If

    Exit Sub


ErrorHandler:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub MDIForm_Activate()

    Dim i As Integer


    fMainForm.mnuTools.Enabled = False

    'New Code added by Mahesh (16/5/02)
    If Forms.Count > 1 Then
        fMainForm.ActiveForm.ZOrder 0
        Exit Sub
    End If
    'New code ends
    
    For i = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next i


    Me.Caption = "frmMain : " & gTit


End Sub

Private Sub MDIForm_Resize()

    On Error GoTo ErrorHandler
            
    With tvwMenu
        .Top = 0
        .Left = 0
''''    .Width = 2895                                    ''''2021.11.30 del jhi
        .Width = 3700                                    ''''2021.11.30 add jhi Tree Menu haba
        .Height = pctExplorer.Height
        .Font.Size = 10                                  ''''2021.12.22 add jhi

    End With


    Exit Sub

ErrorHandler:
    MsgBox Err.Description

End Sub



'*******************************************************************************
'* １．出願受付フェーズ                                                        *
'*******************************************************************************
'*******************************************************************************
'* Web出願データ取込                                                           *
'*******************************************************************************
Private Sub mnuOCR_Click()


    Unload frmBrowse

    frmBrowse.Caption = "frmBrowse : Web出願データ取込"
    frmBrowse.Show

    frmBrowse.ZOrder 0

End Sub

'*******************************************************************************
'* 受験生データの編集                                                          *
'*******************************************************************************
Private Sub mnuMaintainExamineeData_Click()


    gbExamCheckNewShow = True ''''2021.12.22 add jhi


    If gbExamCheckNewShow Then
 
        ''''Unload frmExamCheckPara ''''2021.12.22 add jhi ''''いけない場合があるので2023.01.24 del jhi
        frmExamCheckPara.Caption = "frmExamCheckPara : 受験生データの編集"
        frmExamCheckPara.Show

        ''''コントロールを Z オーダーの最前面に配置します。 コントロールが他のコントロールの上に表示されます (既定値)。
        frmExamCheckPara.ZOrder 0

    Else
        ''''frmExamineeCheck.ZOrder 0 ''''2023.01.24 意味がないのでcomment out
    End If

End Sub

'*******************************************************************************
'* データ確定 処理                                                             *
'*******************************************************************************
Private Sub mnuFixData1_Click()

    On Error GoTo ErrorHandler

    Dim l_frm          As Form
    Dim l_obj_Rst      As New ADODB.Recordset
    Dim rinf           As Integer
    Dim sSQL           As String
    Dim sTmp           As String

    Dim step_no        As Integer


    
step_no = 1

    Select Case f_int_CurrentPhase
    Case 0 '願書受付フェーズ
        sTmp = "願書受付フェーズを確定します。よろしいですか？"
        rinf = MsgBox(sTmp, vbYesNo + vbQuestion, gTit)

    Case 1 '一次試験
        sTmp = "一次試験フェーズを確定します。よろしいですか？"
        rinf = MsgBox(sTmp, vbYesNo + vbQuestion, gTit)

    Case 2 '二次試験
        sTmp = "二次試験フェーズを確定します。よろしいですか？"
        rinf = MsgBox(sTmp, vbYesNo + vbQuestion, gTit)

    Case 3 '入学手続き処理
        sTmp = "入学手続きフェーズを確定します。よろしいですか？"
        rinf = MsgBox(sTmp, vbYesNo + vbQuestion, gTit)

    Case Else
        sTmp = "f_int_CurrentPhase error!"
        rinf = MsgBox(sTmp, vbYesNo + vbQuestion, gTit)

    End Select


    If rinf = vbYes Then

        If f_int_CurrentPhase <> 3 Then
            f_int_CurrentPhase = f_int_CurrentPhase + 1     ' 次のフェーズのフラグをセット
        Else
            '「入学手続き処理」からデータ確定すると「願書受付フェーズ」にセット
            f_int_CurrentPhase = 0
        End If


step_no = 2
        '-----------------------------------------------------------------------
        ' tbSTESystemProfile
        '-----------------------------------------------------------------------
        sSQL = "Update tbSTESystemProfile set iCurrentPhase= " & f_int_CurrentPhase & " where iActiveFlag=1"
        g_obj_Conn.Execute (sSQL)

step_no = 3

        Select Case f_int_CurrentPhase
        Case 0
            sTmp = "入学手続きフェーズを確定します。願書受け付け フェーズに戻ります。"
            MsgBox sTmp, vbInformation, gTit

        Case 1
            sTmp = "願書受け付けフェーズのデータを確定しました。一次試験フェーズを入力してください。"
            MsgBox sTmp, vbInformation, gTit

        Case 2
            sTmp = "一次試験フェーズのデータを確定しました。二次試験フェーズを入力してください。"
            MsgBox sTmp, vbInformation, gTit

        Case 3
            sTmp = "二次試験フェーズのデータを確定しました。入学手続きフェーズを入力してください。"
            MsgBox sTmp, vbInformation, gTit

        Case Else
            sTmp = "入力フェーズの設定フラグ異常です。処理を中断します。"
            MsgBox sTmp, vbInformation, gTit
        End Select
            

        For Each l_frm In Forms
            If l_frm.Name <> "frmMain" Then
                Unload l_frm
            End If
        Next

    End If


''''----------------------------------------------------------------------------
''''2021.12.28 del jhi
''''----------------------------------------------------------------------------
''''tvwMenu.Nodes.Clear
''''Call SetPhaseMenu(CLng(f_int_CurrentPhase))


''''----------------------------------------------------------------------------
''''2021.12.28 add jhi
''''----------------------------------------------------------------------------
    Call Init_TreeView_New(uMenues_)


    Exit Sub


ErrorHandler:

    If (step_no = 2) Then
        sTmp = "処理フェーズのフラグをセットする処理でエラーが発生しました。"
        MsgBox sTmp, vbInformation, gTit
    Else
        MsgBox Err.Number & vbCrLf & Err.Description
    End If

End Sub

'*******************************************************************************
'* ２．１次試験                                                                *
'*******************************************************************************
'*******************************************************************************
'* 会場入力                                                                    *
'*******************************************************************************
Private Sub mnuRoomAllocation_Click()

    Unload frmRoomAlloc

    frmRoomAlloc.Caption = "frmRoomAlloc : 会場入力"
    frmRoomAlloc.Show

    frmRoomAlloc.ZOrder 0

End Sub

'*******************************************************************************
'* 欠席者入力                                                                  *
'*******************************************************************************
Private Sub mnuInputAbsenteeRecord_Click()


''''    If f_int_CurrentPhase <> 1 Then
''''        MsgBox "1フェーズを合わせて画面表示を行ってください。"
''''        Exit Sub
''''    End If


    ' absentee record for the 1st phase

'2021.12.29 del jhi
'    If frmAbsenteeRecord Is Nothing Then
'        Set frmAbsenteeRecord = New frmExamineeStatus
'    Else
'        Unload frmAbsenteeRecord
'    End If

'    With frmAbsenteeRecord
'        .m_int_IntRpt = 1
'        .m_int_Action = 0
'        .Show
'        .Caption = "frmAbsenteeRecord : 欠席者入力"
'        .ZOrder 0
'    End With

'2021.12.29 add jhi
    frm1jikesseki.m_int_IntRpt = 1
    frm1jikesseki.m_int_Action = 0
    frm1jikesseki.Caption = "frm1jikesseki : 1次 欠席者入力"

    frm1jikesseki.Show
    frm1jikesseki.ZOrder 0



End Sub

'*******************************************************************************
'* 1次素点入力                                                                 *
'*******************************************************************************
Private Sub mnuInputRawScore_Click()

    g_int_ExamType = 1 '1次試験フェーズを設定 2021.12.22 add jhi


''''2022.01.24 form変更により del jhi
''''    Unload frmRawScore
''''
''''    frmRawScore.Caption = "frmRawScore : 素点入力"
''''    frmRawScore.Show
''''
''''    frmRawScore.ZOrder 0


''''2022.01.24 form変更により add jhi
    Unload frm1jiSotenInput

    frm1jiSotenInput.Caption = "frm1jiSotenInput : 素点入力"
    frm1jiSotenInput.Show

    frm1jiSotenInput.ZOrder 0


End Sub

'*******************************************************************************
'* 合格者入力                                                                  *
'*******************************************************************************
Private Sub mnuInputPassedPersonData_Click()

    ' passed person data for 1st phase


'2021.12.29 del jhi
'    If frmPassedPersonData Is Nothing Then
'        Set frmPassedPersonData = New frmExamineeStatus
'    End If
'
'    With frmPassedPersonData
'        .m_int_IntRpt = 1
'        .m_int_Action = 1
'        .Show
'        .Caption = "frmPassedPersonData : 合格者入力"
'        .ZOrder 0
'    End With


'2021.12.29 add jhi
    frm1jigoukaku.m_int_Action = 1
    frm1jigoukaku.m_int_IntRpt = 1
    frm1jigoukaku.Caption = "frm1jigoukaku : 1次 合格者入力"
    frm1jigoukaku.Show
    frm1jigoukaku.ZOrder 0


End Sub

'*******************************************************************************
'* 試験日振分                                                                  *
'*******************************************************************************
Private Sub mnuPreparationDay_Click()

    Dim strMsg As String

''''2022.03.09 add jhi 現地で
#If zengo_kubun = 1 Then
    strMsg = "frmPrepSecondExam : 二次試験日振分"
#Else
    strMsg = "frmPrepSecondExam : 二次試験日確定"
#End If



    frmPrepSecondExam.Show
    frmPrepSecondExam.Caption = strMsg
    frmPrepSecondExam.ZOrder 0

End Sub

'*******************************************************************************
'* 試験日変更                                                                  *
'*******************************************************************************
Private Sub mnuManualAllocation_Click()

    frmManualAllocation.Show
    frmManualAllocation.Caption = "frmManualAllocation : 二次試験日変更"
    frmManualAllocation.ZOrder 0

End Sub

'*******************************************************************************
'* データ確定                                                                  *
'*******************************************************************************
Private Sub mnuFixData2_Click()

    Call mnuFixData1_Click

End Sub

'*******************************************************************************
'* ３．２次試験                                                                *
'*******************************************************************************
'*******************************************************************************
'* 2次試験 : 欠席者入力                                                        *
'*******************************************************************************
Private Sub mnuInputAbsenteeRecord2_Click()

    ' absentee record for 2nd phase

'2021.12.29 del jhi
#If 0 Then
    If frmAbsenteeRecord Is Nothing Then
        Set frmAbsenteeRecord = New frmExamineeStatus
    End If

    With frmAbsenteeRecord
        .m_int_IntRpt = 0
        .m_int_Action = 2
        .Show
        .Caption = "frmExamineeStatus : 2次 欠席者入力"
        .ZOrder 0
    End With
#End If


    '2021.12.29 add jhi
    frm2jikesseki.m_int_IntRpt = 0
    frm2jikesseki.m_int_Action = 2

    frm2jikesseki.Caption = "frm2jikesseki : 2次 欠席者入力"
    frm2jikesseki.Show
    frm2jikesseki.ZOrder 0



End Sub

'*******************************************************************************
'* 2次試験 : 面接委員登録                                                      *
'*******************************************************************************
Private Sub mnuTeacherRoomMapInterview_Click()

    '面接委員登録
    If frmIntwrRoomMapInt Is Nothing Then
        Set frmIntwrRoomMapInt = New frmInterviewerRoom
    End If

    With frmIntwrRoomMapInt
        .m_int_IntRpt = 0
        .Show
        .Caption = "frmInterviewerRoom : 面接委員登録"   ''''LoadResString(2301)
        .ZOrder 0
    End With

End Sub

'*******************************************************************************
'* 2次試験 : 面接グループ振分                                                  *
'*******************************************************************************
Private Sub mnuPreparationRoom_Click()

    frmPrepSecondExamGrp.Show
    frmPrepSecondExamGrp.Caption = "frmPrepSecondExamGrp : 面接グループ振分"
    frmPrepSecondExamGrp.ZOrder 0

End Sub

'*******************************************************************************
'* 2次試験 : 面接グループ変更                                                  *
'*******************************************************************************
Private Sub mnuManualAllocationGrp_Click()

    frmManualAllocationGrp.Show
    frmManualAllocationGrp.Caption = "frmManualAllocationGrp : 面接グループ変更"
    frmManualAllocationGrp.ZOrder 0

End Sub

'*******************************************************************************
'* 2次試験 : 小論文採点委員登録                                                *
'*******************************************************************************
Private Sub mnuTeacherRoomMapReport_Click()

    ' Teacher-Room Mapping - Report
    If frmIntwrRoomMapRpt Is Nothing Then
        Set frmIntwrRoomMapRpt = New frmInterviewerReport
    End If

    With frmIntwrRoomMapRpt
        .m_int_IntRpt = 1
        .Show
''''    .Caption = LoadResString(2302) '採点者-小論文会場割り当て-
        .Caption = "frmInterviewerReport : 小論文採点委員登録"  ''''LoadResString(2302) '採点者-小論文会場割り当て-
        .ZOrder 0
    End With

End Sub

'*******************************************************************************
'* 2次試験 : 小論文振分                                                        *
'*******************************************************************************
Private Sub mnuPreparationReport_Click()

    Load frmPrepReport

    ' call zOrder only if the interview 1 has taken place before
    If g_bln_InterviewHappened Then
        frmPrepReport.Show
        frmPrepReport.ZOrder 0
    Else
        Unload frmPrepReport
    End If

End Sub

'*******************************************************************************
'* 2次試験 : 素点入力(小論文)_import                                           *
'* 2021.12.12 add jhi                                                          *
'*******************************************************************************
Private Sub mnuImport_Syoronbun_Click()

    g_int_ExamType = 2

''''MsgBox "素点入力(小論文)_import画面表示"

    Call frmImportSyoronbun.gsSetSecondType(1) '1:小論文

    frmImportSyoronbun.Show
    frmImportSyoronbun.Caption = "frmImportSyoronbun : 素点入力(小論文)_import "
    frmImportSyoronbun.ZOrder 0


End Sub

'*******************************************************************************
'* 2次試験 : 素点入力(小論文)                                                  *
'*******************************************************************************
Private Sub mnuInputRawScoreI_Click()

    ' raw score for second phase - report
    g_int_ExamType = 2

    If frmRawScoreRpt Is Nothing Then
        Set frmRawScoreRpt = New frmRawScore
    End If

    With frmRawScoreRpt
        Call .gsSetSecondType(1)    '1:小論文
        .Show
''''    .Caption = LoadResString(1019) ''''素点入力（小論文）
        .Caption = "frmRawScore : 素点入力(小論文)"
        .ZOrder 0
    End With

End Sub

'*******************************************************************************
'* 2次試験 : 素点入力(面接)_import                                             *
'* 2021.12.12 add jhi                                                          *
'*******************************************************************************
Private Sub mnuImport_Mensetu_Click()

    g_int_ExamType = 2


    Call frmImportMensetu.gsSetSecondType(0) '0:面接

    frmImportMensetu.Show
    frmImportMensetu.Caption = "frmImportMensetu : 素点入力(面接)_import "
    frmImportMensetu.ZOrder 0

End Sub

'*******************************************************************************
'* 2次試験 : 素点入力(面接)                                                    *
'*******************************************************************************
Private Sub mnuInputRawScore2_Click()

    ' raw score for second phase - interview
    g_int_ExamType = 2

    If frmRawScoreInt Is Nothing Then
        Set frmRawScoreInt = New frmRawScore
    End If

    With frmRawScoreInt
        Call .gsSetSecondType(0)    '0:面接
        .Show
''''    .Caption = LoadResString(1047)
        .Caption = "frmRawScore : 素点入力(面接)"
        .ZOrder 0
    End With

End Sub

'*******************************************************************************
'* 2次試験 : 合格者入力                                                        *
'*******************************************************************************
Private Sub mnuInputPassedPersonData2_Click()


'2021.12.30 del jhi
#If 0 Then
    ' passed person data for 2nd phase
    If frmPassedPersonData Is Nothing Then
        Set frmPassedPersonData = New frmExamineeStatus
    End If

    With frmPassedPersonData
        .m_int_IntRpt = 3
        .m_int_Action = 3
        .Show
        .Caption = "frmExamineeStatus : 2次 合格者入力"
        .ZOrder 0
    End With

#End If

'2021.12.30 add jhi
    frm2jigoukaku.m_int_IntRpt = 3
    frm2jigoukaku.m_int_Action = 3

    frm2jigoukaku.Caption = "frm2jigoukaku : 2次 合格者入力"
    frm2jigoukaku.Show
    frm2jigoukaku.ZOrder 0


End Sub

'*******************************************************************************
'* 2次試験 : 補欠者入力                                                        *
'*******************************************************************************
Private Sub mnuWaitList2_Click()


'2021.12.30 del jhi
#If 0 Then

    ' input waiting list
    If frmWaitingList Is Nothing Then
        Set frmWaitingList = New frmExamineeStatus
    End If

    With frmWaitingList
        .m_int_IntRpt = 4
        .m_int_Action = 4
        .Show
        .Caption = "frmExamineeStatus : 補欠者入力"
        .ZOrder 0
    End With


#End If

'2021.12.30 add jhi
    frm2jiHoketusya.m_int_IntRpt = 4
    frm2jiHoketusya.m_int_Action = 4

    frm2jiHoketusya.Caption = "frm2jiHoketusya : 2次 補欠者入力"
    frm2jiHoketusya.Show
    frm2jiHoketusya.ZOrder 0


End Sub

'*******************************************************************************
'* 補欠者順位                                                                  *
'*******************************************************************************
'*******************************************************************************
'* 3.10 補欠者順位 --->ほかのsubsystem 画面からこちらに入れる必要がある★      *
'* 2021.12.02 add jhi                                                          *
'*******************************************************************************
Private Sub mnuHoketusyaJuni_Click()

    '2次試験
    g_int_ExamType = 2

    'frmChooseiReport.Show
    'frmChooseiReport.ZOrder 0

    frm2jiHoketusyaJuni.Caption = "frmHoketusyaJuni : 2次 補欠者順位"
    frm2jiHoketusyaJuni.Show
    frm2jiHoketusyaJuni.ZOrder 0

End Sub

'*******************************************************************************
'* データ確定                                                                  *
'*******************************************************************************
Private Sub mnuFixData3_Click()

    Call mnuFixData1_Click

End Sub

'*******************************************************************************
'* ４．入学手続き処理                                                          *
'*******************************************************************************
'*******************************************************************************
'* 補欠者合格者繰上げ処理                                                      *
'*******************************************************************************
Private Sub mnuUpliftment_Click()


'2021.12.29 del jhi
#If 0 Then
    ' upliftment from waiting list
    If frmUpliftment Is Nothing Then
        Set frmUpliftment = New frmExamineeKuriage
    End If

    With frmUpliftment
        .m_int_IntRpt = 5
        .m_int_Action = 5
        .Show
        .Caption = "frmExamineeKuriage : 補欠者合格繰上げ処理"
        .ZOrder 0
    End With
#End If

    frmExamineeKuriage.m_int_IntRpt = 5
    frmExamineeKuriage.m_int_Action = 5

    frmExamineeKuriage.Caption = "frmExamineeKuriage : 補欠者合格繰上げ処理"
    frmExamineeKuriage.Show
    frmExamineeKuriage.ZOrder 0

End Sub

'*******************************************************************************
'* 辞退                                                                        *
'*******************************************************************************
Private Sub mnuRefuseOffer_Click()

'2021.12.29 del jhi
#If 0 Then
    ' enter/refuse screen
    If frmRefuseOffer Is Nothing Then
'        Set frmRefuseOffer = New frmExamineeStatus
        Set frmRefuseOffer = New frmExamineeKuriage
    End If

    With frmRefuseOffer
        .m_int_IntRpt = 6
        .m_int_Action = 6
        .Show
        .Caption = "frmExamineeKuriage : 辞退"
        .ZOrder 0
    End With
#End If

    frmExamineeJitai.m_int_IntRpt = 6
    frmExamineeJitai.m_int_Action = 6

    frmExamineeJitai.Caption = "frmExamineeJitai : 辞退"
    frmExamineeJitai.Show
    frmExamineeJitai.ZOrder 0


End Sub

'*******************************************************************************
'* データ確定                                                                  *
'*******************************************************************************
Private Sub mnuFixData4_Click()

    Call mnuFixData1_Click

End Sub

'*******************************************************************************
'* ５．マスターメンテナンス Menu                                               *
'*******************************************************************************
'*******************************************************************************
'* 会場・面接グループ                                                          *
'*******************************************************************************
Private Sub mnuRoomProfile_Click()

    frmRoomProfile.Show
    frmRoomProfile.ZOrder 0

End Sub

'*******************************************************************************
'* 採点者プロファイル                                                          *
'*******************************************************************************
Private Sub mnuInterviewerProfile_Click()

    frmInterviewerProfile.Show
    frmInterviewerProfile.ZOrder 0

End Sub

'*******************************************************************************
'* 所属プロファイル                                                            *
'*******************************************************************************
Private Sub mnuInterviewGroupProfile_Click()

    frmInterviewGroupProfile.Show
    frmInterviewGroupProfile.ZOrder 0

End Sub

'*******************************************************************************
'* 入試年度指定                                                                *
'*******************************************************************************
Private Sub mnuSystemData_Click()

    frmSystemData.Show
    frmSystemData.ZOrder 0

End Sub

'*******************************************************************************
'* ６．印刷 Menu                                                               *
'*******************************************************************************
'*******************************************************************************
'* 印刷指示(a61)                                                               *
'*******************************************************************************
Private Sub mnuPrintCommand_Click()

    frmPrintCommand.Show
    frmPrintCommand.ZOrder 0

End Sub

Private Sub mnuPrint_Click()

    On Error GoTo ErrorHandler

    'Mahesh. Commented line g_int_SelectValues = 0 for instances of frmDeptTeacherActivity1
    'To facilitate reprinting of the same data
    Dim f_bln_ClickPrint As Boolean
    Dim f_obj_frm        As Object


    f_bln_ClickPrint = True
    Set f_obj_frm = fMainForm.ActiveForm

    ' call th respective "Print" functions based on the Active form
    If f_obj_frm.Name = "frmPrintCommand" Then
        If f_bln_ClickPrint = True Then
            frmPrintCommand.cmdPrint_Click
            f_bln_ClickPrint = False
        End If
    Else
        fMainForm.ActiveForm.f_void_Print
    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description

End Sub

'*******************************************************************************
'* Excel帳票(a62)                                                              *
'*******************************************************************************
Private Sub mnuExcelReport_Click()

    frmExcelReport.Show
    frmExcelReport.ZOrder 0

End Sub

'*******************************************************************************
'* 度数分布図印刷(a63)                                                          *
'*******************************************************************************
Private Sub mnuPrintDosu_Click()

    frmPrintDosu.Show
    frmPrintDosu.ZOrder 0

End Sub

'*******************************************************************************
'* ７．データ出力                                                              *
'*******************************************************************************
'*******************************************************************************
'* 受験生、素点情報                                                            *
'*******************************************************************************



'*******************************************************************************
'* 以下、未使用                                                                *
'*******************************************************************************
Private Sub mnuAdjustScoreM_Click()

    ' choosei score for Interview in 2nd phase
    g_int_ExamType = 2
    frmChooseiInterview.Show
    frmChooseiInterview.ZOrder 0

End Sub

Private Sub mnuAdjustScoreS_Click()

    ' choosei score for Report in 2nd phase
    g_int_ExamType = 2
    frmChooseiReport.Show
    frmChooseiReport.ZOrder 0

End Sub

Private Sub mnuCascade_Click()

    fMainForm.Arrange 0

End Sub

Public Sub mnuExamineeCheck_Click()

    frmExamineeCheck.Show
    frmExamineeCheck.ZOrder 0

End Sub

Private Sub mnuHelp_Click()

    frmHelp.Show 1

End Sub

Private Sub mnuInputChooseiScore_Click()
    ' choosei score for the 1st phase
    If frmChooseiGrace Is Nothing Then
        Set frmChooseiGrace = New frmChooseiScore1
    End If
    frmChooseiGrace.Show
    frmChooseiGrace.ZOrder 0
End Sub

Private Sub mnuExit_Click()

    Dim l_frm As Form
    Dim rinf  As Long


    rinf = myMsgBox("入試システムを終了します。よろしいですか？", gTit)
    If rinf = vbCancel Then
        Exit Sub
    End If


    For Each l_frm In Forms
        Unload l_frm
    Next

    End


End Sub

Private Sub mnuHighSchoolType_Click()
    frmHighSchoolType.Show
    frmHighSchoolType.ZOrder 0
End Sub
'*******************************************************************************
'* 評定                                                                        *
'*******************************************************************************
Private Sub mnuHyotei_Click()
    ' raw score for the apply phase
    frmRawScore.Show
    frmRawScore.ZOrder 0
End Sub

Private Sub mnuInputChooseiScore2_Click()
    'choosei score for Hyotei
    If frmChooseiHyotei Is Nothing Then
        Set frmChooseiHyotei = New frmChooseiJoken
    End If
    frmChooseiHyotei.m_int_ChoseiJoken = 1
    frmChooseiHyotei.Show
    frmChooseiHyotei.ZOrder 0
End Sub

'Private Sub mnuInputChooseiScorePoint_Click()
'    If frmChooseiPoint Is Nothing Then
'        Set frmChooseiPoint = New frmChooseiJoken
'    End If
'    frmChooseiPoint.Show
'    frmChooseiPoint.ZOrder 0
'End Sub


'*******************************************************************************
'* IVR                                                                         *
'*******************************************************************************
Private Sub mnuIVRTransfer_Click()

''''2021.12.29 del jhi

''''    frmOutputIVR.Show
''''    frmOutputIVR.ZOrder 0

End Sub

Private Sub mnuMaintainExamineeData2_Click()

    Call mnuMaintainExamineeData_Click

End Sub

Private Sub mnuOutputCSV_Click()
    frmCSVOutput.Show
    frmCSVOutput.ZOrder 0
End Sub

'*******************************************************************************
'* 小論文入力  <----使わないようだ2021.12.22 確認                              *
'*******************************************************************************
'add,xzg,2010/12/09,S
Private Sub mnuCommWork_Click()
    frmCompWork.Show
    frmCompWork.ZOrder 0
End Sub
'add,xzg,2010/12/09,E

Private Sub mnuSeisekiIchiran_Click()
    frmSeisekiIchiranProfile.Show
    frmSeisekiIchiranProfile.ZOrder 0
End Sub

Private Sub mnuShowTree_Click()
    If mnuShowTree.Checked Then
        mnuShowTree.Checked = False
        pctExplorer.Visible = False
    Else
        pctExplorer.Visible = True
        mnuShowTree.Checked = True
    End If
End Sub

Private Sub mnuSpecialInterview_Click()

    frmSpecialInterview.Show
    frmSpecialInterview.ZOrder 0

End Sub

Private Sub mnuSubjectProfile_Click()

    frmSubjectProfile.Show
    frmSubjectProfile.ZOrder 0

End Sub

Private Sub mnuSubjectQuestionProfile_Click()

    frmSubjectQuestionProfile.Show
    frmSubjectQuestionProfile.ZOrder 0

End Sub

Private Sub mnuTest_Click()

    fMainForm.Arrange 1

End Sub

Private Sub mnuTileHorizontally_Click()

    fMainForm.Arrange 1

End Sub

Private Sub mnuTileVertically_Click()
    fMainForm.Arrange 2
End Sub

Private Sub mnuToolsQuery_Click()

    Call mnuToolsClear_Click
    fMainForm.ActiveForm.m_bMode = "SEARCH"
    mnuToolsSearch.Enabled = True
    mnuToolsDelete.Enabled = False ' added by mahesh to disable delete in query mode
    fMainForm.Toolbar1.Buttons("Search").Enabled = True
    fMainForm.Toolbar1.Buttons("Delete").Enabled = False
    fMainForm.Toolbar1.Buttons("Clear").Enabled = True

End Sub

Private Sub mnuZipCode_Click()

    frmZipCode.Show
    frmZipCode.ZOrder 0

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If

End Sub

Private Sub mnuToolsCancel_Click()

    'this menu will be enabled only for the data entry form- and dirty mode
    'this will be enabled only if save is enabled
    Call CancelData

End Sub

Private Sub mnuToolsClear_Click()

    'this menu will be enabled only for the data entry form- in short depends on the mode
    Call ClearData

End Sub

Private Sub mnuToolsDelete_Click()
    'this menu will be enabled only for the exiting data- in short depends on the mode
    Call DeleteData
End Sub

Private Sub mnuToolsNew_Click()

    'this menu will be enabled only for the data entry form
   Call NewData

   fMainForm.ActiveForm.lblErrorMsg.Caption = ""

End Sub

Private Sub mnuToolsSave_Click()

    'this menu is enabled only for the data entry form
    Call ValidateAndSaveData

End Sub

Private Sub mnuToolsSearch_Click()

    Call SearchRecords

    mnuToolsSearch.Enabled = False

    Call NewData  'calling this again after search is complete disables delete button Mahesh

End Sub

'*******************************************************************************
'* TreeViewのメニューを表示する                                                *
'* 2021.12.10 comm add jhi                                                     *
'*******************************************************************************
Private Sub Init_TreeView(puMenues_() As prvuMenues_Type)

    On Error GoTo ErrorHandler

    Dim l_obj_NewNode As Object
    Dim lCnt          As Long
    Dim step_no       As Integer


step_no = 1

    '2021.11.11 add jhi
    tvwMenu.Nodes.Clear

    With tvwMenu

        ''''2021.12.09 add jhi Treeメニューに試験区分を入れる
        Set l_obj_NewNode = .Nodes.Add(, , "mnuExamKubun", "【 前期試験 】" & f_int_CurrentPhase) '前期試験

        For lCnt = LBound(puMenues_) To UBound(puMenues_)

            If puMenues_(lCnt).bVisible Then
                If puMenues_(lCnt).lParent = -1 Then
                '親
                    Set l_obj_NewNode = .Nodes.Add(, , puMenues_(lCnt).sTVKey, puMenues_(lCnt).sCaption)
                Else
                '子
                    If puMenues_(puMenues_(lCnt).lParent).bVisible Then
                        Set l_obj_NewNode = .Nodes.Add(puMenues_(puMenues_(lCnt).lParent).sTVKey, tvwChild, puMenues_(lCnt).sTVKey, puMenues_(lCnt).sCaption)
                    End If
                End If
            End If

        Next

    End With

step_no = 2


    Exit Sub


ErrorHandler:
step_no = 3



End Sub


'*******************************************************************************
'* TreeViewのメニューを表示する                                                *
'* 2021.12.10 comm add jhi                                                     *
'*******************************************************************************
Private Sub Init_TreeView_New(puMenues_() As prvuMenues_Type)

    On Error GoTo ErrorHandler

    Dim l_obj_NewNode As Object
    Dim lCnt          As Long
    Dim step_no       As Integer


step_no = 1

    '2021.11.11 add jhi
    tvwMenu.Nodes.Clear

    With tvwMenu


''''----------------------------------------------------------------------------
''''条件付きコンパイル引数の設定 2022.02.01 add jhi
''''----------------------------------------------------------------------------
#If zengo_kubun = 1 Then

        ''''2021.12.09 add jhi Treeメニューに試験区分を入れる
        Set l_obj_NewNode = .Nodes.Add(, , "mnuExamKubun", "■前期試験 -  " & f_int_CurrentPhase + 1 & " フェーズ") 'フェーズ表示

#Else
        Set l_obj_NewNode = .Nodes.Add(, , "mnuExamKubun", "■後期試験 -  " & f_int_CurrentPhase + 1 & " フェーズ") 'フェーズ表示
#End If

        For lCnt = LBound(puMenues_) To UBound(puMenues_)

            If puMenues_(lCnt).bVisible Then
                If puMenues_(lCnt).lParent = -1 Then
                '親
                    Set l_obj_NewNode = .Nodes.Add(, , puMenues_(lCnt).sTVKey, puMenues_(lCnt).sCaption)
                Else
                '子
                    If puMenues_(puMenues_(lCnt).lParent).bVisible Then
                        Set l_obj_NewNode = .Nodes.Add(puMenues_(puMenues_(lCnt).lParent).sTVKey, tvwChild, puMenues_(lCnt).sTVKey, puMenues_(lCnt).sCaption)
                    End If
                End If
            End If

        Next

    End With

step_no = 2


''''----------------------------------------------------------------------------
''''条件付きコンパイル引数の設定 2022.02.01 add jhi
''''----------------------------------------------------------------------------
#If zengo_kubun = 1 Then

''''Call SetTVBackColor(tvwMenu, RGB(240, 248, 255)) 'aliceblue
    Call SetTVBackColor(tvwMenu, RGB(255, 250, 250)) 'Snow

#Else

''''Call SetTVBackColor(tvwMenu, RGB(255, 250, 250)) 'Snow     2022.01.06 del jhi
    Call SetTVBackColor(tvwMenu, RGB(249, 207, 98))  'Yellow系 2022.01.06 add jhi
#End If


    '---------------------------------------------------------------------------
    'Menuを全部広げる
    '2021.12.22 add jhi
    '---------------------------------------------------------------------------
    Dim objNode As Node
    For Each objNode In tvwMenu.Nodes
        objNode.Expanded = True
    Next



    Exit Sub


ErrorHandler:

step_no = 3
    MsgBox "Init_TreeView_New関数でエラーが発生しました。"


End Sub




Private Function lf_StrNullCut(psInStr As String) As String

    Dim lPos As Long

    lPos = InStr(1, psInStr, vbNullChar)

    If lPos > 0 Then
        lf_StrNullCut = Left$(psInStr, lPos - 1)
    Else
        lf_StrNullCut = psInStr
    End If

End Function

'*******************************************************************************
'* Tree Menu 文字列 set                                                        *
'* 2021.12.2 update jhi                                                        *
'*******************************************************************************
Private Sub ls_SetMenues(puMenues_() As prvuMenues_Type)

    Dim lCnt      As Long
    Dim i         As Integer
    Dim strMsg    As String


    Erase puMenues_

    'update,xzg,2010/12/09,S
    'ReDim puMenues_(52)
''''ReDim puMenues_(53)    ''''2021.11.30 del IVR
    'update,xzg,2010/12/09,E

    'index 0-47
    ReDim puMenues_(41)    ''''2021.12.21 update 削除Menu対応


''''lCnt = 0

''''Set puMenues_(lCnt).oMnuObj = mnuExamKubun
''''puMenues_(lCnt).sTVKey = "nodeExamKubun"
''''puMenues_(lCnt).lParent = -1
''''puMenues_(lCnt).sCaption = "前期試験"

''''If g_int_ExamKubun = 1 Then
''''    puMenues_(lCnt).sCaption = "前期試験"
''''Else
''''    puMenues_(lCnt).sCaption = "後期試験"
''''End If

    
    lCnt = 0

    Set puMenues_(lCnt).oMnuObj = mnuApplyPhase
    puMenues_(lCnt).sTVKey = "nodeApplyPhase"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "1. 願書受付フェーズ"     ''''LoadResString(1002) '願書受付フェーズ

    lCnt = lCnt + 1 '1
    Set puMenues_(lCnt).oMnuObj = mnu1stExam
    puMenues_(lCnt).sTVKey = "nodeFirstPhase"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "2. 一次試験"             ''''LoadResString(1008) '一次試験

    lCnt = lCnt + 1 '2
    Set puMenues_(lCnt).oMnuObj = mnu2ndExam
    puMenues_(lCnt).sTVKey = "nodeSecondPhase"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "3. 二次試験"             ''''LoadResString(1016) '二次試験

    lCnt = lCnt + 1 '3
    Set puMenues_(lCnt).oMnuObj = mnuEnterRefuse
    puMenues_(lCnt).sTVKey = "nodeEnterRefuse"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "4. 入学手続き処理"       ''''LoadResString(1024) '入学手続き処理

    lCnt = lCnt + 1 '4
    Set puMenues_(lCnt).oMnuObj = mnuMaster
    puMenues_(lCnt).sTVKey = "nodeMasters"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "5. マスターメンテナンス" ''''LoadResString(1028) 'マスターメンテナンス

    lCnt = lCnt + 1 '5
    Set puMenues_(lCnt).oMnuObj = mnuPrint
    puMenues_(lCnt).sTVKey = "nodePrint"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "6. 印刷"                 ''''LoadResString(1090) '印刷

    lCnt = lCnt + 1 '6
    Set puMenues_(lCnt).oMnuObj = mnuTransfer
    puMenues_(lCnt).sTVKey = "nodeTransfer"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "7. 受験データCSV出力"        ''''LoadResString(1091) 'データ出力


    '***************************************************************************
    '* 願書受付フェーズ Menu                                                   *
    '***************************************************************************
    lCnt = lCnt + 1 '7
    puMenues_(lCnt).sTVKey = "a01"
    puMenues_(lCnt).lParent = 0
    puMenues_(lCnt).sCaption = "Web出願データ取込"       ''''LoadResString(1003)

    lCnt = lCnt + 1 '8
    puMenues_(lCnt).sTVKey = "a02"
    puMenues_(lCnt).lParent = 0
    puMenues_(lCnt).sCaption = "受験生データの編集"      ''''LoadResString(1004)

    lCnt = lCnt + 1 '9
    puMenues_(lCnt).sTVKey = "a03"
    puMenues_(lCnt).lParent = 0
    puMenues_(lCnt).sCaption = "データ確定"              ''''LoadResString(1007)

''''2021.12.01 del jhi 評定
''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a03"
'''' puMenues_(lCnt).lParent = 0
''''puMenues_(lCnt).sCaption = "評定"                    ''''LoadResString(1005)


    '***************************************************************************
    '* 1次試験 Menu                                                            *
    '***************************************************************************
    lCnt = lCnt + 1 '10
    puMenues_(lCnt).sTVKey = "a11"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "会場入力"                ''''LoadResString(1009) '会場入力

    lCnt = lCnt + 1 '11
    puMenues_(lCnt).sTVKey = "a12"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "欠席者入力"              ''''LoadResString(1010) '欠席者入力

    lCnt = lCnt + 1 '12
    puMenues_(lCnt).sTVKey = "a13"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "素点入力"                ''''LoadResString(1011) '素点入力 <---importはしない

    lCnt = lCnt + 1 '13
    puMenues_(lCnt).sTVKey = "a14"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "合格者入力"              ''''LoadResString(1013) '合格者入力

    lCnt = lCnt + 1 '14
    puMenues_(lCnt).sTVKey = "a15"
    puMenues_(lCnt).lParent = 1
''''2022.03.09 del jhi
''''puMenues_(lCnt).sCaption = "二次試験日振分"           ''''LoadResString(1080) '二次試験日振分

#If zengo_kubun = 1 Then
    strMsg = "二次試験日振分"
#Else
    strMsg = "二次試験日確定"
#End If

    puMenues_(lCnt).sCaption = strMsg                     ''''2022.03.09 add jhi 前期、後期のTitileを変更


    lCnt = lCnt + 1 '15
    puMenues_(lCnt).sTVKey = "a16"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "二次試験日変更"           ''''LoadResString(1081) '二次試験日変更

    lCnt = lCnt + 1 '16
    puMenues_(lCnt).sTVKey = "a17"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "データ確定"              ''''LoadResString(1007) 'データ確定


''''2021.12.01 del jhi
''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a14"
''''puMenues_(lCnt).lParent = 1
''''puMenues_(lCnt).sCaption = "科目別調整点入力"        ''''LoadResString(1012) '科目別調整点入力

    'del,xzg,2009/12/02,S----------
    'lCnt = 14      'xx
    'puMenues_(lCnt).sTVKey = "a15"
    'puMenues_(lCnt).lParent = 1
    'puMenues_(lCnt).sCaption = "条件別調整点入力"       ''''LoadResString(1046) '条件別調整点入力
    'del,xzg,2009/12/02,E----------

  
    '***************************************************************************
    '* 2次試験 Menu                                                            *
    '***************************************************************************

    lCnt = lCnt + 1 '17
    puMenues_(lCnt).sTVKey = "a21"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "欠席者入力"    ''''LoadResString(1018)
    
    '---------------------------------------------------------------------------
    ' 面接関連 3Menu
    '---------------------------------------------------------------------------
    lCnt = lCnt + 1 '18
    puMenues_(lCnt).sTVKey = "a22"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "面接委員登録"     ''''LoadResString(1051)

    lCnt = lCnt + 1 '19
    puMenues_(lCnt).sTVKey = "a23"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "面接グループ振分" ''''LoadResString(1082)

    lCnt = lCnt + 1 '20
    puMenues_(lCnt).sTVKey = "a24"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "面接グループ変更" ''''LoadResString(1083)
'-------------------------------------------------------------------------------

    lCnt = lCnt + 1 '21
    puMenues_(lCnt).sTVKey = "a25"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "小論文採点委員登録"    ''''LoadResString(1053)

    lCnt = lCnt + 1 '22 小論文振分
    puMenues_(lCnt).sTVKey = "a26"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "小論文振分"    ''''LoadResString(2433)


    '---------------------------------------------------------------------------
    ''''2021.12.12 add jhi
    lCnt = lCnt + 1 '23
    puMenues_(lCnt).sTVKey = "a27"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "素点入力(小論文)_import"

    lCnt = lCnt + 1 '24 素点入力(小論文)
    puMenues_(lCnt).sTVKey = "a28"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "素点入力(小論文)"        ''''LoadResString(1019)
    '---------------------------------------------------------------------------

    lCnt = lCnt + 1 '25
    puMenues_(lCnt).sTVKey = "a29"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "素点入力(面接)_import"

    lCnt = lCnt + 1 '26
    puMenues_(lCnt).sTVKey = "a30"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "素点入力(面接)"          ''''LoadResString(1047)

    lCnt = lCnt + 1 '27
    puMenues_(lCnt).sTVKey = "a31"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "合格者入力"              ''''LoadResString(1021)

    lCnt = lCnt + 1 '28
    puMenues_(lCnt).sTVKey = "a32"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "補欠者入力"              ''''LoadResString(1022)

    lCnt = lCnt + 1 '29
    puMenues_(lCnt).sTVKey = "a33"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "補欠者順位"              ''''★subsystemから導入した

    lCnt = lCnt + 1 '30
    puMenues_(lCnt).sTVKey = "a34"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "データ確定"              ''''LoadResString(1007)


    'add,xzg,2010/12/09,S-----------
    '小論文入力
''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a27"
''''puMenues_(lCnt).lParent = 2
''''puMenues_(lCnt).sCaption = "小論文入力"           '<---表示されない
    'add,xzg,2010/12/09,E-----------

''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a28"
''''puMenues_(lCnt).lParent = 2
''''puMenues_(lCnt).sCaption = "２次面接グループ生成" '<---表示されない

'-------------------------------------------------------------------------------
' 2021.12.02 del jhi
'-------------------------------------------------------------------------------
'    lCnt = lCnt + 1 'xx 調整点入力(小論文)
'    puMenues_(lCnt).sTVKey = "a30"
'    puMenues_(lCnt).lParent = 2
'    puMenues_(lCnt).sCaption = LoadResString(1048)
'
'    lCnt = lCnt + 1 'xx 調整点入力(面接)
'    puMenues_(lCnt).sTVKey = "a36"
'    puMenues_(lCnt).lParent = 2
'    puMenues_(lCnt).sCaption = LoadResString(1049)
'-------------------------------------------------------------------------------


    '***************************************************************************
    '* 入学手続き処理 Menu                                                     *
    '***************************************************************************

    lCnt = lCnt + 1 '31
    puMenues_(lCnt).sTVKey = "a41"
    puMenues_(lCnt).lParent = 3
    puMenues_(lCnt).sCaption = "補欠者合格繰上げ処理"    ''''LoadResString(1025)

    lCnt = lCnt + 1 '32
    puMenues_(lCnt).sTVKey = "a42"
    puMenues_(lCnt).lParent = 3
    puMenues_(lCnt).sCaption = "辞退"                    ''''LoadResString(1026)

    lCnt = lCnt + 1 '33
    puMenues_(lCnt).sTVKey = "a43"
    puMenues_(lCnt).lParent = 3
    puMenues_(lCnt).sCaption = "データ確定"              ''''LoadResString(1007)


    '***************************************************************************
    '* マスターメンテナンス Menu                                               *
    '***************************************************************************

    lCnt = lCnt + 1 '34
    puMenues_(lCnt).sTVKey = "a51"
    puMenues_(lCnt).lParent = 4
    puMenues_(lCnt).sCaption = "会場・面接グループ"     ''''LoadResString(1031)

    lCnt = lCnt + 1 '35
    puMenues_(lCnt).sTVKey = "a52"
    puMenues_(lCnt).lParent = 4
    puMenues_(lCnt).sCaption = "採点者プロファイル"      ''''LoadResString(1033)

    lCnt = lCnt + 1 '36
    puMenues_(lCnt).sTVKey = "a53"
    puMenues_(lCnt).lParent = 4
    puMenues_(lCnt).sCaption = "所属プロフィール"        ''''LoadResString(2466)

    lCnt = lCnt + 1 '37
    puMenues_(lCnt).sTVKey = "a54"
    puMenues_(lCnt).lParent = 4
    puMenues_(lCnt).sCaption = "入試年度指定"            ''''LoadResString(2600) 'システムパラメータ


''''2021.12.21 del jhi S====
''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a55"
''''puMenues_(lCnt).lParent = 4
''''puMenues_(lCnt).sCaption = "高校区分"                ''''LoadResString(1029)

''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a56"
''''puMenues_(lCnt).lParent = 4
''''puMenues_(lCnt).sCaption = "郵便番号 ID"             ''''LoadResString(1030)
''''2021.12.21 del jhi E====

''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a57"
''''puMenues_(lCnt).lParent = 4
''''puMenues_(lCnt).sCaption = "科目プロファイル"        ''''LoadResString(1032)

''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a58"
''''puMenues_(lCnt).lParent = 4
''''puMenues_(lCnt).sCaption = "科目問題プロファイル"    ''''LoadResString(2458)


    '***************************************************************************
    '* 印刷 Menu                                                               *
    '***************************************************************************
    lCnt = lCnt + 1 '38
    puMenues_(lCnt).sTVKey = "a61"
    puMenues_(lCnt).lParent = 5
    puMenues_(lCnt).sCaption = "印刷指示"                ''''LoadResString(1092)

    lCnt = lCnt + 1 '39
    puMenues_(lCnt).sTVKey = "a62"
    puMenues_(lCnt).lParent = 5
    puMenues_(lCnt).sCaption = "Excel帳票"               ''''LoadResString(1093)

    lCnt = lCnt + 1 '40
    puMenues_(lCnt).sTVKey = "a63"
    puMenues_(lCnt).lParent = 5
    puMenues_(lCnt).sCaption = "度数分布図印刷"          '''''LoadResString(2700)

''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a64"
''''puMenues_(lCnt).lParent = 5
''''puMenues_(lCnt).sCaption = "成績一覧"                ''''LoadResString(1094)


    '***************************************************************************
    '* データ出力 Menu                                                         *
    '***************************************************************************

    lCnt = lCnt + 1 '41
    puMenues_(lCnt).sTVKey = "a71"
    puMenues_(lCnt).lParent = 6
    puMenues_(lCnt).sCaption = "受験生＋素点情報"        ''''LoadResString(1096)


''''2021.11.30 del IVRシステムへのデータ転送
''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a71"
''''puMenues_(lCnt).lParent = 6
''''puMenues_(lCnt).sCaption = LoadResString(1095)


''''Debug.Print "配列数: lCnt=" & lCnt


    '***************************************************************************
    '* 設定 Menu Key　内容をBufferに設定                                       *
    '***************************************************************************
    For i = LBound(puMenues_) To UBound(puMenues_)
        puMenues_(i).sIniKey = puMenues_(i).sTVKey
''''    puMenues_(i).bVisible = False
    Next i

''''Debug.Print "i=" & i


End Sub

Private Function lf_GetMenuIndex(puMenues_() As prvuMenues_Type, lKeyID As Long, sKeyData As String) As Long

    Dim lCnt As Long

    On Error GoTo ErrProc

    lf_GetMenuIndex = -1

    Select Case lKeyID
    Case 0
        For lCnt = LBound(puMenues_) To UBound(puMenues_)
            If puMenues_(lCnt).sTVKey = sKeyData Then
                lf_GetMenuIndex = lCnt
                Exit Function
            End If
        Next
    Case 1
        For lCnt = LBound(puMenues_) To UBound(puMenues_)
            If puMenues_(lCnt).sIniKey = sKeyData Then
                lf_GetMenuIndex = lCnt
                Exit Function
            End If
        Next
    Case 2
        For lCnt = LBound(puMenues_) To UBound(puMenues_)
            If puMenues_(lCnt).oMnuObj.Name = sKeyData Then
                lf_GetMenuIndex = lCnt
                Exit Function
            End If
        Next
    End Select

Exit Function

ErrProc:

End Function

'*******************************************************************************
'* ini file の MENU Section 所得                                               *
'* 2022.02.01 update jhi                                                       *
'*******************************************************************************
Private Sub SetPhaseMenu(f_int_CurrentPhase As Long)

    On Error GoTo ERR_HANDLE

    Dim oRs          As ADODB.Recordset
    Dim sSQL         As String

'ユーザ、MACアドレス、業務PHASEより表示するメニューを決定する

    'MACアドレスの取得
    Dim lAdptCnt     As Long
    Dim sErrMsg      As String
    Dim lCnt         As Long
    Dim sMacAddr     As String
    Dim sCnvMacAddr  As String
    Dim sCnvUserID   As String
    Dim sMenuIDStr   As String
    Dim lMenuID      As Long
    Dim sMenuString  As String
    Dim sMenuSection As String
    Dim sProfileName As String
    Dim sFile        As String
    Dim oGao         As Object
    Dim bRet         As Boolean
    Dim sKey         As String

    Dim lRtn         As Long
    Dim sRtn         As String

''''2021.12.28 del jhi globalに宣言
''''Dim uMenues_() As prvuMenues_Type

    Dim sUserPass    As String
    Dim sMacPass     As String
    Dim sMenuPass    As String
    Dim sMenuGPass   As String
    Dim sIniPass     As String


    lAdptCnt = mAdptInf.gfLoadAdptData(sErrMsg)

    If lAdptCnt < 1 Then
        MsgBox "ＭＡＣアドレスの取得に失敗しました。" & vbCrLf & sErrMsg, vbOKOnly, "初期処理不正"
        End
    End If

    'ユーザ、MACアドレスより表示可能メニュー初期化ファイルのセクション名を取得する
    '一応、アダプタの数だけループするようにしておく（２枚差しがあるので）
    sMenuString = ""
    Set oGao = CreateObject("GaoEncode.GaoeAPI")

    For lCnt = 0 To lAdptCnt - 1

        sMacAddr = Replace(mAdptInf.getMacAddr(lCnt), "-", "")
Call log("1-----> sMacAddr=" & sMacAddr)


        'ユーザＩＤを暗号化
        sUserPass = GetSetting("Nyushi", "Settings", "USER", "USER")
        sCnvUserID = Replace(oGao.EncodeStr(Trim(str(glUserID)), sUserPass, 0), vbCrLf, "")
Call log("2-----> sUserPass=" & sUserPass & " sCnvUserID=" & sCnvUserID)



        'MACアドレスを暗号化
        sMacPass = GetSetting("Nyushi", "Settings", "MAC", "MAC")
        sCnvMacAddr = Replace(oGao.EncodeStr(sMacAddr, sMacPass, 0), vbCrLf, "")
Call log("3-----> sMacPass=" & sMacPass & " sCnvMacAddr=" & sCnvMacAddr)

        '暗号化データをキーにメニューグループを取得
        sSQL = ""
        sSQL = sSQL & "SELECT vDATA1 "
        sSQL = sSQL & "FROM tbSTEWorkTbl "

Call log("4-----> sSQL=" & sSQL)

        'update,xzg,2009/12/02,S------------
        'sSQL = sSQL & "WHERE vKEY1 = '" & sCnvMacAddr & "' "
        'sSQL = sSQL & " AND vKEY2 = '" & sCnvUserID & "' "
        sSQL = sSQL & " WHERE vKEY2 = '" & sCnvUserID & "' "
        'update,xzg,2009/12/02,E------------

#If 0 Then
SELECT
    *
From
    tbSTEWorkTbl
Where
    vKEY2 = 'XjHatuXhQdcCAAAAAAAAAM8hjIrRnAKo'
#End If


        Set oRs = g_obj_Conn.Execute(sSQL)

        If Not oRs.EOF Then
            sMenuIDStr = oRs.Fields(0)
Call log("5-----> sMenuIDStr=" & sMenuIDStr)

            oRs.Close
            Set oRs = Nothing

            'メニューグループデータを復号
            sMenuGPass = GetSetting("Nyushi", "Settings", "MENUG", "MENUG")
            lMenuID = CLng(Replace(oGao.DecodeStr(sMenuIDStr, sMenuGPass, 0), vbCrLf, ""))

Call log("6-----> sMenuGPass=" & sMenuGPass & " lMenuID=" & lMenuID)



            '暗号化したセクション名を取得
            sSQL = ""
            sSQL = sSQL & "SELECT vMenuString "
            sSQL = sSQL & "FROM tbSTEMenuGroup "
            sSQL = sSQL & "WHERE iMenuGroupID = " & str(lMenuID)

Call log("7-----> sSQL=" & sSQL)


            Set oRs = g_obj_Conn.Execute(sSQL)

            If Not oRs.EOF Then
                sMenuString = oRs.Fields(0)
Call log("8-----> sMenuString=" & sMenuString)

                oRs.Close
                Set oRs = Nothing
                'セクションを復号

                sMenuPass = GetSetting("Nyushi", "Settings", "MENU", "MENU")
                sMenuSection = Replace(oGao.DecodeStr(sMenuString, sMenuPass, 0), vbCrLf, "")

Call log("9-----> sMenuPass=" & sMenuPass & " sMenuSection=" & sMenuSection)

                Exit For
            Else
                Set oRs = Nothing
            End If
        Else
            Set oRs = Nothing
        End If
    Next

    If sMenuString = "" Then
        MsgBox "本システムを使用する権限がありません。", vbOKOnly, "初期処理"
        End
    End If

    '***************************************************************************
    '* TreeView Menu情報を Type Member配列に設定する(初期化)関数               *
    '* TreeView Menu 文字を設定                                                *
    '***************************************************************************
    Call ls_SetMenues(uMenues_)


    '初期化ファイルを復号
    If Right(App.Path, 1) = "\" Then
        sProfileName = App.Path & prvsProfileName & ".LZH"
    Else
        sProfileName = App.Path & "\" & prvsProfileName & ".LZH"
    End If

    oGao.Disguise = 4

    sIniPass = GetSetting("Nyushi", "Settings", "CDPC", "CDPC")
    bRet = oGao.DecodeFile(sProfileName, sIniPass, 0, App.Path)
    
    Set oGao = Nothing
    'If Not bRet Then Exit Sub

    '初期化ファイルを復号
    If Right(App.Path, 1) = "\" Then
        sProfileName = App.Path & prvsProfileName & ".ini"
    Else
        sProfileName = App.Path & "\" & prvsProfileName & ".ini"
    End If

    '**************************************************************************
    '* Passcheck.ini ファイル変更                                             *
    '*------------------------------------------------------------------------*
    '* 2021.12.22 add jhi                                                     *
    '**************************************************************************

''''条件付きコンパイル引数の設定 2022.02.01 add jhi
#If zengo_kubun = 1 Then
    sProfileName = App.Path & "\" & prvsProfileName & "_zenki.ini"     ''''Passcheck_zenki.ini
#Else
    sProfileName = App.Path & "\" & prvsProfileName & "_goki.ini"      ''''Passcheck_goki.ini
#End If


    '初期化ファイルを読取
    '***************************************************************************
    '* Passcheck.Ini ファイルを読取,[MENU2]Sectionのkey(a01=1)を読込み         *
    '* そのmenuを表示するか? 設定する                                          *
    '***************************************************************************
    For lCnt = LBound(uMenues_) To UBound(uMenues_)

        sRtn = Space(4)
        lRtn = GetPrivateProfileString(sMenuSection, uMenues_(lCnt).sIniKey, "0", sRtn, 40, sProfileName)

        'key(a01=1)が設定していればそのmenuは見えるようにする
        If lRtn > 0 Then
            uMenues_(lCnt).bVisible = (lf_StrNullCut(sRtn) = "1")
        End If

    Next

''''MsgBox "lCnt=" & lCnt

    'Kill sProfileName
'*******************************************************************************
'* 2021.12.09 del jhi S                                                        *
'*******************************************************************************
#If 0 Then

    Select Case f_int_CurrentPhase
    Case 0
        g_int_ExamType = 0
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeApplyPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeFirstPhase")).bVisible = False
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeSecondPhase")).bVisible = False
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeEnterRefuse")).bVisible = False
    Case 1
        g_int_ExamType = 1
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeApplyPhase")).bVisible = False
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeFirstPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeSecondPhase")).bVisible = False
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeEnterRefuse")).bVisible = False
    Case 2
        g_int_ExamType = 2
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeApplyPhase")).bVisible = False
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeFirstPhase")).bVisible = False
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeSecondPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeEnterRefuse")).bVisible = False
    Case 3
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeApplyPhase")).bVisible = False
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeFirstPhase")).bVisible = False
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeSecondPhase")).bVisible = False
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeEnterRefuse")).bVisible = True
    End Select

#End If
'*******************************************************************************
'* 2021.12.09 del jhi E                                                        *
'*******************************************************************************



    '***************************************************************************
    '* TreeViewのメニュー項目を全て表示するように表示する                      *
    '* 2021.12.09 add jhi                                                      *
    '***************************************************************************
    Select Case f_int_CurrentPhase
    Case 0
        g_int_ExamType = 0

''''    uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeExamKubun")).bVisible = True     '前期試験、上手に出来ないので

        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeApplyPhase")).bVisible = True    '願書受付フェース
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeFirstPhase")).bVisible = True    '1次試験
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeSecondPhase")).bVisible = True   '2次試験
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeEnterRefuse")).bVisible = True   '入学手続き処理
    Case 1
        g_int_ExamType = 1
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeApplyPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeFirstPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeSecondPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeEnterRefuse")).bVisible = True
    Case 2
        g_int_ExamType = 2
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeApplyPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeFirstPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeSecondPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeEnterRefuse")).bVisible = True
    Case 3
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeApplyPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeFirstPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeSecondPhase")).bVisible = True
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeEnterRefuse")).bVisible = True
    End Select


    uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeMasters")).bVisible = True  'マスタメインテナンス
    uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodePrint")).bVisible = True    '印刷
    uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeTransfer")).bVisible = True 'データ出力

    mnuApplyPhase.Enabled = uMenues_(lf_GetMenuIndex(uMenues_, 2, "mnuApplyPhase")).bVisible
    mnu1stExam.Enabled = uMenues_(lf_GetMenuIndex(uMenues_, 2, "mnu1stExam")).bVisible
    mnu2ndExam.Enabled = uMenues_(lf_GetMenuIndex(uMenues_, 2, "mnu2ndExam")).bVisible
    mnuEnterRefuse.Enabled = uMenues_(lf_GetMenuIndex(uMenues_, 2, "mnuEnterRefuse")).bVisible


'ForDebug
'    For lCnt = LBound(uMenues_) To UBound(uMenues_)
'        If uMenues_(lCnt).lParent <> -1 Then uMenues_(lCnt).bVisible = uMenues_(uMenues_(lCnt).lParent).bVisible
'    Next


    ' Initialize the Tree View
''''Call Init_TreeView(uMenues_)     ''''2021.12.28 del jhi

    Call Init_TreeView_New(uMenues_) ''''2021.12.28 add jhi
    Call lsShowMenuBar(uMenues_)




'    Call Init_TreeView_Old
'
'    Select Case f_int_CurrentPhase
'    Case 0  ' apply phase
'        mnuApplyPhase.Enabled = True
'        mnu1stExam.Enabled = False
'        mnu2ndExam.Enabled = False
'        mnuEnterRefuse.Enabled = False
'        tvwMenu.Nodes.Remove "nodeFirstPhase"
'        tvwMenu.Nodes.Remove "nodeSecondPhase"
'        tvwMenu.Nodes.Remove "nodeEnterRefuse"
'        g_int_ExamType = 0
'    Case 1  ' 1st phase
'        mnuApplyPhase.Enabled = False
'        mnu1stExam.Enabled = True
'        mnu2ndExam.Enabled = False
'        mnuEnterRefuse.Enabled = False
'        tvwMenu.Nodes.Remove "nodeApplyPhase"
'        tvwMenu.Nodes.Remove "nodeSecondPhase"
'        tvwMenu.Nodes.Remove "nodeEnterRefuse"
'        g_int_ExamType = 1
'    Case 2  ' 2nd phase
'        mnuApplyPhase.Enabled = False
'        mnu1stExam.Enabled = False
'        mnu2ndExam.Enabled = True
'        mnuEnterRefuse.Enabled = False
'        tvwMenu.Nodes.Remove "nodeApplyPhase"
'        tvwMenu.Nodes.Remove "nodeFirstPhase"
'        tvwMenu.Nodes.Remove "nodeEnterRefuse"
'        g_int_ExamType = 2
'    Case 3  ' enter/refuse phase
'        mnuApplyPhase.Enabled = False
'        mnu1stExam.Enabled = False
'        mnu2ndExam.Enabled = False
'        mnuEnterRefuse.Enabled = True
'        tvwMenu.Nodes.Remove "nodeApplyPhase"
'        tvwMenu.Nodes.Remove "nodeFirstPhase"
'        tvwMenu.Nodes.Remove "nodeSecondPhase"
'    Case Else
'        mnuApplyPhase.Enabled = False
'        mnu1stExam.Enabled = False
'        mnu2ndExam.Enabled = False
'        mnuEnterRefuse.Enabled = False
'        tvwMenu.Nodes.Remove "nodeApplyPhase"
'        tvwMenu.Nodes.Remove "nodeFirstPhase"
'        tvwMenu.Nodes.Remove "nodeSecondPhase"
'        tvwMenu.Nodes.Remove "nodeEnterRefuse"
'    End Select

    Exit Sub

ERR_HANDLE:
    Set oGao = Nothing
    MsgBox Err.Description

End Sub

Private Sub pctExplorer_Resize()

    On Error GoTo ErrorHandler

    With tvwMenu
        .Top = 0
        .Left = 0
''''    .Width = 2895 ''''2021.11.30 del jhi
        .Width = 3960 ''''2021.11.30 add jhi
        .Height = pctExplorer.Height
    End With

    Exit Sub

ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim l_int_position As Integer
    Dim l_str_Cap As String
    On Error GoTo ErrorHandler
    
    l_str_Cap = fMainForm.ActiveForm.Caption
    l_int_position = InStr(1, l_str_Cap, "_")

    If l_int_position > 0 Then
        l_str_Cap = Mid(l_str_Cap, 1, l_int_position - 1)
    End If

    Select Case Button.Key
        Case "New"
            Call NewData
            fMainForm.ActiveForm.lblErrorMsg.Caption = ""
        Case "Clear" ' retrieve
            mnuToolsSearch_Click
        Case "Cancel"
            Call CancelData
        Case "Delete"
            Call DeleteData
        Case "Save"
            Call ValidateAndSaveData
        Case "Search"
            'New code to display current mode of master maint forms
            l_str_Cap = l_str_Cap & "_" & "Search"     ''''LoadResString(1054) 2021.12.08 update jhi
            fMainForm.ActiveForm.Caption = l_str_Cap
            'New code ends
             mnuToolsQuery_Click
    End Select

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Public Sub InitToolbar()

    Dim imgX As ListImage

    'SetMenuToolbar
     
    'Initialize Listimages
    ' Load icons into the ImageList control collection.
    ' If icon or bmp' have been removed, proceed further
    
    Set imgX = Me.ImageList1.ListImages.Add(, "New", LoadPicture(NEWICON))
    Set imgX = Me.ImageList1.ListImages.Add(, "Clear", LoadPicture(CLEARICON))
    Set imgX = Me.ImageList1.ListImages.Add(, "Cancel", LoadPicture(CANCELICON))
    Set imgX = Me.ImageList1.ListImages.Add(, "Delete", LoadPicture(DELETEICON))
    Set imgX = Me.ImageList1.ListImages.Add(, "Save", LoadPicture(SAVEICON))
    Set imgX = Me.ImageList1.ListImages.Add(, "Search", LoadPicture(SEARCHICON))
    
' set the Toolbar images

    Me.Toolbar1.ImageList = Me.ImageList1
    
    Me.Toolbar1.Buttons("New").Image = "New"
    Me.Toolbar1.Buttons("Clear").Image = "Clear"
    Me.Toolbar1.Buttons("Cancel").Image = "Cancel"
    Me.Toolbar1.Buttons("Delete").Image = "Delete"
    Me.Toolbar1.Buttons("Save").Image = "Save"
    Me.Toolbar1.Buttons("Search").Image = "Search"
    
    Me.Toolbar1.Buttons("New").ToolTipText = LoadResString(1041)
    Me.Toolbar1.Buttons("Clear").ToolTipText = LoadResString(1036)
    Me.Toolbar1.Buttons("Cancel").ToolTipText = LoadResString(1039)
    Me.Toolbar1.Buttons("Delete").ToolTipText = LoadResString(1038)
    Me.Toolbar1.Buttons("Save").ToolTipText = LoadResString(1037)
    Me.Toolbar1.Buttons("Search").ToolTipText = "検索" ''''LoadResString(1054)

End Sub

'*******************************************************************************
'* TreeView Menuから選択した際の処理                                           *
'*******************************************************************************
Private Sub tvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)


    On Error GoTo ErrorHandler

    Select Case Node.Key

    '---------------------------------------------------------------------------
    ' 願書受付フェーズ(0)
    '---------------------------------------------------------------------------
    Case "a01"     'Web出願データ取込
''''    Call Phase_FlagSet(0)
        mnuOCR_Click

    Case "a02"      '受験生データの編集
''''    Call Phase_FlagSet(0)
        mnuMaintainExamineeData_Click

    Case "a03"      'データ確定
''''    Call Phase_FlagSet(1)
        f_int_CurrentPhase = 0
        mnuFixData1_Click

''''Case "a03"      '評定
''''    Call Phase_FlagSet(0)
''''    mnuHyotei_Click


        
    '---------------------------------------------------------------------------
    ' 1次試験(1)
    '---------------------------------------------------------------------------
    Case "a11"     '会場入力
''''    Call Phase_FlagSet(1)
        mnuRoomAllocation_Click

    Case "a12"     '欠席者入力
''''    Call Phase_FlagSet(1)
        mnuInputAbsenteeRecord_Click

    Case "a13"     '素点入力
''''    Call Phase_FlagSet(1)
        mnuInputRawScore_Click

    Case "a14"      '合格者入力
''''    Call Phase_FlagSet(1)
        mnuInputPassedPersonData_Click

    Case "a15"      '2次試験日振分
''''    Call Phase_FlagSet(1)
        mnuPreparationDay_Click

    Case "a16"      '2次試験日変更
''''    Call Phase_FlagSet(1)
        mnuManualAllocation_Click

    Case "a17"      'データ確定
        Call Phase_FlagSet(2)
        f_int_CurrentPhase = 1
        mnuFixData2_Click


''''----------------------------------------------------------------------------
''''Case "a14"     ' input choosei score - grace
''''    Call Phase_FlagSet(1)
''''    mnuInputChooseiScore_Click
''''
''''Case "a15"      'input choosei score - particular student
''''    Call Phase_FlagSet(1)
''''    mnuInputChooseiScore2_Click

'add,xzg,2009/12/02,S-----------
'    Case "a73"
'        mnuInputChooseiScorePoint_Click
'add,xzg,2009/12/02,E-----------
''''----------------------------------------------------------------------------



    '---------------------------------------------------------------------------
    ' 2次試験(2)
    '---------------------------------------------------------------------------
    Case "a21"     '欠席者入力
''''    Call Phase_FlagSet(2)
        mnuInputAbsenteeRecord2_Click

    Case "a22"     '面接委員登録
''''    Call Phase_FlagSet(2)
        mnuTeacherRoomMapInterview_Click

    Case "a23"     '面接グループ振分
''''    Call Phase_FlagSet(2)
        mnuPreparationRoom_Click

    Case "a24"     '面接グループ変更
''''    Call Phase_FlagSet(2)
'       mnuSpecialInterview_Click
        mnuManualAllocationGrp_Click

    Case "a25"     '小論文採点委員登録
''''    Call Phase_FlagSet(2)
        mnuTeacherRoomMapReport_Click

    Case "a26"     '小論文振分
''''    Call Phase_FlagSet(2)
        mnuPreparationReport_Click

    '---------------------------------------------------------------------------
    Case "a27"     '素点入力(小論文)_import
''''    Call Phase_FlagSet(2)
        mnuImport_Syoronbun_Click

    Case "a28"     '素点入力(小論文)
''''    Call Phase_FlagSet(2)
        mnuInputRawScoreI_Click

    Case "a29"     '素点入力(面接)_import
''''    Call Phase_FlagSet(2)
        mnuImport_Mensetu_Click

    Case "a30"     '素点入力(面接)
''''    Call Phase_FlagSet(2)
        mnuInputRawScore2_Click
    '---------------------------------------------------------------------------

    Case "a31"     '合格者入力
''''    Call Phase_FlagSet(2)
        mnuInputPassedPersonData2_Click

    Case "a32"      '補欠者入力
''''    Call Phase_FlagSet(2)
        mnuWaitList2_Click

    '----------------------------------------------------------------------------
    ' 2021.12.02 add jhi S
    '----------------------------------------------------------------------------
    Case "a33"      '補欠者順位(sub-systemより統合)
''''    Call Phase_FlagSet(2)
        mnuHoketusyaJuni_Click
    '----------------------------------------------------------------------------
    ' 2021.12.02 add jhi E
    '----------------------------------------------------------------------------

   Case "a34"      'データ確定
        Call Phase_FlagSet(3)
        f_int_CurrentPhase = 2
        mnuFixData3_Click


    '----------------------------------------------------------------------------
    ' 入学手続き処理
    '----------------------------------------------------------------------------
    Case "a41"     '補欠者合格者繰上げ処理
''''    Call Phase_FlagSet(3)
        mnuUpliftment_Click

    Case "a42"     '辞退
''''    Call Phase_FlagSet(3)
        mnuRefuseOffer_Click

   Case "a43"      'データ確定
''''    Call Phase_FlagSet(0)
        f_int_CurrentPhase = 3
        mnuFixData4_Click

    '----------------------------------------------------------------------------
    ' マスターメインテナンス
    '----------------------------------------------------------------------------
    Case "a51"     '会場・面接グループ
        mnuRoomProfile_Click

    Case "a52"     '採点者プロファイル
        mnuInterviewerProfile_Click

    Case "a53"     '所属プロフィール
        mnuInterviewGroupProfile_Click

    Case "a54"     '入試年度指定
        mnuSystemData_Click

    '----------------------------------------------------------------------------
    ' 印刷 Menu
    '----------------------------------------------------------------------------
    Case "a61"     '印刷指示
        mnuPrintCommand_Click

    Case "a62"     ' Excel帳票
        mnuExcelReport_Click

    Case "a63"     ' 度数分布図印刷
        mnuPrintDosu_Click

    '----------------------------------------------------------------------------
    ' データ出力
    '----------------------------------------------------------------------------
    Case "a71"      'ＣＳＶファイル出力
        mnuOutputCSV_Click


    '----------------------------------------------------------------------------
    ' 以下、未使用
    '----------------------------------------------------------------------------

''''    Case "a22"     ' special interview
''''        mnuSpecialInterview_Click
''''
''''    Case "a34"      '
''''        mnuPreparationRoom_Click
''''
''''    Case "a30"     ' adjust score at Shoronbun
''''        mnuAdjustScoreS_Click
''''
''''    Case "a36"     ' adjust score at Mensetsu
''''        mnuAdjustScoreM_Click
''''
''''    Case "a51"     ' High SChool Type
''''        mnuHighSchoolType_Click
''''
''''    Case "a52"     ' Zip Code
''''        mnuZipCode_Click
''''
''''    Case "a53"     ' Room Profile
''''        mnuRoomProfile_Click
''''
''''    Case "a54"     ' Subject Profile
''''        mnuSubjectProfile_Click
''''
''''    Case "a57"     ' Subject Question Profile
''''        mnuSubjectQuestionProfile_Click
''''
''''    Case "a62"     ' 成績一覧印刷指示
''''        mnuSeisekiIchiran_Click
''''
''''    Case "a71"      'データ転送
''''        mnuIVRTransfer_Click
''''
''''
'''''add,xzg,2010/12/09,S-----------
''''    Case "a73"
''''        mnuCommWork_Click
'''''add,xzg,2010/12/09,E-----------

    End Select

    Exit Sub

ErrorHandler:
    MsgBox Err.Description

End Sub

'*******************************************************************************
'* Menu Barで　表示されるのを設定する                                          *
'* Tree Menuに合わせた.[tool]-[Menu Editor]よりも設定できる                    *
'*-----------------------------------------------------------------------------*
'* 2021.12.03 update jhi                                                       *
'*******************************************************************************
Private Sub lsShowMenuBar(puMenues_() As prvuMenues_Type)

    On Error GoTo ErrorHandler

    Dim lLoopCnt As Long


    For lLoopCnt = 0 To UBound(puMenues_)

        Select Case puMenues_(lLoopCnt).sTVKey

        Case "nodeApplyPhase"     '
            mnuApplyPhase.Visible = puMenues_(lLoopCnt).bVisible

        Case "nodeFirstPhase"     '
            mnu1stExam.Visible = puMenues_(lLoopCnt).bVisible

        Case "nodeSecondPhase"     '
            mnu2ndExam.Visible = puMenues_(lLoopCnt).bVisible

        Case "nodeEnterRefuse"     '
            mnuEnterRefuse.Visible = puMenues_(lLoopCnt).bVisible

        Case "nodeMasters"     '
            mnuMaster.Visible = puMenues_(lLoopCnt).bVisible

        Case "nodePrint"     '
            mnuPrintMenu.Visible = puMenues_(lLoopCnt).bVisible

        Case "nodeTransfer"     '
            mnuTransfer.Visible = puMenues_(lLoopCnt).bVisible

        '***********************************************************************
        '* 出願受付フェーズ                                                    *
        '***********************************************************************
        Case "a01"     'Web出願データ取り込
            mnuOCR.Visible = puMenues_(lLoopCnt).bVisible

        Case "a02"     '受験生データ編集
            mnuMaintainExamineeData.Visible = puMenues_(lLoopCnt).bVisible

''''    Case "a03"    ' hyotei
''''        mnuHyotei.Visible = puMenues_(lLoopCnt).bVisible

        Case "a03"     'データ確定
            mnuFixData1.Visible = puMenues_(lLoopCnt).bVisible

        '***********************************************************************
        '* 一次試験                                                            *
        '***********************************************************************
        Case "a11"     '会場入力
            mnuRoomAllocation.Visible = puMenues_(lLoopCnt).bVisible

        Case "a12"     '欠席者入力
            mnuInputAbsenteeRecord.Visible = puMenues_(lLoopCnt).bVisible

        Case "a13"     '素点入力
            mnuInputRawScore.Visible = puMenues_(lLoopCnt).bVisible

''''    Case "a14"     ' input choosei score - grace
''''        mnuInputChooseiScore.Visible = puMenues_(lLoopCnt).bVisible
''''    Case "a15"     ' input choosei score - particular student
''''        mnuInputChooseiScore2.Visible = puMenues_(lLoopCnt).bVisible

        Case "a14"     '合格者入力
            mnuInputPassedPersonData.Visible = puMenues_(lLoopCnt).bVisible

        Case "a15"      '試験日振分
            mnuPreparationDay.Visible = puMenues_(lLoopCnt).bVisible

        Case "a16"     '試験日変更
            mnuManualAllocation.Visible = puMenues_(lLoopCnt).bVisible

'        Case "a19"     ' Manual Allocation
'            mnuPreparationRoom.Visible = puMenues_(lLoopCnt).bVisible

        Case "a17"     'データ確定
            mnuFixData2.Visible = puMenues_(lLoopCnt).bVisible


        '***********************************************************************
        '*  2次試験 Menu                                                       *
        '***********************************************************************
        Case "a21"     '欠席者入力
            mnuInputAbsenteeRecord2.Visible = puMenues_(lLoopCnt).bVisible

        Case "a22"     '面接委員登録
            mnuTeacherRoomMapInterview.Visible = puMenues_(lLoopCnt).bVisible

        Case "a23"     '面接グループ振分
            mnuPreparationRoom.Visible = puMenues_(lLoopCnt).bVisible

        Case "a24"     '面接グループ変更
            mnuManualAllocationGrp.Visible = puMenues_(lLoopCnt).bVisible

        Case "a25"     '小論文採点委員登
            mnuTeacherRoomMapReport.Visible = puMenues_(lLoopCnt).bVisible

        Case "a26"     '小論文振分
            mnuPreparationReport.Visible = puMenues_(lLoopCnt).bVisible

        Case "a27"     '素点入力(小論文)_import
            mnuImport_Syoronbun.Visible = puMenues_(lLoopCnt).bVisible

        Case "a28"     '素点入力(小論文)
            mnuInputRawScoreI.Visible = puMenues_(lLoopCnt).bVisible

        Case "a29"     '素点入力(面接)_import"
            mnuImport_Mensetu.Visible = puMenues_(lLoopCnt).bVisible

        Case "a30"     '素点入力(面接)
            mnuInputRawScore2.Visible = puMenues_(lLoopCnt).bVisible 'Menuで表示されないようにする2021.12.03 del jhi

        Case "a31"      '合格者入
            mnuInputPassedPersonData2.Visible = puMenues_(lLoopCnt).bVisible

        Case "a32"      '補欠者入力
            mnuWaitList2.Visible = puMenues_(lLoopCnt).bVisible

        Case "a33"      '補欠者順位
            mnuHoketusyaJuni.Visible = puMenues_(lLoopCnt).bVisible

        Case "a34"      'データ確定
            mnuFixData3.Visible = puMenues_(lLoopCnt).bVisible
 
''''    Case "a36"     ' adjust score at Mensetsu
''''        mnuAdjustScoreM.Visible = puMenues_(lLoopCnt).bVisible

        '***********************************************************************
        '* 入学手続き処理 Menu                                                 *
        '***********************************************************************
        Case "a41"     '補欠者合格者繰上げ処理
            mnuUpliftment.Visible = puMenues_(lLoopCnt).bVisible

        Case "a42"     '辞退
            mnuRefuseOffer.Visible = puMenues_(lLoopCnt).bVisible

        Case "a43"     'データ確定
            mnuFixData4.Visible = puMenues_(lLoopCnt).bVisible
        

    '***************************************************************************
    '* マスターメンテナンス Menu                                               *
    '***************************************************************************

        Case "a51"     '会場・面接グループ
            mnuRoomProfile.Visible = puMenues_(lLoopCnt).bVisible

        Case "a52"     '採点者プロファイル
            mnuInterviewerProfile.Visible = puMenues_(lLoopCnt).bVisible

        Case "a53"     '所属プロフィール
            mnuInterviewGroupProfile.Visible = puMenues_(lLoopCnt).bVisible

        Case "a54"     '入試年度指定
            mnuSystemData.Visible = puMenues_(lLoopCnt).bVisible

''''    Case "a55"     ' Interviewer Profile
''''        mnuInterviewerProfile.Visible = puMenues_(lLoopCnt).bVisible
''''
''''    Case "a56"     ' Interview Group Profile
''''        mnuInterviewGroupProfile.Visible = puMenues_(lLoopCnt).bVisible

''''    Case "a57"     ' Subject Question Profile
''''        mnuSubjectQuestionProfile.Visible = puMenues_(lLoopCnt).bVisible

''''    Case "a58"     ' Subject Question Profile
''''        mnuSystemData.Visible = puMenues_(lLoopCnt).bVisible


    '***************************************************************************
    '* 印刷 Menu                                                               *
    '***************************************************************************
        Case "a61"     ' 印刷指示
            mnuPrintCommand.Visible = puMenues_(lLoopCnt).bVisible

        Case "a62"     'Excel帳票
            mnuExcelReport.Visible = puMenues_(lLoopCnt).bVisible

        Case "a63"     '度数分布図印刷
            mnuPrintDosu.Visible = puMenues_(lLoopCnt).bVisible

    '***************************************************************************
    '* データ出力 Menu                                                         *
    '***************************************************************************

        Case "a71"      '受験生＋素点情報
            mnuOutputCSV.Visible = puMenues_(lLoopCnt).bVisible

        End Select

    Next

    Exit Sub

ErrorHandler:
    MsgBox Err.Description

End Sub


