VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H80000004&
   Caption         =   "�����V�X�e��"
   ClientHeight    =   8535
   ClientLeft      =   2280
   ClientTop       =   1740
   ClientWidth     =   12495
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0000
   Tag             =   "1905"
   WindowState     =   2  '�ő剻
   Begin VB.PictureBox pctExplorer 
      Align           =   3  '������
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
            Name            =   "�l�r �S�V�b�N"
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
      Align           =   1  '�㑵��
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
      Caption         =   "�����敪"
      Begin VB.Menu mnuExamZenki 
         Caption         =   "�O������"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "���j���["
      Begin VB.Menu mnuApplyPhase 
         Caption         =   "�菑��t�t�F�[�Y"
         Begin VB.Menu mnuOCR 
            Caption         =   "Web�o��f�[�^�捞"
         End
         Begin VB.Menu mnuMaintainExamineeData 
            Caption         =   "�󌱎҃f�[�^�̕ҏW"
         End
         Begin VB.Menu mnuExamineeCheck 
            Caption         =   "�󌱎ҏ�񃁃��e�i���X"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFixData1 
            Caption         =   "�f�[�^�m��"
         End
      End
      Begin VB.Menu mnu1stExam 
         Caption         =   "�ꎟ����"
         Begin VB.Menu mnuRoomAllocation 
            Caption         =   "������"
         End
         Begin VB.Menu mnuInputAbsenteeRecord 
            Caption         =   "���Ȏғ���"
         End
         Begin VB.Menu mnuInputRawScore 
            Caption         =   "�f�_����"
         End
         Begin VB.Menu mnuInputChooseiScore2 
            Caption         =   "�����ʒ����_����"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuInputChooseiScore 
            Caption         =   "�Ȗڕʒ����_����"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuInputPassedPersonData 
            Caption         =   "���i�ғ���"
         End
         Begin VB.Menu mnuPreparationDay 
            Caption         =   "�������U��"
         End
         Begin VB.Menu mnuManualAllocation 
            Caption         =   "�������ύX"
         End
         Begin VB.Menu mnuFixData2 
            Caption         =   "�f�[�^�m��"
         End
         Begin VB.Menu mnuMaintainExamineeData2 
            Caption         =   "�󌱎҃f�[�^�̕ҏW"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu2ndExam 
         Caption         =   "�j������"
         Begin VB.Menu mnuInputAbsenteeRecord2 
            Caption         =   "���Ȏғ���"
         End
         Begin VB.Menu mnuTeacherRoomMapInterview 
            Caption         =   "�ʐڈψ��o�^"
         End
         Begin VB.Menu mnuPreparationRoom 
            Caption         =   "�ʐڃO���[�v�U��"
         End
         Begin VB.Menu mnuManualAllocationGrp 
            Caption         =   "�ʐڃO���[�v�ύX"
         End
         Begin VB.Menu mnuTeacherRoomMapReport 
            Caption         =   "���_���̓_�ψ��o�^"
         End
         Begin VB.Menu mnuPreparationReport 
            Caption         =   "���_���U��"
         End
         Begin VB.Menu mnuImport_Syoronbun 
            Caption         =   "�f�_����(���_��)_import"
         End
         Begin VB.Menu mnuInputRawScoreI 
            Caption         =   "�f�_����(���_��)"
         End
         Begin VB.Menu mnuImport_Mensetu 
            Caption         =   "�f�_����(�ʐ�)_import"
         End
         Begin VB.Menu mnuInputRawScore2 
            Caption         =   "�f�_����(�ʐ�)"
         End
         Begin VB.Menu mnuInputPassedPersonData2 
            Caption         =   "���i�ғ���"
         End
         Begin VB.Menu mnuWaitList2 
            Caption         =   "�⌇�ғ���"
         End
         Begin VB.Menu mnuHoketusyaJuni 
            Caption         =   "�⌇�ҏ���"
         End
         Begin VB.Menu mnuFixData3 
            Caption         =   "�f�[�^�m��"
         End
         Begin VB.Menu mnuAdjustScoreM 
            Caption         =   "�����_����(�ʐ�)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAdjustScoreS 
            Caption         =   "�����_����(���_��)"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuEnterRefuse 
         Caption         =   "���w�葱������"
         Begin VB.Menu mnuUpliftment 
            Caption         =   "�⌇�ҍ��i�ҌJ�グ����"
         End
         Begin VB.Menu mnuRefuseOffer 
            Caption         =   "����"
         End
         Begin VB.Menu mnuFixData4 
            Caption         =   "�f�[�^�m��"
         End
      End
      Begin VB.Menu mnuMaster 
         Caption         =   "�}�X�^�[�����e�i���X"
         Begin VB.Menu mnuRoomProfile 
            Caption         =   "���E�ʐڃO���[�v"
         End
         Begin VB.Menu mnuInterviewerProfile 
            Caption         =   "�̓_�҃v���t�@�C��"
         End
         Begin VB.Menu mnuInterviewGroupProfile 
            Caption         =   "�����v���t�B�[��"
         End
         Begin VB.Menu mnuSystemData 
            Caption         =   "�����N�x�ݒ�"
         End
      End
      Begin VB.Menu mnuPrintMenu 
         Caption         =   "���"
         Begin VB.Menu mnuPrintCommand 
            Caption         =   "����w��"
         End
         Begin VB.Menu mnuExcelReport 
            Caption         =   "Excel���["
         End
         Begin VB.Menu mnuPrintDosu 
            Caption         =   "�x�����z�}���"
         End
      End
      Begin VB.Menu mnuTransfer 
         Caption         =   "�󌱃f�[�^CSV�o��"
         Begin VB.Menu mnuOutputCSV 
            Caption         =   "�󌱐��{�f�_���"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "�c�[��"
      Visible         =   0   'False
      Begin VB.Menu mnuToolsSearch 
         Caption         =   "���R�[�h��\��"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuToolsSave 
         Caption         =   "�ۑ�"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuToolsDelete 
         Caption         =   "�폜"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuToolsCancel 
         Caption         =   "�L�����Z��"
      End
      Begin VB.Menu mnuToolsNew 
         Caption         =   "�V�K"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuToolsQuery 
         Caption         =   "�N�G��"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTreeMenu 
      Caption         =   "�c���[���j���["
      Begin VB.Menu mnuShowTree 
         Caption         =   "���j���[�\��"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "���"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "�E�C���h�E"
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
      Caption         =   "�w���v"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuExit 
      Caption         =   "�I��"
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

''''Public f_int_CurrentPhase  As Integer       'modNyushi.bas�Ɉړ� 2021.12.28 del jhi

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

''''2021.12.28 del jhi global�ɐ錾����
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
        MsgBox "�A�v���P�[�V�����̏������Ɏ��s���܂����B���΂炭��A�Ď��s���Ă��������B", vbCritical, gTit
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
        MsgBox "�A�v���P�[�V�����̏������Ɏ��s���܂����B���΂炭��A�Ď��s���Ă��������B", vbCritical, gTit

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
''''Init_TreeView_New()�ɓ��� S
''''2021.12.28 del jhi
''''----------------------------------------------------------------------------

''''treeview�ɔw�i�F��ݒ�
''''Call SetTVBackColor(tvwMenu, RGB(230, 230, 250)) 'Lavender
''''Call SetTVBackColor(tvwMenu, RGB(240, 255, 240)) 'Honeydew
''''Call SetTVBackColor(tvwMenu, RGB(220, 220, 220)) 'Gainsboro
'   Call SetTVBackColor(tvwMenu, RGB(255, 250, 250)) 'Snow


    '---------------------------------------------------------------------------
    'Menu��S���L����
    '2021.12.22 add jhi
    '---------------------------------------------------------------------------
'    Dim objNode As Node
'    For Each objNode In tvwMenu.Nodes
''''''    If (objNode Is TreeView1.SelectedItem) Then
'            objNode.Expanded = True
''''''    End If
'    Next


''''----------------------------------------------------------------------------
''''Init_TreeView_New()�ɓ��� E
''''2021.12.28 add
''''----------------------------------------------------------------------------

''''----------------------------------------------------------------------------
''''�����t���R���p�C�������̐ݒ� 2022.02.01 add jhi
''''----------------------------------------------------------------------------
#If zengo_kubun = 1 Then
    mnuExamZenki.Caption = "�O������"
#Else
    mnuExamZenki.Caption = "�������"
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
'* �P�D�o���t�t�F�[�Y                                                        *
'*******************************************************************************
'*******************************************************************************
'* Web�o��f�[�^�捞                                                           *
'*******************************************************************************
Private Sub mnuOCR_Click()


    Unload frmBrowse

    frmBrowse.Caption = "frmBrowse : Web�o��f�[�^�捞"
    frmBrowse.Show

    frmBrowse.ZOrder 0

End Sub

'*******************************************************************************
'* �󌱐��f�[�^�̕ҏW                                                          *
'*******************************************************************************
Private Sub mnuMaintainExamineeData_Click()


    gbExamCheckNewShow = True ''''2021.12.22 add jhi


    If gbExamCheckNewShow Then
 
        ''''Unload frmExamCheckPara ''''2021.12.22 add jhi ''''�����Ȃ��ꍇ������̂�2023.01.24 del jhi
        frmExamCheckPara.Caption = "frmExamCheckPara : �󌱐��f�[�^�̕ҏW"
        frmExamCheckPara.Show

        ''''�R���g���[���� Z �I�[�_�[�̍őO�ʂɔz�u���܂��B �R���g���[�������̃R���g���[���̏�ɕ\������܂� (����l)�B
        frmExamCheckPara.ZOrder 0

    Else
        ''''frmExamineeCheck.ZOrder 0 ''''2023.01.24 �Ӗ����Ȃ��̂�comment out
    End If

End Sub

'*******************************************************************************
'* �f�[�^�m�� ����                                                             *
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
    Case 0 '�菑��t�t�F�[�Y
        sTmp = "�菑��t�t�F�[�Y���m�肵�܂��B��낵���ł����H"
        rinf = MsgBox(sTmp, vbYesNo + vbQuestion, gTit)

    Case 1 '�ꎟ����
        sTmp = "�ꎟ�����t�F�[�Y���m�肵�܂��B��낵���ł����H"
        rinf = MsgBox(sTmp, vbYesNo + vbQuestion, gTit)

    Case 2 '�񎟎���
        sTmp = "�񎟎����t�F�[�Y���m�肵�܂��B��낵���ł����H"
        rinf = MsgBox(sTmp, vbYesNo + vbQuestion, gTit)

    Case 3 '���w�葱������
        sTmp = "���w�葱���t�F�[�Y���m�肵�܂��B��낵���ł����H"
        rinf = MsgBox(sTmp, vbYesNo + vbQuestion, gTit)

    Case Else
        sTmp = "f_int_CurrentPhase error!"
        rinf = MsgBox(sTmp, vbYesNo + vbQuestion, gTit)

    End Select


    If rinf = vbYes Then

        If f_int_CurrentPhase <> 3 Then
            f_int_CurrentPhase = f_int_CurrentPhase + 1     ' ���̃t�F�[�Y�̃t���O���Z�b�g
        Else
            '�u���w�葱�������v����f�[�^�m�肷��Ɓu�菑��t�t�F�[�Y�v�ɃZ�b�g
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
            sTmp = "���w�葱���t�F�[�Y���m�肵�܂��B�菑�󂯕t�� �t�F�[�Y�ɖ߂�܂��B"
            MsgBox sTmp, vbInformation, gTit

        Case 1
            sTmp = "�菑�󂯕t���t�F�[�Y�̃f�[�^���m�肵�܂����B�ꎟ�����t�F�[�Y����͂��Ă��������B"
            MsgBox sTmp, vbInformation, gTit

        Case 2
            sTmp = "�ꎟ�����t�F�[�Y�̃f�[�^���m�肵�܂����B�񎟎����t�F�[�Y����͂��Ă��������B"
            MsgBox sTmp, vbInformation, gTit

        Case 3
            sTmp = "�񎟎����t�F�[�Y�̃f�[�^���m�肵�܂����B���w�葱���t�F�[�Y����͂��Ă��������B"
            MsgBox sTmp, vbInformation, gTit

        Case Else
            sTmp = "���̓t�F�[�Y�̐ݒ�t���O�ُ�ł��B�����𒆒f���܂��B"
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
        sTmp = "�����t�F�[�Y�̃t���O���Z�b�g���鏈���ŃG���[���������܂����B"
        MsgBox sTmp, vbInformation, gTit
    Else
        MsgBox Err.Number & vbCrLf & Err.Description
    End If

End Sub

'*******************************************************************************
'* �Q�D�P������                                                                *
'*******************************************************************************
'*******************************************************************************
'* ������                                                                    *
'*******************************************************************************
Private Sub mnuRoomAllocation_Click()

    Unload frmRoomAlloc

    frmRoomAlloc.Caption = "frmRoomAlloc : ������"
    frmRoomAlloc.Show

    frmRoomAlloc.ZOrder 0

End Sub

'*******************************************************************************
'* ���Ȏғ���                                                                  *
'*******************************************************************************
Private Sub mnuInputAbsenteeRecord_Click()


''''    If f_int_CurrentPhase <> 1 Then
''''        MsgBox "1�t�F�[�Y�����킹�ĉ�ʕ\�����s���Ă��������B"
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
'        .Caption = "frmAbsenteeRecord : ���Ȏғ���"
'        .ZOrder 0
'    End With

'2021.12.29 add jhi
    frm1jikesseki.m_int_IntRpt = 1
    frm1jikesseki.m_int_Action = 0
    frm1jikesseki.Caption = "frm1jikesseki : 1�� ���Ȏғ���"

    frm1jikesseki.Show
    frm1jikesseki.ZOrder 0



End Sub

'*******************************************************************************
'* 1���f�_����                                                                 *
'*******************************************************************************
Private Sub mnuInputRawScore_Click()

    g_int_ExamType = 1 '1�������t�F�[�Y��ݒ� 2021.12.22 add jhi


''''2022.01.24 form�ύX�ɂ�� del jhi
''''    Unload frmRawScore
''''
''''    frmRawScore.Caption = "frmRawScore : �f�_����"
''''    frmRawScore.Show
''''
''''    frmRawScore.ZOrder 0


''''2022.01.24 form�ύX�ɂ�� add jhi
    Unload frm1jiSotenInput

    frm1jiSotenInput.Caption = "frm1jiSotenInput : �f�_����"
    frm1jiSotenInput.Show

    frm1jiSotenInput.ZOrder 0


End Sub

'*******************************************************************************
'* ���i�ғ���                                                                  *
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
'        .Caption = "frmPassedPersonData : ���i�ғ���"
'        .ZOrder 0
'    End With


'2021.12.29 add jhi
    frm1jigoukaku.m_int_Action = 1
    frm1jigoukaku.m_int_IntRpt = 1
    frm1jigoukaku.Caption = "frm1jigoukaku : 1�� ���i�ғ���"
    frm1jigoukaku.Show
    frm1jigoukaku.ZOrder 0


End Sub

'*******************************************************************************
'* �������U��                                                                  *
'*******************************************************************************
Private Sub mnuPreparationDay_Click()

    Dim strMsg As String

''''2022.03.09 add jhi ���n��
#If zengo_kubun = 1 Then
    strMsg = "frmPrepSecondExam : �񎟎������U��"
#Else
    strMsg = "frmPrepSecondExam : �񎟎������m��"
#End If



    frmPrepSecondExam.Show
    frmPrepSecondExam.Caption = strMsg
    frmPrepSecondExam.ZOrder 0

End Sub

'*******************************************************************************
'* �������ύX                                                                  *
'*******************************************************************************
Private Sub mnuManualAllocation_Click()

    frmManualAllocation.Show
    frmManualAllocation.Caption = "frmManualAllocation : �񎟎������ύX"
    frmManualAllocation.ZOrder 0

End Sub

'*******************************************************************************
'* �f�[�^�m��                                                                  *
'*******************************************************************************
Private Sub mnuFixData2_Click()

    Call mnuFixData1_Click

End Sub

'*******************************************************************************
'* �R�D�Q������                                                                *
'*******************************************************************************
'*******************************************************************************
'* 2������ : ���Ȏғ���                                                        *
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
        .Caption = "frmExamineeStatus : 2�� ���Ȏғ���"
        .ZOrder 0
    End With
#End If


    '2021.12.29 add jhi
    frm2jikesseki.m_int_IntRpt = 0
    frm2jikesseki.m_int_Action = 2

    frm2jikesseki.Caption = "frm2jikesseki : 2�� ���Ȏғ���"
    frm2jikesseki.Show
    frm2jikesseki.ZOrder 0



End Sub

'*******************************************************************************
'* 2������ : �ʐڈψ��o�^                                                      *
'*******************************************************************************
Private Sub mnuTeacherRoomMapInterview_Click()

    '�ʐڈψ��o�^
    If frmIntwrRoomMapInt Is Nothing Then
        Set frmIntwrRoomMapInt = New frmInterviewerRoom
    End If

    With frmIntwrRoomMapInt
        .m_int_IntRpt = 0
        .Show
        .Caption = "frmInterviewerRoom : �ʐڈψ��o�^"   ''''LoadResString(2301)
        .ZOrder 0
    End With

End Sub

'*******************************************************************************
'* 2������ : �ʐڃO���[�v�U��                                                  *
'*******************************************************************************
Private Sub mnuPreparationRoom_Click()

    frmPrepSecondExamGrp.Show
    frmPrepSecondExamGrp.Caption = "frmPrepSecondExamGrp : �ʐڃO���[�v�U��"
    frmPrepSecondExamGrp.ZOrder 0

End Sub

'*******************************************************************************
'* 2������ : �ʐڃO���[�v�ύX                                                  *
'*******************************************************************************
Private Sub mnuManualAllocationGrp_Click()

    frmManualAllocationGrp.Show
    frmManualAllocationGrp.Caption = "frmManualAllocationGrp : �ʐڃO���[�v�ύX"
    frmManualAllocationGrp.ZOrder 0

End Sub

'*******************************************************************************
'* 2������ : ���_���̓_�ψ��o�^                                                *
'*******************************************************************************
Private Sub mnuTeacherRoomMapReport_Click()

    ' Teacher-Room Mapping - Report
    If frmIntwrRoomMapRpt Is Nothing Then
        Set frmIntwrRoomMapRpt = New frmInterviewerReport
    End If

    With frmIntwrRoomMapRpt
        .m_int_IntRpt = 1
        .Show
''''    .Caption = LoadResString(2302) '�̓_��-���_����ꊄ�蓖��-
        .Caption = "frmInterviewerReport : ���_���̓_�ψ��o�^"  ''''LoadResString(2302) '�̓_��-���_����ꊄ�蓖��-
        .ZOrder 0
    End With

End Sub

'*******************************************************************************
'* 2������ : ���_���U��                                                        *
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
'* 2������ : �f�_����(���_��)_import                                           *
'* 2021.12.12 add jhi                                                          *
'*******************************************************************************
Private Sub mnuImport_Syoronbun_Click()

    g_int_ExamType = 2

''''MsgBox "�f�_����(���_��)_import��ʕ\��"

    Call frmImportSyoronbun.gsSetSecondType(1) '1:���_��

    frmImportSyoronbun.Show
    frmImportSyoronbun.Caption = "frmImportSyoronbun : �f�_����(���_��)_import "
    frmImportSyoronbun.ZOrder 0


End Sub

'*******************************************************************************
'* 2������ : �f�_����(���_��)                                                  *
'*******************************************************************************
Private Sub mnuInputRawScoreI_Click()

    ' raw score for second phase - report
    g_int_ExamType = 2

    If frmRawScoreRpt Is Nothing Then
        Set frmRawScoreRpt = New frmRawScore
    End If

    With frmRawScoreRpt
        Call .gsSetSecondType(1)    '1:���_��
        .Show
''''    .Caption = LoadResString(1019) ''''�f�_���́i���_���j
        .Caption = "frmRawScore : �f�_����(���_��)"
        .ZOrder 0
    End With

End Sub

'*******************************************************************************
'* 2������ : �f�_����(�ʐ�)_import                                             *
'* 2021.12.12 add jhi                                                          *
'*******************************************************************************
Private Sub mnuImport_Mensetu_Click()

    g_int_ExamType = 2


    Call frmImportMensetu.gsSetSecondType(0) '0:�ʐ�

    frmImportMensetu.Show
    frmImportMensetu.Caption = "frmImportMensetu : �f�_����(�ʐ�)_import "
    frmImportMensetu.ZOrder 0

End Sub

'*******************************************************************************
'* 2������ : �f�_����(�ʐ�)                                                    *
'*******************************************************************************
Private Sub mnuInputRawScore2_Click()

    ' raw score for second phase - interview
    g_int_ExamType = 2

    If frmRawScoreInt Is Nothing Then
        Set frmRawScoreInt = New frmRawScore
    End If

    With frmRawScoreInt
        Call .gsSetSecondType(0)    '0:�ʐ�
        .Show
''''    .Caption = LoadResString(1047)
        .Caption = "frmRawScore : �f�_����(�ʐ�)"
        .ZOrder 0
    End With

End Sub

'*******************************************************************************
'* 2������ : ���i�ғ���                                                        *
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
        .Caption = "frmExamineeStatus : 2�� ���i�ғ���"
        .ZOrder 0
    End With

#End If

'2021.12.30 add jhi
    frm2jigoukaku.m_int_IntRpt = 3
    frm2jigoukaku.m_int_Action = 3

    frm2jigoukaku.Caption = "frm2jigoukaku : 2�� ���i�ғ���"
    frm2jigoukaku.Show
    frm2jigoukaku.ZOrder 0


End Sub

'*******************************************************************************
'* 2������ : �⌇�ғ���                                                        *
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
        .Caption = "frmExamineeStatus : �⌇�ғ���"
        .ZOrder 0
    End With


#End If

'2021.12.30 add jhi
    frm2jiHoketusya.m_int_IntRpt = 4
    frm2jiHoketusya.m_int_Action = 4

    frm2jiHoketusya.Caption = "frm2jiHoketusya : 2�� �⌇�ғ���"
    frm2jiHoketusya.Show
    frm2jiHoketusya.ZOrder 0


End Sub

'*******************************************************************************
'* �⌇�ҏ���                                                                  *
'*******************************************************************************
'*******************************************************************************
'* 3.10 �⌇�ҏ��� --->�ق���subsystem ��ʂ��炱����ɓ����K�v�����遚      *
'* 2021.12.02 add jhi                                                          *
'*******************************************************************************
Private Sub mnuHoketusyaJuni_Click()

    '2������
    g_int_ExamType = 2

    'frmChooseiReport.Show
    'frmChooseiReport.ZOrder 0

    frm2jiHoketusyaJuni.Caption = "frmHoketusyaJuni : 2�� �⌇�ҏ���"
    frm2jiHoketusyaJuni.Show
    frm2jiHoketusyaJuni.ZOrder 0

End Sub

'*******************************************************************************
'* �f�[�^�m��                                                                  *
'*******************************************************************************
Private Sub mnuFixData3_Click()

    Call mnuFixData1_Click

End Sub

'*******************************************************************************
'* �S�D���w�葱������                                                          *
'*******************************************************************************
'*******************************************************************************
'* �⌇�ҍ��i�ҌJ�グ����                                                      *
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
        .Caption = "frmExamineeKuriage : �⌇�ҍ��i�J�グ����"
        .ZOrder 0
    End With
#End If

    frmExamineeKuriage.m_int_IntRpt = 5
    frmExamineeKuriage.m_int_Action = 5

    frmExamineeKuriage.Caption = "frmExamineeKuriage : �⌇�ҍ��i�J�グ����"
    frmExamineeKuriage.Show
    frmExamineeKuriage.ZOrder 0

End Sub

'*******************************************************************************
'* ����                                                                        *
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
        .Caption = "frmExamineeKuriage : ����"
        .ZOrder 0
    End With
#End If

    frmExamineeJitai.m_int_IntRpt = 6
    frmExamineeJitai.m_int_Action = 6

    frmExamineeJitai.Caption = "frmExamineeJitai : ����"
    frmExamineeJitai.Show
    frmExamineeJitai.ZOrder 0


End Sub

'*******************************************************************************
'* �f�[�^�m��                                                                  *
'*******************************************************************************
Private Sub mnuFixData4_Click()

    Call mnuFixData1_Click

End Sub

'*******************************************************************************
'* �T�D�}�X�^�[�����e�i���X Menu                                               *
'*******************************************************************************
'*******************************************************************************
'* ���E�ʐڃO���[�v                                                          *
'*******************************************************************************
Private Sub mnuRoomProfile_Click()

    frmRoomProfile.Show
    frmRoomProfile.ZOrder 0

End Sub

'*******************************************************************************
'* �̓_�҃v���t�@�C��                                                          *
'*******************************************************************************
Private Sub mnuInterviewerProfile_Click()

    frmInterviewerProfile.Show
    frmInterviewerProfile.ZOrder 0

End Sub

'*******************************************************************************
'* �����v���t�@�C��                                                            *
'*******************************************************************************
Private Sub mnuInterviewGroupProfile_Click()

    frmInterviewGroupProfile.Show
    frmInterviewGroupProfile.ZOrder 0

End Sub

'*******************************************************************************
'* �����N�x�w��                                                                *
'*******************************************************************************
Private Sub mnuSystemData_Click()

    frmSystemData.Show
    frmSystemData.ZOrder 0

End Sub

'*******************************************************************************
'* �U�D��� Menu                                                               *
'*******************************************************************************
'*******************************************************************************
'* ����w��(a61)                                                               *
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
'* Excel���[(a62)                                                              *
'*******************************************************************************
Private Sub mnuExcelReport_Click()

    frmExcelReport.Show
    frmExcelReport.ZOrder 0

End Sub

'*******************************************************************************
'* �x�����z�}���(a63)                                                          *
'*******************************************************************************
Private Sub mnuPrintDosu_Click()

    frmPrintDosu.Show
    frmPrintDosu.ZOrder 0

End Sub

'*******************************************************************************
'* �V�D�f�[�^�o��                                                              *
'*******************************************************************************
'*******************************************************************************
'* �󌱐��A�f�_���                                                            *
'*******************************************************************************



'*******************************************************************************
'* �ȉ��A���g�p                                                                *
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


    rinf = myMsgBox("�����V�X�e�����I�����܂��B��낵���ł����H", gTit)
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
'* �]��                                                                        *
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
'* ���_������  <----�g��Ȃ��悤��2021.12.22 �m�F                              *
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
'* TreeView�̃��j���[��\������                                                *
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

        ''''2021.12.09 add jhi Tree���j���[�Ɏ����敪������
        Set l_obj_NewNode = .Nodes.Add(, , "mnuExamKubun", "�y �O������ �z" & f_int_CurrentPhase) '�O������

        For lCnt = LBound(puMenues_) To UBound(puMenues_)

            If puMenues_(lCnt).bVisible Then
                If puMenues_(lCnt).lParent = -1 Then
                '�e
                    Set l_obj_NewNode = .Nodes.Add(, , puMenues_(lCnt).sTVKey, puMenues_(lCnt).sCaption)
                Else
                '�q
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
'* TreeView�̃��j���[��\������                                                *
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
''''�����t���R���p�C�������̐ݒ� 2022.02.01 add jhi
''''----------------------------------------------------------------------------
#If zengo_kubun = 1 Then

        ''''2021.12.09 add jhi Tree���j���[�Ɏ����敪������
        Set l_obj_NewNode = .Nodes.Add(, , "mnuExamKubun", "���O������ -  " & f_int_CurrentPhase + 1 & " �t�F�[�Y") '�t�F�[�Y�\��

#Else
        Set l_obj_NewNode = .Nodes.Add(, , "mnuExamKubun", "��������� -  " & f_int_CurrentPhase + 1 & " �t�F�[�Y") '�t�F�[�Y�\��
#End If

        For lCnt = LBound(puMenues_) To UBound(puMenues_)

            If puMenues_(lCnt).bVisible Then
                If puMenues_(lCnt).lParent = -1 Then
                '�e
                    Set l_obj_NewNode = .Nodes.Add(, , puMenues_(lCnt).sTVKey, puMenues_(lCnt).sCaption)
                Else
                '�q
                    If puMenues_(puMenues_(lCnt).lParent).bVisible Then
                        Set l_obj_NewNode = .Nodes.Add(puMenues_(puMenues_(lCnt).lParent).sTVKey, tvwChild, puMenues_(lCnt).sTVKey, puMenues_(lCnt).sCaption)
                    End If
                End If
            End If

        Next

    End With

step_no = 2


''''----------------------------------------------------------------------------
''''�����t���R���p�C�������̐ݒ� 2022.02.01 add jhi
''''----------------------------------------------------------------------------
#If zengo_kubun = 1 Then

''''Call SetTVBackColor(tvwMenu, RGB(240, 248, 255)) 'aliceblue
    Call SetTVBackColor(tvwMenu, RGB(255, 250, 250)) 'Snow

#Else

''''Call SetTVBackColor(tvwMenu, RGB(255, 250, 250)) 'Snow     2022.01.06 del jhi
    Call SetTVBackColor(tvwMenu, RGB(249, 207, 98))  'Yellow�n 2022.01.06 add jhi
#End If


    '---------------------------------------------------------------------------
    'Menu��S���L����
    '2021.12.22 add jhi
    '---------------------------------------------------------------------------
    Dim objNode As Node
    For Each objNode In tvwMenu.Nodes
        objNode.Expanded = True
    Next



    Exit Sub


ErrorHandler:

step_no = 3
    MsgBox "Init_TreeView_New�֐��ŃG���[���������܂����B"


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
'* Tree Menu ������ set                                                        *
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
    ReDim puMenues_(41)    ''''2021.12.21 update �폜Menu�Ή�


''''lCnt = 0

''''Set puMenues_(lCnt).oMnuObj = mnuExamKubun
''''puMenues_(lCnt).sTVKey = "nodeExamKubun"
''''puMenues_(lCnt).lParent = -1
''''puMenues_(lCnt).sCaption = "�O������"

''''If g_int_ExamKubun = 1 Then
''''    puMenues_(lCnt).sCaption = "�O������"
''''Else
''''    puMenues_(lCnt).sCaption = "�������"
''''End If

    
    lCnt = 0

    Set puMenues_(lCnt).oMnuObj = mnuApplyPhase
    puMenues_(lCnt).sTVKey = "nodeApplyPhase"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "1. �菑��t�t�F�[�Y"     ''''LoadResString(1002) '�菑��t�t�F�[�Y

    lCnt = lCnt + 1 '1
    Set puMenues_(lCnt).oMnuObj = mnu1stExam
    puMenues_(lCnt).sTVKey = "nodeFirstPhase"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "2. �ꎟ����"             ''''LoadResString(1008) '�ꎟ����

    lCnt = lCnt + 1 '2
    Set puMenues_(lCnt).oMnuObj = mnu2ndExam
    puMenues_(lCnt).sTVKey = "nodeSecondPhase"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "3. �񎟎���"             ''''LoadResString(1016) '�񎟎���

    lCnt = lCnt + 1 '3
    Set puMenues_(lCnt).oMnuObj = mnuEnterRefuse
    puMenues_(lCnt).sTVKey = "nodeEnterRefuse"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "4. ���w�葱������"       ''''LoadResString(1024) '���w�葱������

    lCnt = lCnt + 1 '4
    Set puMenues_(lCnt).oMnuObj = mnuMaster
    puMenues_(lCnt).sTVKey = "nodeMasters"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "5. �}�X�^�[�����e�i���X" ''''LoadResString(1028) '�}�X�^�[�����e�i���X

    lCnt = lCnt + 1 '5
    Set puMenues_(lCnt).oMnuObj = mnuPrint
    puMenues_(lCnt).sTVKey = "nodePrint"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "6. ���"                 ''''LoadResString(1090) '���

    lCnt = lCnt + 1 '6
    Set puMenues_(lCnt).oMnuObj = mnuTransfer
    puMenues_(lCnt).sTVKey = "nodeTransfer"
    puMenues_(lCnt).lParent = -1
    puMenues_(lCnt).sCaption = "7. �󌱃f�[�^CSV�o��"        ''''LoadResString(1091) '�f�[�^�o��


    '***************************************************************************
    '* �菑��t�t�F�[�Y Menu                                                   *
    '***************************************************************************
    lCnt = lCnt + 1 '7
    puMenues_(lCnt).sTVKey = "a01"
    puMenues_(lCnt).lParent = 0
    puMenues_(lCnt).sCaption = "Web�o��f�[�^�捞"       ''''LoadResString(1003)

    lCnt = lCnt + 1 '8
    puMenues_(lCnt).sTVKey = "a02"
    puMenues_(lCnt).lParent = 0
    puMenues_(lCnt).sCaption = "�󌱐��f�[�^�̕ҏW"      ''''LoadResString(1004)

    lCnt = lCnt + 1 '9
    puMenues_(lCnt).sTVKey = "a03"
    puMenues_(lCnt).lParent = 0
    puMenues_(lCnt).sCaption = "�f�[�^�m��"              ''''LoadResString(1007)

''''2021.12.01 del jhi �]��
''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a03"
'''' puMenues_(lCnt).lParent = 0
''''puMenues_(lCnt).sCaption = "�]��"                    ''''LoadResString(1005)


    '***************************************************************************
    '* 1������ Menu                                                            *
    '***************************************************************************
    lCnt = lCnt + 1 '10
    puMenues_(lCnt).sTVKey = "a11"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "������"                ''''LoadResString(1009) '������

    lCnt = lCnt + 1 '11
    puMenues_(lCnt).sTVKey = "a12"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "���Ȏғ���"              ''''LoadResString(1010) '���Ȏғ���

    lCnt = lCnt + 1 '12
    puMenues_(lCnt).sTVKey = "a13"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "�f�_����"                ''''LoadResString(1011) '�f�_���� <---import�͂��Ȃ�

    lCnt = lCnt + 1 '13
    puMenues_(lCnt).sTVKey = "a14"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "���i�ғ���"              ''''LoadResString(1013) '���i�ғ���

    lCnt = lCnt + 1 '14
    puMenues_(lCnt).sTVKey = "a15"
    puMenues_(lCnt).lParent = 1
''''2022.03.09 del jhi
''''puMenues_(lCnt).sCaption = "�񎟎������U��"           ''''LoadResString(1080) '�񎟎������U��

#If zengo_kubun = 1 Then
    strMsg = "�񎟎������U��"
#Else
    strMsg = "�񎟎������m��"
#End If

    puMenues_(lCnt).sCaption = strMsg                     ''''2022.03.09 add jhi �O���A�����Titile��ύX


    lCnt = lCnt + 1 '15
    puMenues_(lCnt).sTVKey = "a16"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "�񎟎������ύX"           ''''LoadResString(1081) '�񎟎������ύX

    lCnt = lCnt + 1 '16
    puMenues_(lCnt).sTVKey = "a17"
    puMenues_(lCnt).lParent = 1
    puMenues_(lCnt).sCaption = "�f�[�^�m��"              ''''LoadResString(1007) '�f�[�^�m��


''''2021.12.01 del jhi
''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a14"
''''puMenues_(lCnt).lParent = 1
''''puMenues_(lCnt).sCaption = "�Ȗڕʒ����_����"        ''''LoadResString(1012) '�Ȗڕʒ����_����

    'del,xzg,2009/12/02,S----------
    'lCnt = 14      'xx
    'puMenues_(lCnt).sTVKey = "a15"
    'puMenues_(lCnt).lParent = 1
    'puMenues_(lCnt).sCaption = "�����ʒ����_����"       ''''LoadResString(1046) '�����ʒ����_����
    'del,xzg,2009/12/02,E----------

  
    '***************************************************************************
    '* 2������ Menu                                                            *
    '***************************************************************************

    lCnt = lCnt + 1 '17
    puMenues_(lCnt).sTVKey = "a21"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "���Ȏғ���"    ''''LoadResString(1018)
    
    '---------------------------------------------------------------------------
    ' �ʐڊ֘A 3Menu
    '---------------------------------------------------------------------------
    lCnt = lCnt + 1 '18
    puMenues_(lCnt).sTVKey = "a22"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "�ʐڈψ��o�^"     ''''LoadResString(1051)

    lCnt = lCnt + 1 '19
    puMenues_(lCnt).sTVKey = "a23"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "�ʐڃO���[�v�U��" ''''LoadResString(1082)

    lCnt = lCnt + 1 '20
    puMenues_(lCnt).sTVKey = "a24"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "�ʐڃO���[�v�ύX" ''''LoadResString(1083)
'-------------------------------------------------------------------------------

    lCnt = lCnt + 1 '21
    puMenues_(lCnt).sTVKey = "a25"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "���_���̓_�ψ��o�^"    ''''LoadResString(1053)

    lCnt = lCnt + 1 '22 ���_���U��
    puMenues_(lCnt).sTVKey = "a26"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "���_���U��"    ''''LoadResString(2433)


    '---------------------------------------------------------------------------
    ''''2021.12.12 add jhi
    lCnt = lCnt + 1 '23
    puMenues_(lCnt).sTVKey = "a27"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "�f�_����(���_��)_import"

    lCnt = lCnt + 1 '24 �f�_����(���_��)
    puMenues_(lCnt).sTVKey = "a28"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "�f�_����(���_��)"        ''''LoadResString(1019)
    '---------------------------------------------------------------------------

    lCnt = lCnt + 1 '25
    puMenues_(lCnt).sTVKey = "a29"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "�f�_����(�ʐ�)_import"

    lCnt = lCnt + 1 '26
    puMenues_(lCnt).sTVKey = "a30"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "�f�_����(�ʐ�)"          ''''LoadResString(1047)

    lCnt = lCnt + 1 '27
    puMenues_(lCnt).sTVKey = "a31"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "���i�ғ���"              ''''LoadResString(1021)

    lCnt = lCnt + 1 '28
    puMenues_(lCnt).sTVKey = "a32"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "�⌇�ғ���"              ''''LoadResString(1022)

    lCnt = lCnt + 1 '29
    puMenues_(lCnt).sTVKey = "a33"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "�⌇�ҏ���"              ''''��subsystem���瓱������

    lCnt = lCnt + 1 '30
    puMenues_(lCnt).sTVKey = "a34"
    puMenues_(lCnt).lParent = 2
    puMenues_(lCnt).sCaption = "�f�[�^�m��"              ''''LoadResString(1007)


    'add,xzg,2010/12/09,S-----------
    '���_������
''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a27"
''''puMenues_(lCnt).lParent = 2
''''puMenues_(lCnt).sCaption = "���_������"           '<---�\������Ȃ�
    'add,xzg,2010/12/09,E-----------

''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a28"
''''puMenues_(lCnt).lParent = 2
''''puMenues_(lCnt).sCaption = "�Q���ʐڃO���[�v����" '<---�\������Ȃ�

'-------------------------------------------------------------------------------
' 2021.12.02 del jhi
'-------------------------------------------------------------------------------
'    lCnt = lCnt + 1 'xx �����_����(���_��)
'    puMenues_(lCnt).sTVKey = "a30"
'    puMenues_(lCnt).lParent = 2
'    puMenues_(lCnt).sCaption = LoadResString(1048)
'
'    lCnt = lCnt + 1 'xx �����_����(�ʐ�)
'    puMenues_(lCnt).sTVKey = "a36"
'    puMenues_(lCnt).lParent = 2
'    puMenues_(lCnt).sCaption = LoadResString(1049)
'-------------------------------------------------------------------------------


    '***************************************************************************
    '* ���w�葱������ Menu                                                     *
    '***************************************************************************

    lCnt = lCnt + 1 '31
    puMenues_(lCnt).sTVKey = "a41"
    puMenues_(lCnt).lParent = 3
    puMenues_(lCnt).sCaption = "�⌇�ҍ��i�J�グ����"    ''''LoadResString(1025)

    lCnt = lCnt + 1 '32
    puMenues_(lCnt).sTVKey = "a42"
    puMenues_(lCnt).lParent = 3
    puMenues_(lCnt).sCaption = "����"                    ''''LoadResString(1026)

    lCnt = lCnt + 1 '33
    puMenues_(lCnt).sTVKey = "a43"
    puMenues_(lCnt).lParent = 3
    puMenues_(lCnt).sCaption = "�f�[�^�m��"              ''''LoadResString(1007)


    '***************************************************************************
    '* �}�X�^�[�����e�i���X Menu                                               *
    '***************************************************************************

    lCnt = lCnt + 1 '34
    puMenues_(lCnt).sTVKey = "a51"
    puMenues_(lCnt).lParent = 4
    puMenues_(lCnt).sCaption = "���E�ʐڃO���[�v"     ''''LoadResString(1031)

    lCnt = lCnt + 1 '35
    puMenues_(lCnt).sTVKey = "a52"
    puMenues_(lCnt).lParent = 4
    puMenues_(lCnt).sCaption = "�̓_�҃v���t�@�C��"      ''''LoadResString(1033)

    lCnt = lCnt + 1 '36
    puMenues_(lCnt).sTVKey = "a53"
    puMenues_(lCnt).lParent = 4
    puMenues_(lCnt).sCaption = "�����v���t�B�[��"        ''''LoadResString(2466)

    lCnt = lCnt + 1 '37
    puMenues_(lCnt).sTVKey = "a54"
    puMenues_(lCnt).lParent = 4
    puMenues_(lCnt).sCaption = "�����N�x�w��"            ''''LoadResString(2600) '�V�X�e���p�����[�^


''''2021.12.21 del jhi S====
''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a55"
''''puMenues_(lCnt).lParent = 4
''''puMenues_(lCnt).sCaption = "���Z�敪"                ''''LoadResString(1029)

''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a56"
''''puMenues_(lCnt).lParent = 4
''''puMenues_(lCnt).sCaption = "�X�֔ԍ� ID"             ''''LoadResString(1030)
''''2021.12.21 del jhi E====

''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a57"
''''puMenues_(lCnt).lParent = 4
''''puMenues_(lCnt).sCaption = "�Ȗڃv���t�@�C��"        ''''LoadResString(1032)

''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a58"
''''puMenues_(lCnt).lParent = 4
''''puMenues_(lCnt).sCaption = "�Ȗږ��v���t�@�C��"    ''''LoadResString(2458)


    '***************************************************************************
    '* ��� Menu                                                               *
    '***************************************************************************
    lCnt = lCnt + 1 '38
    puMenues_(lCnt).sTVKey = "a61"
    puMenues_(lCnt).lParent = 5
    puMenues_(lCnt).sCaption = "����w��"                ''''LoadResString(1092)

    lCnt = lCnt + 1 '39
    puMenues_(lCnt).sTVKey = "a62"
    puMenues_(lCnt).lParent = 5
    puMenues_(lCnt).sCaption = "Excel���["               ''''LoadResString(1093)

    lCnt = lCnt + 1 '40
    puMenues_(lCnt).sTVKey = "a63"
    puMenues_(lCnt).lParent = 5
    puMenues_(lCnt).sCaption = "�x�����z�}���"          '''''LoadResString(2700)

''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a64"
''''puMenues_(lCnt).lParent = 5
''''puMenues_(lCnt).sCaption = "���шꗗ"                ''''LoadResString(1094)


    '***************************************************************************
    '* �f�[�^�o�� Menu                                                         *
    '***************************************************************************

    lCnt = lCnt + 1 '41
    puMenues_(lCnt).sTVKey = "a71"
    puMenues_(lCnt).lParent = 6
    puMenues_(lCnt).sCaption = "�󌱐��{�f�_���"        ''''LoadResString(1096)


''''2021.11.30 del IVR�V�X�e���ւ̃f�[�^�]��
''''lCnt = lCnt + 1 'xx
''''puMenues_(lCnt).sTVKey = "a71"
''''puMenues_(lCnt).lParent = 6
''''puMenues_(lCnt).sCaption = LoadResString(1095)


''''Debug.Print "�z��: lCnt=" & lCnt


    '***************************************************************************
    '* �ݒ� Menu Key�@���e��Buffer�ɐݒ�                                       *
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
'* ini file �� MENU Section ����                                               *
'* 2022.02.01 update jhi                                                       *
'*******************************************************************************
Private Sub SetPhaseMenu(f_int_CurrentPhase As Long)

    On Error GoTo ERR_HANDLE

    Dim oRs          As ADODB.Recordset
    Dim sSQL         As String

'���[�U�AMAC�A�h���X�A�Ɩ�PHASE���\�����郁�j���[�����肷��

    'MAC�A�h���X�̎擾
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

''''2021.12.28 del jhi global�ɐ錾
''''Dim uMenues_() As prvuMenues_Type

    Dim sUserPass    As String
    Dim sMacPass     As String
    Dim sMenuPass    As String
    Dim sMenuGPass   As String
    Dim sIniPass     As String


    lAdptCnt = mAdptInf.gfLoadAdptData(sErrMsg)

    If lAdptCnt < 1 Then
        MsgBox "�l�`�b�A�h���X�̎擾�Ɏ��s���܂����B" & vbCrLf & sErrMsg, vbOKOnly, "���������s��"
        End
    End If

    '���[�U�AMAC�A�h���X���\���\���j���[�������t�@�C���̃Z�N�V���������擾����
    '�ꉞ�A�A�_�v�^�̐��������[�v����悤�ɂ��Ă����i�Q������������̂Łj
    sMenuString = ""
    Set oGao = CreateObject("GaoEncode.GaoeAPI")

    For lCnt = 0 To lAdptCnt - 1

        sMacAddr = Replace(mAdptInf.getMacAddr(lCnt), "-", "")
Call log("1-----> sMacAddr=" & sMacAddr)


        '���[�U�h�c���Í���
        sUserPass = GetSetting("Nyushi", "Settings", "USER", "USER")
        sCnvUserID = Replace(oGao.EncodeStr(Trim(str(glUserID)), sUserPass, 0), vbCrLf, "")
Call log("2-----> sUserPass=" & sUserPass & " sCnvUserID=" & sCnvUserID)



        'MAC�A�h���X���Í���
        sMacPass = GetSetting("Nyushi", "Settings", "MAC", "MAC")
        sCnvMacAddr = Replace(oGao.EncodeStr(sMacAddr, sMacPass, 0), vbCrLf, "")
Call log("3-----> sMacPass=" & sMacPass & " sCnvMacAddr=" & sCnvMacAddr)

        '�Í����f�[�^���L�[�Ƀ��j���[�O���[�v���擾
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

            '���j���[�O���[�v�f�[�^�𕜍�
            sMenuGPass = GetSetting("Nyushi", "Settings", "MENUG", "MENUG")
            lMenuID = CLng(Replace(oGao.DecodeStr(sMenuIDStr, sMenuGPass, 0), vbCrLf, ""))

Call log("6-----> sMenuGPass=" & sMenuGPass & " lMenuID=" & lMenuID)



            '�Í��������Z�N�V���������擾
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
                '�Z�N�V�����𕜍�

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
        MsgBox "�{�V�X�e�����g�p���錠��������܂���B", vbOKOnly, "��������"
        End
    End If

    '***************************************************************************
    '* TreeView Menu���� Type Member�z��ɐݒ肷��(������)�֐�               *
    '* TreeView Menu ������ݒ�                                                *
    '***************************************************************************
    Call ls_SetMenues(uMenues_)


    '�������t�@�C���𕜍�
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

    '�������t�@�C���𕜍�
    If Right(App.Path, 1) = "\" Then
        sProfileName = App.Path & prvsProfileName & ".ini"
    Else
        sProfileName = App.Path & "\" & prvsProfileName & ".ini"
    End If

    '**************************************************************************
    '* Passcheck.ini �t�@�C���ύX                                             *
    '*------------------------------------------------------------------------*
    '* 2021.12.22 add jhi                                                     *
    '**************************************************************************

''''�����t���R���p�C�������̐ݒ� 2022.02.01 add jhi
#If zengo_kubun = 1 Then
    sProfileName = App.Path & "\" & prvsProfileName & "_zenki.ini"     ''''Passcheck_zenki.ini
#Else
    sProfileName = App.Path & "\" & prvsProfileName & "_goki.ini"      ''''Passcheck_goki.ini
#End If


    '�������t�@�C����ǎ�
    '***************************************************************************
    '* Passcheck.Ini �t�@�C����ǎ�,[MENU2]Section��key(a01=1)��Ǎ���         *
    '* ����menu��\�����邩? �ݒ肷��                                          *
    '***************************************************************************
    For lCnt = LBound(uMenues_) To UBound(uMenues_)

        sRtn = Space(4)
        lRtn = GetPrivateProfileString(sMenuSection, uMenues_(lCnt).sIniKey, "0", sRtn, 40, sProfileName)

        'key(a01=1)���ݒ肵�Ă���΂���menu�͌�����悤�ɂ���
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
    '* TreeView�̃��j���[���ڂ�S�ĕ\������悤�ɕ\������                      *
    '* 2021.12.09 add jhi                                                      *
    '***************************************************************************
    Select Case f_int_CurrentPhase
    Case 0
        g_int_ExamType = 0

''''    uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeExamKubun")).bVisible = True     '�O�������A���ɏo���Ȃ��̂�

        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeApplyPhase")).bVisible = True    '�菑��t�t�F�[�X
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeFirstPhase")).bVisible = True    '1������
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeSecondPhase")).bVisible = True   '2������
        uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeEnterRefuse")).bVisible = True   '���w�葱������
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


    uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeMasters")).bVisible = True  '�}�X�^���C���e�i���X
    uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodePrint")).bVisible = True    '���
    uMenues_(lf_GetMenuIndex(uMenues_, 0, "nodeTransfer")).bVisible = True '�f�[�^�o��

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
    Me.Toolbar1.Buttons("Search").ToolTipText = "����" ''''LoadResString(1054)

End Sub

'*******************************************************************************
'* TreeView Menu����I�������ۂ̏���                                           *
'*******************************************************************************
Private Sub tvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)


    On Error GoTo ErrorHandler

    Select Case Node.Key

    '---------------------------------------------------------------------------
    ' �菑��t�t�F�[�Y(0)
    '---------------------------------------------------------------------------
    Case "a01"     'Web�o��f�[�^�捞
''''    Call Phase_FlagSet(0)
        mnuOCR_Click

    Case "a02"      '�󌱐��f�[�^�̕ҏW
''''    Call Phase_FlagSet(0)
        mnuMaintainExamineeData_Click

    Case "a03"      '�f�[�^�m��
''''    Call Phase_FlagSet(1)
        f_int_CurrentPhase = 0
        mnuFixData1_Click

''''Case "a03"      '�]��
''''    Call Phase_FlagSet(0)
''''    mnuHyotei_Click


        
    '---------------------------------------------------------------------------
    ' 1������(1)
    '---------------------------------------------------------------------------
    Case "a11"     '������
''''    Call Phase_FlagSet(1)
        mnuRoomAllocation_Click

    Case "a12"     '���Ȏғ���
''''    Call Phase_FlagSet(1)
        mnuInputAbsenteeRecord_Click

    Case "a13"     '�f�_����
''''    Call Phase_FlagSet(1)
        mnuInputRawScore_Click

    Case "a14"      '���i�ғ���
''''    Call Phase_FlagSet(1)
        mnuInputPassedPersonData_Click

    Case "a15"      '2���������U��
''''    Call Phase_FlagSet(1)
        mnuPreparationDay_Click

    Case "a16"      '2���������ύX
''''    Call Phase_FlagSet(1)
        mnuManualAllocation_Click

    Case "a17"      '�f�[�^�m��
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
    ' 2������(2)
    '---------------------------------------------------------------------------
    Case "a21"     '���Ȏғ���
''''    Call Phase_FlagSet(2)
        mnuInputAbsenteeRecord2_Click

    Case "a22"     '�ʐڈψ��o�^
''''    Call Phase_FlagSet(2)
        mnuTeacherRoomMapInterview_Click

    Case "a23"     '�ʐڃO���[�v�U��
''''    Call Phase_FlagSet(2)
        mnuPreparationRoom_Click

    Case "a24"     '�ʐڃO���[�v�ύX
''''    Call Phase_FlagSet(2)
'       mnuSpecialInterview_Click
        mnuManualAllocationGrp_Click

    Case "a25"     '���_���̓_�ψ��o�^
''''    Call Phase_FlagSet(2)
        mnuTeacherRoomMapReport_Click

    Case "a26"     '���_���U��
''''    Call Phase_FlagSet(2)
        mnuPreparationReport_Click

    '---------------------------------------------------------------------------
    Case "a27"     '�f�_����(���_��)_import
''''    Call Phase_FlagSet(2)
        mnuImport_Syoronbun_Click

    Case "a28"     '�f�_����(���_��)
''''    Call Phase_FlagSet(2)
        mnuInputRawScoreI_Click

    Case "a29"     '�f�_����(�ʐ�)_import
''''    Call Phase_FlagSet(2)
        mnuImport_Mensetu_Click

    Case "a30"     '�f�_����(�ʐ�)
''''    Call Phase_FlagSet(2)
        mnuInputRawScore2_Click
    '---------------------------------------------------------------------------

    Case "a31"     '���i�ғ���
''''    Call Phase_FlagSet(2)
        mnuInputPassedPersonData2_Click

    Case "a32"      '�⌇�ғ���
''''    Call Phase_FlagSet(2)
        mnuWaitList2_Click

    '----------------------------------------------------------------------------
    ' 2021.12.02 add jhi S
    '----------------------------------------------------------------------------
    Case "a33"      '�⌇�ҏ���(sub-system��蓝��)
''''    Call Phase_FlagSet(2)
        mnuHoketusyaJuni_Click
    '----------------------------------------------------------------------------
    ' 2021.12.02 add jhi E
    '----------------------------------------------------------------------------

   Case "a34"      '�f�[�^�m��
        Call Phase_FlagSet(3)
        f_int_CurrentPhase = 2
        mnuFixData3_Click


    '----------------------------------------------------------------------------
    ' ���w�葱������
    '----------------------------------------------------------------------------
    Case "a41"     '�⌇�ҍ��i�ҌJ�グ����
''''    Call Phase_FlagSet(3)
        mnuUpliftment_Click

    Case "a42"     '����
''''    Call Phase_FlagSet(3)
        mnuRefuseOffer_Click

   Case "a43"      '�f�[�^�m��
''''    Call Phase_FlagSet(0)
        f_int_CurrentPhase = 3
        mnuFixData4_Click

    '----------------------------------------------------------------------------
    ' �}�X�^�[���C���e�i���X
    '----------------------------------------------------------------------------
    Case "a51"     '���E�ʐڃO���[�v
        mnuRoomProfile_Click

    Case "a52"     '�̓_�҃v���t�@�C��
        mnuInterviewerProfile_Click

    Case "a53"     '�����v���t�B�[��
        mnuInterviewGroupProfile_Click

    Case "a54"     '�����N�x�w��
        mnuSystemData_Click

    '----------------------------------------------------------------------------
    ' ��� Menu
    '----------------------------------------------------------------------------
    Case "a61"     '����w��
        mnuPrintCommand_Click

    Case "a62"     ' Excel���[
        mnuExcelReport_Click

    Case "a63"     ' �x�����z�}���
        mnuPrintDosu_Click

    '----------------------------------------------------------------------------
    ' �f�[�^�o��
    '----------------------------------------------------------------------------
    Case "a71"      '�b�r�u�t�@�C���o��
        mnuOutputCSV_Click


    '----------------------------------------------------------------------------
    ' �ȉ��A���g�p
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
''''    Case "a62"     ' ���шꗗ����w��
''''        mnuSeisekiIchiran_Click
''''
''''    Case "a71"      '�f�[�^�]��
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
'* Menu Bar�Ł@�\�������̂�ݒ肷��                                          *
'* Tree Menu�ɍ��킹��.[tool]-[Menu Editor]�����ݒ�ł���                    *
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
        '* �o���t�t�F�[�Y                                                    *
        '***********************************************************************
        Case "a01"     'Web�o��f�[�^��荞
            mnuOCR.Visible = puMenues_(lLoopCnt).bVisible

        Case "a02"     '�󌱐��f�[�^�ҏW
            mnuMaintainExamineeData.Visible = puMenues_(lLoopCnt).bVisible

''''    Case "a03"    ' hyotei
''''        mnuHyotei.Visible = puMenues_(lLoopCnt).bVisible

        Case "a03"     '�f�[�^�m��
            mnuFixData1.Visible = puMenues_(lLoopCnt).bVisible

        '***********************************************************************
        '* �ꎟ����                                                            *
        '***********************************************************************
        Case "a11"     '������
            mnuRoomAllocation.Visible = puMenues_(lLoopCnt).bVisible

        Case "a12"     '���Ȏғ���
            mnuInputAbsenteeRecord.Visible = puMenues_(lLoopCnt).bVisible

        Case "a13"     '�f�_����
            mnuInputRawScore.Visible = puMenues_(lLoopCnt).bVisible

''''    Case "a14"     ' input choosei score - grace
''''        mnuInputChooseiScore.Visible = puMenues_(lLoopCnt).bVisible
''''    Case "a15"     ' input choosei score - particular student
''''        mnuInputChooseiScore2.Visible = puMenues_(lLoopCnt).bVisible

        Case "a14"     '���i�ғ���
            mnuInputPassedPersonData.Visible = puMenues_(lLoopCnt).bVisible

        Case "a15"      '�������U��
            mnuPreparationDay.Visible = puMenues_(lLoopCnt).bVisible

        Case "a16"     '�������ύX
            mnuManualAllocation.Visible = puMenues_(lLoopCnt).bVisible

'        Case "a19"     ' Manual Allocation
'            mnuPreparationRoom.Visible = puMenues_(lLoopCnt).bVisible

        Case "a17"     '�f�[�^�m��
            mnuFixData2.Visible = puMenues_(lLoopCnt).bVisible


        '***********************************************************************
        '*  2������ Menu                                                       *
        '***********************************************************************
        Case "a21"     '���Ȏғ���
            mnuInputAbsenteeRecord2.Visible = puMenues_(lLoopCnt).bVisible

        Case "a22"     '�ʐڈψ��o�^
            mnuTeacherRoomMapInterview.Visible = puMenues_(lLoopCnt).bVisible

        Case "a23"     '�ʐڃO���[�v�U��
            mnuPreparationRoom.Visible = puMenues_(lLoopCnt).bVisible

        Case "a24"     '�ʐڃO���[�v�ύX
            mnuManualAllocationGrp.Visible = puMenues_(lLoopCnt).bVisible

        Case "a25"     '���_���̓_�ψ��o
            mnuTeacherRoomMapReport.Visible = puMenues_(lLoopCnt).bVisible

        Case "a26"     '���_���U��
            mnuPreparationReport.Visible = puMenues_(lLoopCnt).bVisible

        Case "a27"     '�f�_����(���_��)_import
            mnuImport_Syoronbun.Visible = puMenues_(lLoopCnt).bVisible

        Case "a28"     '�f�_����(���_��)
            mnuInputRawScoreI.Visible = puMenues_(lLoopCnt).bVisible

        Case "a29"     '�f�_����(�ʐ�)_import"
            mnuImport_Mensetu.Visible = puMenues_(lLoopCnt).bVisible

        Case "a30"     '�f�_����(�ʐ�)
            mnuInputRawScore2.Visible = puMenues_(lLoopCnt).bVisible 'Menu�ŕ\������Ȃ��悤�ɂ���2021.12.03 del jhi

        Case "a31"      '���i�ғ�
            mnuInputPassedPersonData2.Visible = puMenues_(lLoopCnt).bVisible

        Case "a32"      '�⌇�ғ���
            mnuWaitList2.Visible = puMenues_(lLoopCnt).bVisible

        Case "a33"      '�⌇�ҏ���
            mnuHoketusyaJuni.Visible = puMenues_(lLoopCnt).bVisible

        Case "a34"      '�f�[�^�m��
            mnuFixData3.Visible = puMenues_(lLoopCnt).bVisible
 
''''    Case "a36"     ' adjust score at Mensetsu
''''        mnuAdjustScoreM.Visible = puMenues_(lLoopCnt).bVisible

        '***********************************************************************
        '* ���w�葱������ Menu                                                 *
        '***********************************************************************
        Case "a41"     '�⌇�ҍ��i�ҌJ�グ����
            mnuUpliftment.Visible = puMenues_(lLoopCnt).bVisible

        Case "a42"     '����
            mnuRefuseOffer.Visible = puMenues_(lLoopCnt).bVisible

        Case "a43"     '�f�[�^�m��
            mnuFixData4.Visible = puMenues_(lLoopCnt).bVisible
        

    '***************************************************************************
    '* �}�X�^�[�����e�i���X Menu                                               *
    '***************************************************************************

        Case "a51"     '���E�ʐڃO���[�v
            mnuRoomProfile.Visible = puMenues_(lLoopCnt).bVisible

        Case "a52"     '�̓_�҃v���t�@�C��
            mnuInterviewerProfile.Visible = puMenues_(lLoopCnt).bVisible

        Case "a53"     '�����v���t�B�[��
            mnuInterviewGroupProfile.Visible = puMenues_(lLoopCnt).bVisible

        Case "a54"     '�����N�x�w��
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
    '* ��� Menu                                                               *
    '***************************************************************************
        Case "a61"     ' ����w��
            mnuPrintCommand.Visible = puMenues_(lLoopCnt).bVisible

        Case "a62"     'Excel���[
            mnuExcelReport.Visible = puMenues_(lLoopCnt).bVisible

        Case "a63"     '�x�����z�}���
            mnuPrintDosu.Visible = puMenues_(lLoopCnt).bVisible

    '***************************************************************************
    '* �f�[�^�o�� Menu                                                         *
    '***************************************************************************

        Case "a71"      '�󌱐��{�f�_���
            mnuOutputCSV.Visible = puMenues_(lLoopCnt).bVisible

        End Select

    Next

    Exit Sub

ErrorHandler:
    MsgBox Err.Description

End Sub


