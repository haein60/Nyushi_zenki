VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmImportMensetu 
   Caption         =   "frmImportMensetu : �f�_����(�ʐ�)_import"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmImportMensetu.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   14580
   WindowState     =   2  '�ő剻
   Begin VB.CommandButton cmdDataSet 
      Caption         =   "CSV�f�[�^��DB�ɔ��f"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   10425
      TabIndex        =   7
      Top             =   5880
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6705
      Left            =   435
      TabIndex        =   6
      Top             =   2055
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   11827
      _Version        =   393216
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdCsvDataDisp 
      Caption         =   "CSV�f�[�^�\��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   10410
      TabIndex        =   2
      Top             =   3615
      Width           =   2775
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10815
      TabIndex        =   1
      Top             =   1455
      Width           =   450
   End
   Begin VB.TextBox txtCSVPathFile 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1470
      Width           =   7455
   End
   Begin MSComDlg.CommonDialog cdlCSVPath 
      Left            =   11655
      Top             =   525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select CSV File"
      Filter          =   "CSV Files (*.csv)|*.csv|���̑��e�L�X�g�t�@�C��(*)|*.*|"
   End
   Begin VB.Label lblMsg2 
      BackStyle       =   0  '����
      Caption         =   "lblMsg2"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   465
      TabIndex        =   9
      Top             =   9180
      Width           =   11340
   End
   Begin VB.Label lblGuid1 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "�w�b�_�Ȃ���csv�t�@�C�������w�肭�������B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6375
      TabIndex        =   8
      Top             =   1215
      Width           =   4890
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '����
      Caption         =   "CSV�f�[�^�t�@�C����I��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   450
      TabIndex        =   5
      Top             =   1515
      Width           =   2880
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  '����
      Caption         =   "lblMsg"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   465
      TabIndex        =   0
      Top             =   8880
      Width           =   11340
   End
   Begin VB.Label lblCSVPathFile 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '����
      Caption         =   "�f�_����(�ʐ�)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   450
      TabIndex        =   4
      Top             =   1230
      Width           =   1920
   End
End
Attribute VB_Name = "frmImportMensetu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
'Form Name   : �f�_����(�ʐ�)_import(frmImportMensetu)
'Author      : jhi
'Created On  : 2021.12.21
'Description :
'Reference   :
'*******************************************************************************
Private m_SecondExam_Type    As Long      '�ʐڂ����_����flag
Private CurrentRowNo         As Integer   'active cell�̍s���擾

Dim gFN_CSV                  As String    'Import����csv�t�@�C����
Dim giNendo                  As Long      '�����N�x
Dim gupKensu                 As Long      'update��������

Dim gDestFile                As String    '�ʐ�csv�t�@�C�����T�[�o�ɓ����t�@�C����

Private Type MenCsv_Type
    No                  As Integer
    iNendo              As String
    juno                As Integer
    subcd               As Integer  '20:�ʐڇT 21:�ʐڇU
    Meniin1             As String
    MenSco1             As Single
    Meniin2             As String
    MenSco2             As Single
    Meniin3             As String
    MenSco3             As Single
    fAvg                As Single    ''''Single�^ �P���x���������_��(4byte)
End Type

Private mencsv()    As MenCsv_Type   '�ʐ�csv data


Private Type MenData_Type
    No                  As Integer
    iNendo              As Integer
    juno                As Integer
    subcd               As Integer  '20:�ʐڇT 21:�ʐڇU
    Meniin1             As String
    MenSco1             As Single
    Meniin2             As String
    MenSco2             As Single
    Meniin3             As String
    MenSco3             As Single
    fAvg                As Single    ''''Single�^ �P���x���������_��(4byte)
    iExamineeProfileId  As Long
    idbsetflg           As Integer
End Type

Private MenData()    As MenData_Type   '�ʐ�DB data



'*******************************************************************************
'* Form_Load �֐�                                                              *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler
    Dim rinf    As Long


    Me.Caption = "frmImportMensetu : �f�_����(�ʐ�)_import"
    lblMsg.Caption = ""
    lblMsg2.Caption = ""


Call log("1-----> Form_Load")


''''LoadResStrings Me
''''Call g_void_SetFontProperties(Me)    'set the font properties

    'MSFlexGrid1������
    Call MSFlexGrid1_Mensetu

    If Trim(txtCSVPathFile.Text) = "" Then
        cmdCsvDataDisp.Enabled = False
        cmdDataSet.Enabled = False
    End If

    giNendo = g_int_CurrentNendo
    ''''MsgBox (g_int_CurrentNendo) 'global variable in form load


    rinf = DB_Data_Disp_Men


    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "�G���["

End Sub

'*******************************************************************************
'* MSFlexGrid��csv�f�[�^��\������ �֐�                                        *
'*******************************************************************************
Private Sub cmdCsvDataDisp_Click()

    On Error GoTo ErrorHandler

    Dim sTmp       As String

    Dim rinf       As Long
    Dim step_no    As Long
    Dim i          As Integer


    lblMsg.Caption = ""
    lblMsg2.Caption = ""

    'import����csv�t�@�C���w��L���`�F�b�N
    sTmp = txtCSVPathFile.Text

    If (sTmp = "") Then
        MsgBox "Import����csv�t�@�C�����w�肵�Ă��������B"
        Exit Sub
    End If


    MSFlexGrid1.Clear
    MSFlexGrid1.Refresh

    '---------------------------------------------------------------------------
    ' MSFlexGrid1 �����ݒ�
    '---------------------------------------------------------------------------
    Call MSFlexGrid1_Mensetu


    '�f�[�^�Ǎ��\������
    rinf = ReadCsvFile_Mensetu(sTmp)
    If rinf <> 0 Then
        step_no = 3
        GoTo ErrorHandler
    End If


    MSFlexGrid1.Visible = False        '��U��\���Ɂi�Ǎ��������Ȃ�j

    For i = 0 To UBound(mencsv)

        '�Ǎ��񂾃f�[�^���Z���ɑ��
        MSFlexGrid1.Rows = i + 2
        MSFlexGrid1.Row = i + 1
        MSFlexGrid1.RowHeight(i + 1) = 320

        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = Format$(mencsv(i).No, "###0")       'no

        MSFlexGrid1.Col = 1
        MSFlexGrid1.Text = Format$(mencsv(i).iNendo, "###0")   '�N�x

        MSFlexGrid1.Col = 2
        MSFlexGrid1.Text = Format$(mencsv(i).juno, "000#")     '�󌱔ԍ�

        MSFlexGrid1.Col = 3
        MSFlexGrid1.Text = mencsv(i).subcd                     '�Ȗ�Code20:�ʐڇT 21:�ʐڇU
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter

        MSFlexGrid1.Col = 4
        MSFlexGrid1.Text = mencsv(i).Meniin1                   '�ʐڈψ�1 - A,B,C,D,E
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter

        MSFlexGrid1.Col = 5
        MSFlexGrid1.Text = mencsv(i).MenSco1                   '�ʐڈψ�1 - score
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter

        MSFlexGrid1.Col = 6
        MSFlexGrid1.Text = mencsv(i).Meniin2                   '�ʐڈψ�2 - A,B,C,D,E
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter

        MSFlexGrid1.Col = 7
        MSFlexGrid1.Text = mencsv(i).MenSco2                   '�ʐڈψ�2 - score
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter

        MSFlexGrid1.Col = 8
        MSFlexGrid1.Text = mencsv(i).Meniin3                   '�ʐڈψ�3 - A,B,C,D,E
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter

        MSFlexGrid1.Col = 9
        MSFlexGrid1.Text = mencsv(i).MenSco3                   '�ʐڈψ�3 - score
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter

        MSFlexGrid1.Col = 10
        MSFlexGrid1.Text = Format$(mencsv(i).fAvg, "#0.0")     '���v

    Next i


     '�J�����g�Z�����z�[���|�W�V������
     MSFlexGrid1.Row = 1
     MSFlexGrid1.Col = 1
     MSFlexGrid1.TopRow = 1

     MSFlexGrid1.Visible = True         '�ĕ\��
''''MSFlexGrid1.SetFocus


    '--------------------------------------------------------------------------
    '�ʐڂ�work Table�쐬�֐����ďo��
    '--------------------------------------------------------------------------
    Call Set_Menwork_table(gFN_CSV)

    Exit Sub

ErrorHandler:

    If step_no = 1 Then
        MsgBox "import����csv�t�@�C���̎擾�Ɏ��s���܂����B(�t�@�C����=" & sTmp & ")"

    ElseIf step_no = 2 Then
        MsgBox "import����csv�t�@�C����Open�Ŏ��s���܂����B"

    ElseIf step_no = 3 Then
        MsgBox "import����csv�t�@�C���̗񐔂Ō�肪����܂��Bcsv�t�@�C�������m�F���������B(No=" & rinf & ")"

    ElseIf step_no = 4 Then 'no check
        MsgBox "import����csv�t�@�C������No�쐬�Ɏ��s���܂����B"

    ElseIf step_no = 5 Then '�󌱔ԍ�
        MsgBox "�󌱔ԍ��ݒ�Ɍ�肪����܂����B(No=" & i & ")"

    ElseIf step_no = 6 Then
        MsgBox "�f�_�ݒ�Ɍ�肪����܂����B(No=" & i & ")"

    ElseIf step_no = 7 Then
        MsgBox "import����csv�t�@�C������cnt�쐬�Ɏ��s���܂����B(cnt=" & i & ")"

    ElseIf step_no = 8 Then
        MsgBox "import����csv�t�@�C����Close�Ŏ��s���܂���"

    ElseIf step_no = 9 Then
        MsgBox "�J�����g�Z�����z�[���|�W�V�����ɏ����ŃG���[���������܂����B"

    Else
        MsgBox "import����csv�t�@�C������G���[���������܂����B(step_no=" & step_no & ")"
    End If

    Call MSFlexGrid1_Mensetu



End Sub

'*******************************************************************************
' �ʐ�csv file���e�� work table�ɓ���鏈��
'*******************************************************************************
Private Sub Set_Menwork_table(csvFName As String)

    On Error GoTo ErrorHandler

    Dim oRs  As New ADODB.Recordset
    Dim sSQL As String


    gDestFile = "C:\share\score_mensetsu20_" & giNendo & ".csv"


    'tmp_csvscore20 �ʐ�Table clear
    sSQL = ""
    sSQL = sSQL & "delete tmp_csvscore20" & vbCrLf
    sSQL = sSQL & "where iNendo=" & giNendo

Call log("2-----> sSQL" & sSQL)


    Set oRs = g_obj_Conn.Execute(sSQL)

    'release the object variable
    Set oRs = Nothing


    'csv file���e��tmp_csvscore20 table�ɓ����sql������
    sSQL = ""
    sSQL = sSQL & "BULK INSERT tmp_csvscore20" & vbCrLf
''''sSQL = sSQL & "FROM '" & csvFName & "'" & vbCrLf    ''''2022.02.08 del jhi
    sSQL = sSQL & "FROM '" & gDestFile & "'" & vbCrLf   ''''2022.02.08 add jhi
    sSQL = sSQL & "WITH" & vbCrLf
    sSQL = sSQL & "(" & vbCrLf
    sSQL = sSQL & "   FIELDTERMINATOR = ','," & vbCrLf
    sSQL = sSQL & "   ROWTERMINATOR = '\n'" & vbCrLf
    sSQL = sSQL & ");"

Call log("3-----> sSQL" & sSQL)


    Set oRs = g_obj_Conn.Execute(sSQL)
    
    'release the object variable
    Set oRs = Nothing


    'tbSTEcsvscore20 Table clear
    sSQL = ""
    sSQL = sSQL & "delete tbSTEcsvscore20" & vbCrLf
    sSQL = sSQL & "where iNendo=" & giNendo

Call log("4-----> sSQL" & sSQL)


    Set oRs = g_obj_Conn.Execute(sSQL)

Call log("---->4.5=")


    'release the object variable
    Set oRs = Nothing


    sSQL = ""
    sSQL = sSQL & "insert into tbSTEcsvscore20" & vbCrLf
    sSQL = sSQL & "select" & vbCrLf
    sSQL = sSQL & "    a.*" & vbCrLf
    sSQL = sSQL & "   ,b.iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "   ,0" & vbCrLf                      'idbsetflg ������0���Z�b�g����
    sSQL = sSQL & "   ,GETDATE()" & vbCrLf              'dtCreate
    sSQL = sSQL & "   ,GETDATE()" & vbCrLf              'dtUpdate
    sSQL = sSQL & "from" & vbCrLf
    sSQL = sSQL & "    tmp_csvscore20       a" & vbCrLf
    sSQL = sSQL & "   ,tbSTEExamineeProfile b" & vbCrLf
    sSQL = sSQL & "where" & vbCrLf
    sSQL = sSQL & "        a.iJukenno = b.iJukenNumber" & vbCrLf
    sSQL = sSQL & "    and a.iNendo   = b.iNendo" & vbCrLf
    sSQL = sSQL & "order by" & vbCrLf
    sSQL = sSQL & "    a.iNendo" & vbCrLf
    sSQL = sSQL & "   ,a.iJukenno" & vbCrLf

Call log("---->5=" & sSQL)


    Set oRs = g_obj_Conn.Execute(sSQL)

Call log("---->6=")

    
    'release the object variable
    Set oRs = Nothing

    Exit Sub

ErrorHandler:

    MsgBox "�ʐڍ��v�_��tbSTEcsvscore20 work Table�쐬���G���[���������܂����B(Set_Menwork_table)"

End Sub

'*******************************************************************************
'* Csv�t�@�C���� �I��                                                          *
'*******************************************************************************
Private Sub cmdBrowse_Click()

    On Error GoTo ErrorHandler


    lblMsg.Caption = ""
    lblMsg2.Caption = ""

    Err.Clear
    cdlCSVPath.ShowOpen


    ' check for cancel error
    If Err.Number = 0 Then
''''    txtCSVPathFile.Text = Left(cdlCSVPath.FileName, InStrRev(cdlCSVPath.FileName, "\"))
        txtCSVPathFile.Text = cdlCSVPath.FileName
    End If

    'csv file�����Z�b�g
    gFN_CSV = txtCSVPathFile.Text

    If Trim(gFN_CSV) <> "" Then
        cmdCsvDataDisp.Enabled = True
        cmdDataSet.Enabled = True
  
        '�ʐڂ̐���import �t�@�C�����T�[�o����copy����
        Call fCopy(gFN_CSV, "W:\score_mensetsu20_" & giNendo & ".csv")
''''    Call fCopy(gFN_CSV, "c:\share\score_mensetsu20_" & giNendo & ".csv") ''''2023.02.06 for local kankyo test�̏ꍇ

  End If

    Exit Sub


ErrorHandler:
    MsgBox Err.Description, vbInformation, "�G���[" ''''LoadResString(1729)

End Sub

'*******************************************************************************
'* �ʐ� csv�t�@�C����Ǎ��� type array �ɃZ�b�g����                            *
'*******************************************************************************
Private Function ReadCsvFile_Mensetu(fName As String) As Long

    On Error GoTo GO_ERR

    Dim intFileNo   As Integer
    Dim blnOpenFlg  As Boolean
    Dim lngCnt      As Long

    Dim strFname    As String
    Dim strBuf      As String
    Dim vntBuf      As Variant
    
    Dim rCnt        As Long '�scount
    Dim cCnt        As Long '��count



    '�����l�ݒ�
    blnOpenFlg = False
    
    '�t�@�C�����ݒ�
    strFname = fName
    intFileNo = FreeFile()
    
    '�t�@�C���I�[�v��
    Open strFname For Input As #intFileNo

    '�t�@�C���I�[�v��������t���OOn
    blnOpenFlg = True
    
    rCnt = 1

    Do While Not EOF(intFileNo)

        '�f�[�^����s�P�ʂœǍ�
        Line Input #intFileNo, strBuf

        'Split�������āA���o����������𕪉�
        vntBuf = Split(strBuf, ",")

        '�v�f�����擾
        cCnt = UBound(vntBuf) + 1
        
''''    MsgBox CStr(lngCnt) & "�s�ڂ̍��ڐ���" & CStr(lngDataCnt) & "�ł��B"
        If (cCnt <> 10) Then
            GoTo GO_ERR
        End If

        ReDim Preserve mencsv(rCnt - 1)
                
        mencsv(rCnt - 1).No = rCnt
        mencsv(rCnt - 1).iNendo = vntBuf(0)
        mencsv(rCnt - 1).juno = vntBuf(1)
        mencsv(rCnt - 1).subcd = vntBuf(2)
        mencsv(rCnt - 1).Meniin1 = vntBuf(3)
        mencsv(rCnt - 1).MenSco1 = CSng(vntBuf(4))
        mencsv(rCnt - 1).Meniin2 = vntBuf(5)
        mencsv(rCnt - 1).MenSco2 = CSng(vntBuf(6))
        mencsv(rCnt - 1).Meniin3 = vntBuf(7)
        mencsv(rCnt - 1).MenSco3 = CSng(vntBuf(8))
        mencsv(rCnt - 1).fAvg = CSng(vntBuf(9))
        
        rCnt = rCnt + 1
    Loop
    

    '�t�@�C�������
    If blnOpenFlg = True Then
        Close #intFileNo
        blnOpenFlg = False
    End If

    ReadCsvFile_Mensetu = 0
    Exit Function


GO_ERR:
    
    '�t�@�C�������
    If blnOpenFlg = True Then
        Close #intFileNo
        blnOpenFlg = False
    End If

    ReadCsvFile_Mensetu = rCnt

End Function

'*******************************************************************************
'* MSFlexGrid �̏����ݒ�                                                       *
'*******************************************************************************
Private Sub MSFlexGrid1_Mensetu()

    Dim i As Integer


    'MSFlexGrid �̏����ݒ�
    With MSFlexGrid1

        .Rows = 21                  '�s�̑����i�Œ�s�܂ށj
        .cols = 11                  '��̑����i�Œ��܂ށj
        .FixedRows = 1              '�Œ�s�̐� Rows���P�ȏ㏭�Ȃ���
        .FixedCols = 1              '�Œ��̐� Cols���P�ȏ㏭�Ȃ���
        .Row = 0

        '�񕝂̐ݒ�
        .ColWidth(0) = 600          'No
        .ColWidth(1) = 700          '�N�x
        .ColWidth(2) = 1100         '�󌱔ԍ�

        .ColWidth(3) = 900          '�ȖڃR�[�h

        .ColWidth(4) = 900          '�ʐڈψ�1
        .ColWidth(5) = 800          '�_��1

        .ColWidth(6) = 900          '�ʐڈψ�2
        .ColWidth(7) = 800          '�_��2

        .ColWidth(8) = 900          '�ʐڈψ�3
        .ColWidth(9) = 800          '�_��3

        .ColWidth(10) = 850         '���v�_


        .RowHeight(0) = 320         '�s�̍���

        .Col = 0
        .Text = "No"
        .CellAlignment = flexAlignCenterCenter '�Y���Z�����@����^����@�\��

        .Col = 1
        .Text = "�N�x"
        .CellAlignment = flexAlignCenterCenter

        .Col = 2
        .Text = "�󌱔ԍ�"
        .CellAlignment = flexAlignCenterCenter

        .Col = 3
        .Text = "�ȖڃR�[�h"
        .CellAlignment = flexAlignCenterCenter
 
        .Col = 4
        .Text = "�ʐڈψ�1"
        .CellAlignment = flexAlignCenterCenter

        .Col = 5
        .Text = "�_��1"
        .CellAlignment = flexAlignCenterCenter

        .Col = 6
        .Text = "�ʐڈψ�2"
        .CellAlignment = flexAlignCenterCenter

        .Col = 7
        .Text = "�_��2"
        .CellAlignment = flexAlignCenterCenter

        .Col = 8
        .Text = "�ʐڈψ�3"
        .CellAlignment = flexAlignCenterCenter
 
        .Col = 9
        .Text = "�_��3"
        .CellAlignment = flexAlignCenterCenter

        .Col = 10
        .Text = "���v�_"
        .CellAlignment = flexAlignCenterCenter

        .Col = 0
        For i = 1 To .Rows - 1
            .RowHeight(i) = 320     '�s�̍���
            .Row = i
            .Text = i               '�s�ԍ���\��
        Next i

        .Col = 1
        .Row = 1

        '�J�����g�Z���𔽓]�\���i�����\������΃J�����g�Z��������₷���j
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways

   End With



End Sub

'*******************************************************************************
'* �ʐ� �f�_�� DB tbSTEScoreProfile table�ɔ��f����                            *
'*******************************************************************************
Private Sub cmdDataSet_Click()

    On Error GoTo ErrorHandler
    Dim rinf As Long


    lblMsg.Caption = ""
    lblMsg2.Caption = ""

    rinf = myMsgBox("Import���܂����ʐځA���ϓ_��DB�ɔ��f���܂��B��낵���ł����H", gTit)
    If rinf = vbCancel Then
        Exit Sub
    End If

    cmdDataSet.Enabled = False

    Call Set_Hon_table_Mensetu

''''lblMsg.Caption = gupKensu & "���̖ʐړ_���𐳏��DB�ɔ��f���܂����B"


    cmdDataSet.Enabled = True

    Exit Sub


ErrorHandler:

    MsgBox "�ʐځA���ϓ_��DB�ɔ��f�������G���[���������܂����B"

End Sub

'-------------------------------------------------------------------------------
' csv file�̎w�� fRawScore��tbSTEScoreProfile Table��fRawScore�𔽉f����
'-------------------------------------------------------------------------------
Private Sub Set_Hon_table_Mensetu()

    On Error GoTo ErrorHandler
  
    Dim oRs    As New ADODB.Recordset
    Dim sSQL   As String
    Dim ikensu As Long


    g_obj_Conn.BeginTrans


'*******************************************************************************
'* �d�l�R���Update�̂ݍl�����ꍇ�̏���                                        *
'* �ォ�番�������̂ŉ�ʂœ��͂�����insert�������̂Ń��R�[�h���Ȃ����      *
'* insert���K�v�ɂȂ�̂ŏC������                                              *
'*******************************************************************************
#If 0 Then
    sSQL = ""
    sSQL = sSQL & "update ta" & vbCrLf
    sSQL = sSQL & "   set ta.fRawScore = so.fAvg" & vbCrLf
    sSQL = sSQL & "from" & vbCrLf
    sSQL = sSQL & "    tbSTEScoreProfile ta" & vbCrLf
    sSQL = sSQL & "    inner join" & vbCrLf
    sSQL = sSQL & "    tbSTEcsvscore20 so on" & vbCrLf
    sSQL = sSQL & "            ta.iExamineeProfileId = so.iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "        and substring(convert(nvarchar,ta.dtCreate,111),1,4)=" & giNendo & vbCrLf
    sSQL = sSQL & "        and ta.iSubjectProfileId = 20" & vbCrLf
    sSQL = sSQL & "        and ta.iAbsentFlag       = 0" & vbCrLf
    sSQL = sSQL & "        and so.iNendo            = " & giNendo

'-------------------------------------------------------------------------------
'2021.12.17 add jhi
'-------------------------------------------------------------------------------
Update ta
   Set ta.fRawScore = so.fAvg
From
    tbSTEScoreProfile ta
    Inner Join
    tbSTEcsvscore20 so on
            ta.iExamineeProfileId = so.iExamineeProfileId
        and substring(convert(nvarchar,ta.dtCreate,111),1,4)='2020'
        and ta.iSubjectProfileId = 20
        and ta.iAbsentFlag       = 0
        and so.iNendo            = 2020
--where
--    so.iNendo =2020
'-------------------------------------------------------------------------------
#End If

    '***************************************************************************
    '* tbSTEScoreProfile Table�ɖʐڍ��v�_������                             *
    '* 2022.02.01 update jhi                                                   *
    '***************************************************************************

    '***************************************************************************
    '* insert������܂��̂�SQL�����C��                                         *
    '* 2022.01.26 add jhi                                                      *
    '***************************************************************************
    sSQL = ""
    sSQL = sSQL & "MERGE INTO" & vbCrLf
    sSQL = sSQL & "    tbSTEScoreProfile AS sp" & vbCrLf
    sSQL = sSQL & "USING (" & vbCrLf
    sSQL = sSQL & "    select" & vbCrLf
    sSQL = sSQL & "        (ROW_NUMBER() OVER(ORDER BY iJukenno) + sp.iScoreProfileId )as iScoreProfileId" & vbCrLf
    sSQL = sSQL & "       ,iSubcd    as iSubjectProfileId" & vbCrLf
    sSQL = sSQL & "       ,iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "       ,fAvg" & vbCrLf
    sSQL = sSQL & "       ,0.00      as fChoseiScore" & vbCrLf
    sSQL = sSQL & "       ,0         as iAbsentFlag" & vbCrLf
    sSQL = sSQL & "       ,GETDATE() as dtCreate" & vbCrLf
    sSQL = sSQL & "       ,GETDATE() as dtUpdate" & vbCrLf
    sSQL = sSQL & "    from" & vbCrLf
    sSQL = sSQL & "        tbSTEcsvscore20" & vbCrLf
    sSQL = sSQL & "       ,(select" & vbCrLf
    sSQL = sSQL & "             MAX(iScoreProfileId) as iScoreProfileId" & vbCrLf
    sSQL = sSQL & "         from" & vbCrLf
    sSQL = sSQL & "             tbSTEScoreProfile" & vbCrLf
    sSQL = sSQL & "         where" & vbCrLf
    sSQL = sSQL & "             convert(varchar(4),dtcreate,112) >= '" & g_int_CurrentNendo & "'" & vbCrLf
    sSQL = sSQL & "        ) sp" & vbCrLf
    sSQL = sSQL & "    where" & vbCrLf
    sSQL = sSQL & "        iNendo=" & g_int_CurrentNendo & vbCrLf
    sSQL = sSQL & "    ) cs" & vbCrLf
    sSQL = sSQL & "    on sp.iExamineeProfileId = cs.iExamineeProfileId and (sp.iSubjectProfileID=20 or sp.iSubjectProfileID=21)" & vbCrLf

    sSQL = sSQL & "WHEN MATCHED THEN" & vbCrLf
    sSQL = sSQL & "    UPDATE SET" & vbCrLf
    sSQL = sSQL & "       fRawScore=cs.fAvg" & vbCrLf
    sSQL = sSQL & "      ,dtUpdate =GETDATE()" & vbCrLf

    sSQL = sSQL & "WHEN NOT MATCHED THEN" & vbCrLf
    sSQL = sSQL & "    INSERT(iScoreProfileId,iSubjectProfileId,iExamineeProfileId,fRawScore,fChoseiScore,iAbsentFlag,dtCreate,dtUpdate)" & vbCrLf
    sSQL = sSQL & "    VALUES" & vbCrLf
    sSQL = sSQL & "    (" & vbCrLf
    sSQL = sSQL & "        cs.iScoreProfileId" & vbCrLf
    sSQL = sSQL & "       ,cs.iSubjectProfileId" & vbCrLf
    sSQL = sSQL & "       ,cs.iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "       ,cs.fAvg" & vbCrLf
    sSQL = sSQL & "       ,cs.fChoseiScore" & vbCrLf
    sSQL = sSQL & "       ,cs.iAbsentFlag" & vbCrLf
    sSQL = sSQL & "       ,cs.dtCreate" & vbCrLf
    sSQL = sSQL & "       ,cs.dtUpdate" & vbCrLf
    sSQL = sSQL & "    )" & vbCrLf
    sSQL = sSQL & ";"

Call log("----->MERGE INTO sSQL:" & vbCrLf & sSQL)

'*******************************************************************************
#If 0 Then
Merge Into
    tbSTEScoreProfile As sp
USING (
    select
        (ROW_NUMBER() OVER(ORDER BY iJukenno) + sp.iScoreProfileId )as iScoreProfileId
       ,iSubcd    as iSubjectProfileId
       ,iExamineeProfileId
       ,fAvg
       ,0.00      as fChoseiScore
       ,0         as iAbsentFlag
       ,GETDATE() as dtCreate
       ,GETDATE() as dtUpdate
    From
        tbSTEcsvscore20
       ,(select
             MAX(iScoreProfileId) As iScoreProfileId
         From
             tbSTEScoreProfile
         Where
             convert(varchar(4),dtcreate,112) >= '2022'
        ) sp
    Where
        iNendo = 2022
    ) cs
    on sp.iExamineeProfileId = cs.iExamineeProfileId and (sp.iSubjectProfileID=20 or sp.iSubjectProfileID=21)
WHEN MATCHED THEN
    UPDATE SET
       fRawScore = cs.fAvg
      ,dtUpdate =GETDATE()
WHEN NOT MATCHED THEN
    INSERT(iScoreProfileId,iSubjectProfileId,iExamineeProfileId,fRawScore,fChoseiScore,iAbsentFlag,dtCreate,dtUpdate)
    Values
    (
        cs.iScoreProfileId
       ,cs.iSubjectProfileId
       ,cs.iExamineeProfileId
       ,cs.fAvg
       ,cs.fChoseiScore
       ,cs.iAbsentFlag
       ,cs.dtCreate
       ,cs.dtUpdate
    )
;


#End If
'*******************************************************************************

    Set oRs = g_obj_Conn.Execute(sSQL)

    Set oRs = Nothing

    sSQL = ""
    sSQL = sSQL & "select @@ROWCOUNT;"
    oRs.Open sSQL, g_obj_Conn

    gupKensu = oRs.Fields(0)

    oRs.Close
    Set oRs = Nothing

    lblMsg.Caption = gupKensu & "���̖ʐڍ��v�_��DB�ɔ��f���܂����B"


    '***************************************************************************
    '* tbSTEScoreDetail Table�ɖʐ�3�l�̕]���_�_������                       *
    '* 2022.02.01 add jhi                                                      *
    '***************************************************************************


    '---------------------------------------------------------------------------
    '�����O�ɁA���łɂ���΍폜����
    '---------------------------------------------------------------------------
    sSQL = ""
    sSQL = sSQL & "delete tbSTEScoreDetail" & vbCrLf
    sSQL = sSQL & "where" & vbCrLf
    sSQL = sSQL & "    iScoreProfileId in" & vbCrLf
    sSQL = sSQL & "    (" & vbCrLf
    sSQL = sSQL & "        select" & vbCrLf
    sSQL = sSQL & "            iScoreProfileId" & vbCrLf
    sSQL = sSQL & "        from" & vbCrLf
    sSQL = sSQL & "            tbSTEScoreProfile" & vbCrLf
    sSQL = sSQL & "        where" & vbCrLf
    sSQL = sSQL & "                convert(varchar(4),dtCreate,112) ='" & g_int_CurrentNendo & "'" & vbCrLf
    sSQL = sSQL & "            and iSubjectProfileId in(20,21)" & vbCrLf
    sSQL = sSQL & "    )"

Call log("----->delete tbSTEScoreDetail sSQL:" & vbCrLf & sSQL)

#If 0 Then
#End If

    Set oRs = g_obj_Conn.Execute(sSQL)

    Set oRs = Nothing

    '---------------------------------------------------------------------------
    'tbSTEScoreProfile�ɂ���󌱐��̖ʐ�Detail�f�[�^��insert����
    '---------------------------------------------------------------------------
    sSQL = ""
    sSQL = sSQL & "INSERT INTO tbSTEScoreDetail" & vbCrLf
    sSQL = sSQL & "(" & vbCrLf
    sSQL = sSQL & "    iScoreDetailId" & vbCrLf
    sSQL = sSQL & "   ,iScoreProfileId" & vbCrLf
    sSQL = sSQL & "   ,iSubjectQuestionId" & vbCrLf
    sSQL = sSQL & "   ,siSerialNo" & vbCrLf
    sSQL = sSQL & "   ,fDetailScore" & vbCrLf
    sSQL = sSQL & "   ,dtCreate" & vbCrLf
    sSQL = sSQL & "   ,dtUpdate" & vbCrLf
    sSQL = sSQL & ")" & vbCrLf
    sSQL = sSQL & "select" & vbCrLf
    sSQL = sSQL & ""
    sSQL = sSQL & "    (ROW_NUMBER() OVER(ORDER BY a.iNendo,a.ijukenno,a.iExamineeProfileId,a.seq) + sd.iScoreDetailId ) as iScoreDetailId" & vbCrLf
    sSQL = sSQL & "   ,b.iScoreProfileId as iScoreProfileId" & vbCrLf
    sSQL = sSQL & "   ,a.seq             as iSubjectQuestionId" & vbCrLf
    sSQL = sSQL & "   ,a.seq             as siSerialNo" & vbCrLf
    sSQL = sSQL & "   ,a.rawsco          as fDetailScore" & vbCrLf
    sSQL = sSQL & "   ,getdate()         as dtCreate" & vbCrLf
    sSQL = sSQL & "   ,getdate()         as dtUpdate" & vbCrLf
    sSQL = sSQL & "from" & vbCrLf
    sSQL = sSQL & "    vwSTEScore20      a" & vbCrLf
    sSQL = sSQL & "inner join" & vbCrLf
    sSQL = sSQL & "    tbSTEScoreProfile b" & vbCrLf
    sSQL = sSQL & "on" & vbCrLf
    sSQL = sSQL & "        convert(varchar(4),dtCreate,112) =" & g_int_CurrentNendo & vbCrLf
    sSQL = sSQL & "    and a.iExamineeProfileId=b.iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "    and (b.iSubjectProfileId=20 or b.iSubjectProfileId=21)" & vbCrLf
    sSQL = sSQL & "" & vbCrLf
    sSQL = sSQL & "   ,(select" & vbCrLf
    sSQL = sSQL & "             MAX(iScoreDetailId ) As iScoreDetailId" & vbCrLf
    sSQL = sSQL & "         from" & vbCrLf
    sSQL = sSQL & "             tbSTEScoreDetail" & vbCrLf
    sSQL = sSQL & "         where" & vbCrLf
    sSQL = sSQL & "             convert(varchar(4),dtCreate,112) >= " & g_int_CurrentNendo & vbCrLf
    sSQL = sSQL & "    ) sd"

Call log("----->INSERT INTO tbSTEScoreDetail sSQL:" & vbCrLf & sSQL)

'****************************************
#If 0 Then
INSERT INTO tbSTEScoreDetail
(
    iScoreDetailId
   ,iScoreProfileId
   ,iSubjectQuestionId
   ,siSerialNo
   ,fDetailScore
   ,dtCreate
   ,dtUpdate
)
select
    (ROW_NUMBER() OVER(ORDER BY a.iNendo,a.ijukenno,a.iExamineeProfileId,a.seq) + sd.iScoreDetailId ) as iScoreDetailId
   ,b.iScoreProfileId as iScoreProfileId
   ,a.seq             as iSubjectQuestionId
   ,a.seq             as siSerialNo
   ,a.rawsco          as fDetailScore
   ,getdate()         as dtCreate
   ,getdate()         as dtUpdate
From
    vwSTEScore20 a
Inner Join
    tbSTEScoreProfile b
on
        Convert(VarChar(4), dtCreate, 112) = 2022
    and a.iExamineeProfileId=b.iExamineeProfileId
    and (b.iSubjectProfileId=20 or b.iSubjectProfileId=21)

   ,(select
             MAX(iScoreDetailId) As iScoreDetailId
         From
             tbSTEScoreDetail
         Where
             convert(varchar(4),dtCreate,112) >= 2022
    ) sd

#End If
'****************************************

    Set oRs = g_obj_Conn.Execute(sSQL)

    Set oRs = Nothing

    sSQL = ""
    sSQL = sSQL & "select @@ROWCOUNT;"
    oRs.Open sSQL, g_obj_Conn

    gupKensu = oRs.Fields(0)

    oRs.Close
    Set oRs = Nothing

    lblMsg2.Caption = gupKensu & "����Detail�ʐڕ]���_��DB�ɔ��f���܂����B"


    sSQL = ""
    sSQL = sSQL & "update tbSTEcsvscore20" & vbCrLf
    sSQL = sSQL & "   set idbsetflg = 1" & vbCrLf
    sSQL = sSQL & "where" & vbCrLf
    sSQL = sSQL & "    iNendo = " & giNendo & vbCrLf

Call log("----->update tbSTEcsvscore20 sSQL:" & vbCrLf & sSQL)

    Set oRs = g_obj_Conn.Execute(sSQL)

    Set oRs = Nothing

    g_obj_Conn.CommitTrans


    Exit Sub


ErrorHandler:
    g_obj_Conn.RollbackTrans
    MsgBox "Set_Hon_table_Mensetu()�֐������ŃG���[���������܂����B"

End Sub

'*******************************************************************************
'* 2021.12.12 add jhi                                                          *
'*******************************************************************************
Public Sub gsSetSecondType(piSType As Long)

    If piSType = 0 Then
        m_SecondExam_Type = 0 '�ʐ�
    Else
        m_SecondExam_Type = 1 '���_��
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call g_void_CloseChildForm

End Sub

'*******************************************************************************
'* MSFlexGrid��DB�ɖʐ�import�f�[�^�𔽉f���Ă���Ε\������ �֐�               *
'*******************************************************************************
Private Function DB_Data_Disp_Men() As Long

    On Error GoTo ErrorHandler

    Dim oRs         As New ADODB.Recordset
    Dim sSQL        As String

    Dim step_no     As Integer
    Dim icnt        As Integer    '�f�[�^�̃J�E���g
    Dim i           As Integer    'loop�J�E���g

    Dim rinf        As Long



step_no = 1

    DB_Data_Disp_Men = 0


    MSFlexGrid1.Clear
    MSFlexGrid1.Refresh

    '---------------------------------------------------------------------------
    ' MSFlexGrid1 �����ݒ�
    '---------------------------------------------------------------------------
    Call MSFlexGrid1_Mensetu


    '---------------------------------------------------------------------------
    ' csv work Table��Ǎ���
    '---------------------------------------------------------------------------
    sSQL = ""
    sSQL = sSQL & "select" & vbCrLf
    sSQL = sSQL & "    iNendo" & vbCrLf
    sSQL = sSQL & "   ,iJukenno" & vbCrLf
    sSQL = sSQL & "   ,iSubcd" & vbCrLf
    sSQL = sSQL & "   ,meniin1" & vbCrLf
    sSQL = sSQL & "   ,menSco1" & vbCrLf
    sSQL = sSQL & "   ,meniin2" & vbCrLf
    sSQL = sSQL & "   ,menSco2" & vbCrLf
    sSQL = sSQL & "   ,meniin3" & vbCrLf
    sSQL = sSQL & "   ,menSco3" & vbCrLf
    sSQL = sSQL & "   ,fAvg" & vbCrLf
    sSQL = sSQL & "   ,iExamineeProfileId" & vbCrLf
    sSQL = sSQL & "   ,idbsetflg" & vbCrLf
    sSQL = sSQL & "from" & vbCrLf
    sSQL = sSQL & "    tbSTEcsvscore20" & vbCrLf
    sSQL = sSQL & "where" & vbCrLf
    sSQL = sSQL & "    iNendo=" & giNendo & vbCrLf

    Set oRs = g_obj_Conn.Execute(sSQL)
    
    If oRs.EOF Then
        DB_Data_Disp_Men = 0
        oRs.Close
        Set oRs = Nothing
        Exit Function
    Else
        oRs.MoveFirst
    End If
 

    icnt = 0
    Do While Not oRs.EOF

        ReDim Preserve MenData(icnt)

        MenData(icnt).No = icnt + 1
        MenData(icnt).iNendo = oRs.Fields(0)                'iNendo
        MenData(icnt).juno = oRs.Fields(1)                  'iJukenno
        MenData(icnt).subcd = oRs.Fields(2)                 '20 or 21

        MenData(icnt).Meniin1 = oRs.Fields(3)               'Meniin1
        MenData(icnt).MenSco1 = oRs.Fields(4)               'MenSco1

        MenData(icnt).Meniin2 = oRs.Fields(5)               'Meniin2
        MenData(icnt).MenSco2 = oRs.Fields(6)               'MenSco2

        MenData(icnt).Meniin3 = oRs.Fields(7)               'Meniin3
        MenData(icnt).MenSco3 = oRs.Fields(8)               'MenSco3

        MenData(icnt).fAvg = oRs.Fields(9)                  'fAvg
        MenData(icnt).iExamineeProfileId = oRs.Fields(10)   'iExamineeProfileId
        MenData(icnt).idbsetflg = oRs.Fields(11)            'idbsetflg

        If MenData(icnt).idbsetflg = 0 Then
            GoTo DBSET_NASI
        End If

        icnt = icnt + 1
        oRs.MoveNext
    Loop

    oRs.Close
    Set oRs = Nothing


    MSFlexGrid1.Visible = False        '��U��\����(�Ǎ��������Ȃ�)

    For i = 0 To UBound(MenData)

        MSFlexGrid1.Rows = i + 2
        MSFlexGrid1.Row = i + 1
        MSFlexGrid1.RowHeight(i + 1) = 320

        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = Format$(i + 1, "###0")                   'no

        MSFlexGrid1.Col = 1
        MSFlexGrid1.Text = Format$(MenData(i).iNendo, "###0")       '�N�x

        MSFlexGrid1.Col = 2
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.Text = Format$(MenData(i).juno, "000#")         '�󌱔ԍ�

        '----
        MSFlexGrid1.Col = 3
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.Text = MenData(i).subcd                         '�ʐډȖ�

        '----
        MSFlexGrid1.Col = 4
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.Text = MenData(i).Meniin1                       '�ʐڈψ�1

        MSFlexGrid1.Col = 5
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.Text = MenData(i).MenSco1                       '�_��1

        '----
        MSFlexGrid1.Col = 6
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.Text = MenData(i).Meniin2                       '�ʐڈψ�2

        MSFlexGrid1.Col = 7
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.Text = MenData(i).MenSco2                       '�_��2

        '----
        MSFlexGrid1.Col = 8
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.Text = MenData(i).Meniin3                       '�ʐڈψ�3

        MSFlexGrid1.Col = 9
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.Text = MenData(i).MenSco3                       '�_��3

        '----
        MSFlexGrid1.Col = 10
        MSFlexGrid1.Text = Format$(MenData(i).fAvg, "#0.0")         '�f�_

    Next i


    '�J�����g�Z�����z�[���|�W�V������
    MSFlexGrid1.Row = 1
    MSFlexGrid1.Col = 1
    MSFlexGrid1.TopRow = 1
    MSFlexGrid1.Visible = True         '�ĕ\��

    DB_Data_Disp_Men = i
    Exit Function


DBSET_NASI:

    oRs.Close
    Set oRs = Nothing
 
    DB_Data_Disp_Men = 0
    Exit Function


ErrorHandler:
    MsgBox "DB_Data_Disp_Men()�֐��ŃG���[���������܂����B"


End Function

'*******************************************************************************
'* Form_Activate �֐�                                                          *
'*******************************************************************************
Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim i As Integer

    fMainForm.mnuTools.Enabled = False  ' disable tools menu

    For i = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "�G���["

End Sub

