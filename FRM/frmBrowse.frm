VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBrowse 
   BackColor       =   &H8000000A&
   Caption         =   "frmBrowse : Web�o��f�[�^�捞��"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12915
   ForeColor       =   &H80000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmBrowse.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   12915
   WindowState     =   2  '�ő剻
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import���s"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4785
      TabIndex        =   5
      Top             =   3210
      Width           =   1920
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7815
      TabIndex        =   4
      Top             =   3225
      Width           =   1920
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3400
      TabIndex        =   3
      Top             =   1740
      Width           =   5000
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "�t�@�C���I��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8385
      TabIndex        =   2
      Top             =   1725
      Width           =   1350
   End
   Begin VB.TextBox txtNendo 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3400
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1230
      Width           =   1140
   End
   Begin VB.CheckBox chkInput 
      BackColor       =   &H00F3F3F3&
      Caption         =   "�����o�^"
      Height          =   360
      Left            =   2655
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   3210
      Value           =   1  '����
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   11655
      Top             =   645
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "CSV�t�@�C����I��"
      Filter          =   "Csv Files (*.csv)|*.csv|���̑��e�L�X�g�t�@�C��(*)|*.*|"
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '����
      Caption         =   "( �ȈՔŃt�@�C���̂��w��͂��Ȃ��ł������� )"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3570
      TabIndex        =   11
      Top             =   2400
      Width           =   3840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "(���Z�R�[�h�A�X�֔ԍ��̃`�F�b�N���s��Ȃ�)"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   3630
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "��CSV�t�@�C���ɂ́A�w�b�_�[����̃t�@�C�������w�肭�������B"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3420
      TabIndex        =   9
      Top             =   2175
      Width           =   4890
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "Import�N�x"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   630
      TabIndex        =   8
      Top             =   1320
      Width           =   1380
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  '����
      Caption         =   "lblMsg"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   585
      TabIndex        =   7
      Top             =   4815
      Width           =   12765
   End
   Begin VB.Label lbl02 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "Web�o�� Import�t�@�C��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   375
      TabIndex        =   6
      Top             =   1830
      Width           =   2955
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim logFile As String

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    ''''LoadResStrings Me

    Me.Caption = "frmBrowse : Web�o��f�[�^�捞��"

''''Call g_void_SetFontProperties(Me)                 'set the font properties


    txtNendo.Text = g_int_CurrentNendo

    lblMsg.FontSize = 11                              '2021.12.21 add jhi
    lblMsg.Caption = ""

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'*******************************************************************************
'* �yImport���s�z                                                              *
'*******************************************************************************
Private Sub cmdImport_Click()

    On Error GoTo ErrHandler

    Dim logD  As New Scripting.FileSystemObject
    Dim objTextD As Object

    Dim errLogFlag As String
    Dim errLine As String
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim objCsv As New Scripting.FileSystemObject
    Dim objTextCsv As TextStream
    
    Dim log  As New Scripting.FileSystemObject
    Dim objText As Object
    
    Dim csvFile As String
    Dim strLineData As String
    Dim strLineArray() As String
    Dim strNendo As String
    Dim curLine As Long
    Dim f_bln_UpdateDatabase As Boolean
    
    Dim col_Nendo As Integer
    Dim col_JyukenNo As Integer
    Dim col_Name As Integer
    Dim col_NameFuri As Integer
    Dim col_BirthDay As Integer
    Dim col_Sex As Integer
    Dim col_zipCode1 As Integer
    Dim col_Nation As Integer
    Dim col_HighSchoolID As Integer
    Dim col_HighSchoolAddr As Integer
    Dim col_HighSchoolType As Integer
    Dim col_HighSchoolName As Integer
    Dim col_Katei As Integer
    Dim col_Gaka As Integer
    Dim col_Admiss1 As Integer
    Dim col_Admiss2 As Integer
    Dim col_CollageName As Integer
    Dim col_CollageType As Integer
    Dim col_Score1 As Integer
    Dim col_Score2 As Integer
    Dim col_Language As Integer '�I�O
    Dim col_Rika As Integer     '�I��
    Dim col_MenSetu As Integer  '�ʐړ�
    Dim col_HeiGan As Integer   '����
    Dim col_AddID As Integer    '�s���{���R�[�h
    Dim col_AddName As Integer  '�s���{����
    Dim col_Add1Name As Integer '�Z���P
    Dim col_Add2Name As Integer '�Z��2
    Dim col_Add3Name As Integer '�Z��3

    Dim rinf   As Long
    Dim strMsg As String

    '---------------------------------------------------------------------------
    '���s�m�F�⍇��
    '---------------------------------------------------------------------------
    rinf = myMsgBox("Web�o��CSV�f�[�^��Import���܂��B��낵���ł����H", gTit)
    If rinf = vbCancel Then
        Exit Sub
    End If


    '---------------------------------------------------------------------------
    '�{�^��������ݒ�
    '---------------------------------------------------------------------------
    cmdImport.Enabled = False
    cmdClose.Enabled = False
    cmdSelect.Enabled = False

    lblMsg.ForeColor = &HFF0000 ''''Blue
    lblMsg.Caption = "Web�o��CSV�f�[�^��Import���Ă��܂��B�I���܂ł��΂炭���҂����������B"
    

    errLogFlag = "0"
    Set log = CreateObject("Scripting.FileSystemObject")
    logFile = App.Path & "\Log\" & "Csvlog" & Year(Now) & ".log"
    
    If log.FileExists(logFile) Then
        Set objText = log.OpenTextFile(logFile, ForAppending)
    Else
        Set objText = log.CreateTextFile(logFile, False)
    End If

    objText.WriteLine chkInput.Value & "Start---------CSV " & Now
    
    Set logD = CreateObject("Scripting.FileSystemObject")

    Set objTextD = log.CreateTextFile(App.Path & "\Log\" & "CsvlogDetail" & Year(Now) & ".log", ForAppending)

    csvFile = txtFile.Text
    strNendo = txtNendo.Text
    
    objText.WriteLine "csvfile " & csvFile
    objTextD.WriteLine "csvfile " & csvFile
    Set objCsv = CreateObject("Scripting.FileSystemObject")

    If Len(csvFile) > 1 Then

        If objCsv.FileExists(csvFile) Then
        
            Set objTextCsv = objCsv.OpenTextFile(csvFile, ForReading)
        
            '�s1 ����
'            objTextCsv.ReadLine
            
            '
            col_Nendo = 0
            col_JyukenNo = 1
            col_Name = 2
            col_NameFuri = 3
            col_BirthDay = 4
            col_Sex = 5
            col_zipCode1 = 7
            col_Nation = 8
            col_HighSchoolID = 13
            col_HighSchoolAddr = 9
            col_HighSchoolType = 11
            col_HighSchoolName = 14
            col_Katei = 15
            col_Gaka = 17
            col_Admiss1 = 19
            col_Admiss2 = 21
            col_CollageName = 22
            col_CollageType = 23
            col_Score1 = 25
            col_Score2 = 26
            col_Language = 27
            col_Rika = 28
            col_MenSetu = 30
            col_HeiGan = 32
            col_AddID = 35
            col_AddName = 36
            col_Add1Name = 37
            col_Add2Name = 38
            col_Add3Name = 39
            strLineData = objTextCsv.ReadLine
            objTextD.WriteLine strLineData
            
            strLineData = Replace(strLineData, """", "")
            strLineArray = Split(Trim(strLineData), ",")
             
            Dim cols As Integer

            If UBound(strLineArray) > 38 Then
            
                For cols = 0 To UBound(strLineArray)
                    If Trim(strLineArray(cols)) = "��No" Then
                        col_JyukenNo = cols
                    ElseIf Trim(strLineArray(cols)) = "����" Then
                        col_Name = cols
                    ElseIf Trim(strLineArray(cols)) = "�t���K�i" Then
                        col_NameFuri = cols
                    ElseIf Trim(strLineArray(cols)) = "���N����" Then
                        col_BirthDay = cols
                    ElseIf Trim(strLineArray(cols)) = "����" Then
                       col_Sex = cols
                    ElseIf Trim(strLineArray(cols)) = "�X�֔ԍ�" Then
                        col_zipCode1 = cols
                    ElseIf Trim(strLineArray(cols)) = "����" Then
                        col_Nation = cols
                    ElseIf Trim(strLineArray(cols)) = "�o�g�Z" Then
                        col_HighSchoolID = cols
                    ElseIf Trim(strLineArray(cols)) = "���Z���ݒn��" Then
                        col_HighSchoolAddr = cols
                    ElseIf Trim(strLineArray(cols)) = "���" Then
                        col_HighSchoolType = cols
                    ElseIf Trim(strLineArray(cols)) = "�o�g�Z��" Then
                        col_HighSchoolName = cols
                    ElseIf Trim(strLineArray(cols)) = "�ے�" Then
                        col_Katei = cols
                    ElseIf Trim(strLineArray(cols)) = "�w��" Then
                        col_Gaka = cols
                    ElseIf Trim(strLineArray(cols)) = "���Q�P" Then
                        col_Admiss1 = cols
                    ElseIf Trim(strLineArray(cols)) = "���Q�Q" Then
                        col_Admiss2 = cols
                    ElseIf Trim(strLineArray(cols)) = "��w��" Then
                        col_CollageName = cols
                    ElseIf Trim(strLineArray(cols)) = "�敪" Then
                        col_CollageType = cols
                    ElseIf Trim(strLineArray(cols)) = "�]��" Then
                        col_Score1 = cols
                    ElseIf Trim(strLineArray(cols)) = "����" Then
                        col_Score2 = cols
                    ElseIf Trim(strLineArray(cols)) = "�I�O" Then
                        col_Language = cols
                    ElseIf Trim(strLineArray(cols)) = "�I��" Then
                        col_Rika = cols
                    ElseIf Trim(strLineArray(cols)) = "�ʐڊ�]��" Then
                        col_MenSetu = cols
                    ElseIf Trim(strLineArray(cols)) = "����" Then
                        col_HeiGan = cols
                    ElseIf Trim(strLineArray(cols)) = "�l���F�s���{���R�[�h" Then
                        col_AddID = cols
                    ElseIf Trim(strLineArray(cols)) = "�l���F�s���{����" Then
                        col_AddName = cols
                    ElseIf Trim(strLineArray(cols)) = "�l���F�Z���P" Then
                        col_Add1Name = cols
                    ElseIf Trim(strLineArray(cols)) = "�l���F�Z���Q" Then
                        col_Add2Name = cols
                    ElseIf Trim(strLineArray(cols)) = "�l���F�Z���R" Then
                        col_Add3Name = cols
                    End If
                Next
            Else
                    objText.WriteLine "Cols�s���G" & UBound(strLineArray) & " " & Now
                    objText.Close
                    Set objText = Nothing
                    Set log = Nothing
                    
                    objTextD.WriteLine "Cols�s���G" & UBound(strLineArray) & " " & Now
                    objTextD.Close
                    Set objTextD = Nothing
                    Set logD = Nothing
                    
                    strMsg = ""
                    strMsg = strMsg & "CSV�t�@�C���̗񂪏��Ȃ��ł��B" & vbCrLf
                    strMsg = strMsg & "(�ȈՔłł͂Ȃ��Acsv�t�@�C�������w�肵�Ă�������)"

                    MsgBox strMsg, vbInformation

                    ''''2023.01.23 add jhi
                    lblMsg.Caption = ""      ''''�K�C�_���XMSG��clear
                    cmdSelect.Enabled = True
                    cmdImport.Enabled = True
                    cmdClose.Enabled = True

                    Exit Sub
            End If
             
            curLine = 1
            g_obj_Conn.BeginTrans
            f_bln_UpdateDatabase = True

            While Not objTextCsv.AtEndOfLine

                DoEvents

                curLine = curLine + 1
                errLine = curLine
                strLineData = objTextCsv.ReadLine
                objTextD.WriteLine strLineData
                strLineData = Replace(strLineData, """", "")
                If Not IsNull(strLineData) Then
                    If Trim(strLineData) <> "" Then
                        strLineArray = Split(Trim(strLineData), ",")
                        If UBound(strLineArray) >= 39 Then
                      
                            SQL = "EXEC uspSTEInsertExamineeCSV "
                           
                            SQL = SQL & strNendo & ","  '�N�x
                            If Len(strLineArray(col_JyukenNo)) < 1 Then
                                objText.WriteLine curLine & "�s�̎󌱔ԍ����Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "�s�̎󌱔ԍ����Ȃ��ł��B", vbInformation
'                                GoTo CsvErrHandler
                                objTextD.WriteLine curLine & "�s�̎󌱔ԍ����Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                                errLogFlag = "1"
                            End If
                            If Not IsNumeric(strLineArray(col_JyukenNo)) Then
                                objText.WriteLine curLine & "�s�̎󌱔ԍ��Ɍ�肪����܂��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "�s�̎󌱔ԍ��Ɍ�肪����܂��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "�s�̎󌱔ԍ��Ɍ�肪����܂��B", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            SQL = SQL & strLineArray(col_JyukenNo) & ","  '�󌱔ԍ�
                            SQL = SQL & "'" & strLineArray(col_Name) & "'," '������
                            SQL = SQL & "'" & strLineArray(col_NameFuri) & "'," '�J�i��
                            
                             If Len(strLineArray(col_BirthDay)) < 1 Then
                                objText.WriteLine curLine & "�s�̐��N�������Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "�s�̐��N�������Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "�s�̐��N�������Ȃ��ł��B", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            If Not IsNumeric(strLineArray(col_BirthDay)) And Len(strLineArray(col_BirthDay)) <> 8 Then
                                objText.WriteLine curLine & "�s�̐��N���Ɍ�肪����܂��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "�s�̐��N���Ɍ�肪����܂��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "�s�̐��N���Ɍ�肪����܂��B", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            SQL = SQL & "'" & strLineArray(col_BirthDay) & "'," '���N����
                            
                         
                            If strLineArray(col_Sex) <> "1" And strLineArray(col_Sex) <> "2" Then
                                objText.WriteLine curLine & "�s�̐���(1Or2)�Ɍ�肪����܂��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "�s�̐���(1Or2)�Ɍ�肪����܂��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "�s�̐���(1Or2)�Ɍ�肪����܂��B", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                        
                            SQL = SQL & "'" & strLineArray(col_Sex) & "'," '����
                            
                            SQL = SQL & "'" & strLineArray(col_zipCode1) & "'," '�X�֔ԍ�
                            SQL = SQL & "'" & strLineArray(col_HighSchoolID) & "'," '���Z�R�[�h
                            SQL = SQL & "'" & strLineArray(col_HighSchoolName) & "'," '���Z��
                            SQL = SQL & "'" & strLineArray(col_Katei) & "'," '�ے�
                            SQL = SQL & "'" & strLineArray(col_Gaka) & "'," '�w��
                            SQL = SQL & "'" & strLineArray(col_Admiss1) & "'," '���Q�P
                            SQL = SQL & "'" & strLineArray(col_Admiss2) & "'," '���Q2
                            SQL = SQL & "'" & strLineArray(col_CollageName) & "'," '��w��
                   
                            SQL = SQL & "'" & strLineArray(col_CollageType) & "'," '��w�敪
                            SQL = SQL & "'" & strLineArray(col_Score1) & "'," '�]��
                            
                            '��@a 999 b
                            If strLineArray(col_Score2) = " " Or strLineArray(col_Score2) = "�@" Then
                                SQL = SQL & "'-1'," '����
                            ElseIf strLineArray(col_Score2) = "999" Then
                                SQL = SQL & "'-2'," '����
                            Else
                                SQL = SQL & "'" & strLineArray(col_Score2) & "'," '����
                            End If
                            
                            SQL = SQL & "0,"                            '�p�� (�Œ�H)
                            
                            If strLineArray(col_Rika) <> "1" And strLineArray(col_Rika) <> "2" And strLineArray(col_Rika) <> "3" Then
                                objText.WriteLine curLine & "�s�̑I��(1Or2Or3)�Ɍ�肪����܂��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "�s�̑I��(1Or2Or3)�Ɍ�肪����܂��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "�s�̑I��(1Or2Or3)�Ɍ�肪����܂��B", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            SQL = SQL & "'" & strLineArray(col_Rika) & "'," '�I��
                            
                            
                            If strLineArray(col_MenSetu) <> "1" And strLineArray(col_MenSetu) <> "2" And strLineArray(col_MenSetu) <> "3" Then
                                objText.WriteLine curLine & "�s�̖ʐڊ�]��(1Or2Or3)�Ɍ�肪����܂��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "�s�̖ʐڊ�]��(1Or2Or3)�Ɍ�肪����܂��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "�s�̖ʐڊ�]��(1Or2Or3)�Ɍ�肪����܂��B", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            SQL = SQL & "'" & strLineArray(col_MenSetu) & "'," '�ʐړ�
                            SQL = SQL & "'" & strLineArray(col_HeiGan) & "'," '����
                            SQL = SQL & "'" & strLineArray(9) & "'" '���Z���ݒn��
                            
                            SQL = SQL & ",'" & strLineArray(col_HighSchoolType) & "'" '���ZType
                            SQL = SQL & ",'" & strLineArray(col_Nation) & "'" '����
                            SQL = SQL & ",'" & strLineArray(col_AddID) & "'" '�l���F�s���{���R�[�h
                            SQL = SQL & ",'" & strLineArray(col_AddName) & "'" '�l���F�s���{����
                            If Len(Trim(strLineArray(col_Add1Name)) & Trim(strLineArray(col_Add2Name)) & Trim(strLineArray(col_Add2Name))) < 1 Then
                                    objText.WriteLine curLine & "�s�̏Z�����Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                                    objTextD.WriteLine curLine & "�s�̏Z�����Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                                    MsgBox curLine & "�s�̏Z�����Ȃ��ł��B", vbInformation
'                                    GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            SQL = SQL & ",'" & strLineArray(col_Add1Name) & "'" '�l���F�Z���P
                            SQL = SQL & ",'" & strLineArray(col_Add2Name) & "'" '�l���F�Z��2
                            SQL = SQL & ",'" & strLineArray(col_Add3Name) & "'" '�l���F�Z��3
                            
                            SQL = SQL & ",'" & chkInput.Value & "'"  '�����o�^
                             objTextD.WriteLine "sql  " & SQL
'                            g_obj_Conn.Execute SQL
                             Set RS = g_obj_Conn.Execute(SQL)
                             If RS.EOF Then
                                    objText.WriteLine curLine & "�s�̃f�[�^���C���|�[�g���鎞�A�V�X�e���G���[�������܂����B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                                    objTextD.WriteLine curLine & "�s�̃f�[�^���C���|�[�g���鎞�A�V�X�e���G���[�������܂����B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                                    MsgBox curLine & "�s�̃f�[�^���C���|�[�g���鎞�A�V�X�e���G���[�������܂����B", vbInformation
'                                    GoTo CsvErrHandler
                                errLogFlag = "1"
                             Else
                                If RS.Fields(0).Value = 0 Then
                                ElseIf RS.Fields(0).Value = 1 Then '���Z�R�[�h�Ȃ�
                                    objText.WriteLine curLine & "�s�̍��Z�R�[�h�����݂��Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo) & "  ���Z�R�[�h:" & strLineArray(col_HighSchoolID)
                                    objTextD.WriteLine curLine & "�s�̍��Z�R�[�h�����݂��Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo) & "  ���Z�R�[�h:" & strLineArray(col_HighSchoolID)
'                                    MsgBox curLine & "�s�̍��Z�R�[�h�����݂��Ȃ��ł��B", vbInformation
'                                    GoTo CsvErrHandler
                                errLogFlag = "1"
                                ElseIf RS.Fields(0).Value = 2 Then '�X�֔ԍ�
                                    objText.WriteLine curLine & "�s�̗X�֔ԍ������݂��Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo) & "  �X�֔ԍ�:" & strLineArray(col_zipCode1)
                                    objTextD.WriteLine curLine & "�s�̗X�֔ԍ������݂��Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo) & "  �X�֔ԍ�:" & strLineArray(col_zipCode1)
'                                    MsgBox curLine & "�s�̗X�֔ԍ������݂��Ȃ��ł��B", vbInformation
'                                    GoTo CsvErrHandler
                                    errLogFlag = "1"
                                ElseIf RS.Fields(0).Value = 3 Then '�Z���Ȃ�
                                    objText.WriteLine curLine & "�s�̏Z�����Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                                    objTextD.WriteLine curLine & "�s�̏Z�����Ȃ��ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                                    MsgBox curLine & "�s�̏Z�����Ȃ��ł��B", vbInformation
'                                    GoTo CsvErrHandler
                                errLogFlag = "1"
                                End If
                             End If
                         Else
                           objText.WriteLine curLine & "�s�̗񐔂��s��v�ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
                           objTextD.WriteLine curLine & "�s�̗񐔂��s��v�ł��B" & "  �󌱔ԍ�:" & strLineArray(col_JyukenNo)
'                            MsgBox curLine & "�s�̗񐔂��s��v�ł��B", vbInformation
'                            GoTo CsvErrHandler
                                errLogFlag = "1"
                         End If
                     Else
                        objText.WriteLine " No cols " & curLine & " " & strLineData
                    End If
                Else
                    objText.WriteLine "null cols " & curLine & " " & strLineData
                End If
                
            Wend
            '
            'CSV�t�@�C����Close
            objTextCsv.Close
            Set objTextCsv = Nothing
            Set objCsv = Nothing
    
    
        Else
            objText.WriteLine "csvfile not exist "
            objTextD.WriteLine "csvfile not exist "
            MsgBox "CSV�t�@�C�������݂��Ă��܂���B"
            GoTo CsvErrHandler
        End If
        
        
    Else
        objText.WriteLine "no csvfile "
        objTextD.WriteLine "no csvfile "
        MsgBox "CSV�t�@�C�������݂��Ă��܂���B"
        GoTo CsvErrHandler
        
    End If
    
    If errLogFlag = "1" Then
    
        If chkInput.Value = 1 Then
            objText.WriteLine "End-----------"
        End If
        objText.Close
        Set objText = Nothing
        Set log = Nothing
        If chkInput.Value = 1 Then
            If f_bln_UpdateDatabase = True Then
                g_obj_Conn.CommitTrans
                f_bln_UpdateDatabase = False
            End If

            lblMsg.ForeColor = &HFF& ''''red
            lblMsg.Caption = "CSV���C���|�[�g���܂����B(CSV���e�Ɍ�肪����܂������A�����C���|�[�g���܂����B���O�����m�F���Ă��������B)"
            MsgBox "CSV���C���|�[�g���܂����B" & Chr(10) & "(CSV���e�Ɍ�肪����܂������A�����C���|�[�g���܂����B���O�����m�F���Ă��������B)"

'            Shell "notepad.exe " & logFile
''''         Me.Visible = False

            cmdSelect.Enabled = True
            cmdImport.Enabled = True
            cmdClose.Enabled = True

            Exit Sub
        Else
            If f_bln_UpdateDatabase = True Then
                g_obj_Conn.RollbackTrans
                f_bln_UpdateDatabase = False
                 
             End If

             cmdImport.Enabled = True
             lblMsg.ForeColor = &HFF& ''''Red
             lblMsg.Caption = "CSV���C���|�[�g���ł��܂���ł����B���O�����m�F���Ă��������B"
             MsgBox "CSV���C���|�[�g���ł��܂���ł����B" & Chr(13) & "���O�����m�F���Ă��������B"

'             Shell "notepad.exe " & logFile
        End If

        cmdSelect.Enabled = True
        cmdImport.Enabled = True
        cmdClose.Enabled = True

        Exit Sub

    End If
    
    If f_bln_UpdateDatabase = True Then
        g_obj_Conn.CommitTrans
        f_bln_UpdateDatabase = False
    End If
    
    
    objText.WriteLine "End-----------"
    objText.Close
    Set objText = Nothing
    Set log = Nothing

    lblMsg.ForeColor = &HFF0000 ''''Blue
    lblMsg.Caption = "Web�o��CSV�f�[�^�𐳏��Import���܂����B(Import����=" & curLine - 1 & ")"
    MsgBox "Web�o��CSV�f�[�^�𐳏��Import���܂����B"
''''Me.Visible = False

    '---------------------------------------------------------------------------
    '�{�^��������߂�
    '---------------------------------------------------------------------------
    cmdImport.Enabled = True
    cmdClose.Enabled = True
    cmdSelect.Enabled = True

    Exit Sub

CsvErrHandler:

    'objText.WriteLine "CSv Error " & curLine
    objText.Close
    Set objText = Nothing
    Set log = Nothing

    If f_bln_UpdateDatabase = True Then
        g_obj_Conn.RollbackTrans
        f_bln_UpdateDatabase = False
    End If
    
'    MsgBox "CSV�t�@�C���ɕs���f�[�^�����݂��Ă܂��B"
''''Me.Visible = False

    lblMsg.ForeColor = &HFF& ''''red
    lblMsg.Caption = "Web�o��CSV�f�[�^�ɕs���f�[�^�����݂��Ă܂��B���O�ƍ��킹�Ă��m�F���������B"

    '---------------------------------------------------------------------------
    '�{�^��������߂�
    '---------------------------------------------------------------------------
    cmdImport.Enabled = True
    cmdClose.Enabled = True
    cmdSelect.Enabled = True

    Exit Sub

ErrHandler:

    If f_bln_UpdateDatabase = True Then
        g_obj_Conn.RollbackTrans
        f_bln_UpdateDatabase = False
    End If

    objTextD.WriteLine "Error " & errLine & "  msg:" & Err.Description
    objTextD.Close
    Set objTextD = Nothing
    Set logD = Nothing

    lblMsg.ForeColor = &HFF& ''''red
    lblMsg.Caption = "Web�o��CSV�f�[�^�Ɏ捞�ݏ������G���[���������܂����B���O�ƍ��킹�Ă��m�F���������B"

    '---------------------------------------------------------------------------
    '�{�^��������߂�
    '---------------------------------------------------------------------------
    cmdImport.Enabled = True
    cmdClose.Enabled = True
    cmdSelect.Enabled = True

    MsgBox Err.Description, vbInformation, "�G���["
    
End Sub

'*******************************************************************************
'* �t�@�C���I���{�^������                                                      *
'*******************************************************************************
Private Sub cmdSelect_Click()

    On Error GoTo ErrHandler

    Err.Clear
    dlgFile.ShowOpen

    ' check for cancel error
    If Err.Number = 0 Then
        If dlgFile.FileName <> "" Then
         txtFile.Text = dlgFile.FileName 'Left(dlgFile.FileName, InStrRev(dlgFile.FileName, "\"))
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub cmdClose_Click()

    Me.Visible = False
    Unload Me

End Sub

