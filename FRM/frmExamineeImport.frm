VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExamineeImport 
   Caption         =   "�C���|�[�g"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CheckBox chkInput 
      Caption         =   "�����o�^"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1020
      Width           =   1365
   End
   Begin VB.TextBox txtNendo 
      Height          =   405
      Left            =   810
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   660
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "�I��"
      Height          =   405
      Left            =   3990
      TabIndex        =   4
      Top             =   300
      Width           =   555
   End
   Begin VB.TextBox txtFile 
      Height          =   345
      Left            =   960
      TabIndex        =   3
      Top             =   300
      Width           =   2985
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   210
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "CSV�t�@�C����I��"
      Filter          =   "Csv Files (*.csv)|*.csv|���̑��e�L�X�g�t�@�C��(*)|*.*|"
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "����"
      Height          =   405
      Left            =   3690
      TabIndex        =   1
      Top             =   1020
      Width           =   855
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "�C���|�[�g"
      Height          =   405
      Left            =   2520
      TabIndex        =   0
      Top             =   1020
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "�t�@�C��"
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   330
      Width           =   735
   End
End
Attribute VB_Name = "frmExamineeImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Me.Visible = False
'    Unload (Me)
End Sub

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
                    objText.WriteLine "Cols�s���G" & UBound(strLineArray) & Now
                    objText.Close
                    Set objText = Nothing
                    Set log = Nothing
                    
                    objTextD.WriteLine "Cols�s���G" & UBound(strLineArray) & Now
                    objTextD.Close
                    Set objTextD = Nothing
                    Set logD = Nothing
                    
                    MsgBox "CSV�t�@�C���̗񂪏��Ȃ��ł��B", vbInformation
                    Exit Sub
            End If
             
            curLine = 1
            g_obj_Conn.BeginTrans
            f_bln_UpdateDatabase = True
            While Not objTextCsv.AtEndOfLine
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
            MsgBox "CSV�t�@�C�����C���|�[�g���܂����B" & Chr(10) & "CSV�t�@�C���Ɍ�肪����܂��B���O���m�F���Ă��������B"
'            Shell "notepad.exe " & logFile
            Me.Visible = False
            Exit Sub
        Else
            If f_bln_UpdateDatabase = True Then
                g_obj_Conn.RollbackTrans
                f_bln_UpdateDatabase = False
                 
             End If
             MsgBox "CSV�t�@�C���C���|�[�g���ł��܂���ł����B" & Chr(13) & "���O���m�F���Ă��������B"
'             Shell "notepad.exe " & logFile
        End If

    
       
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
    MsgBox "CSV�t�@�C�����C���|�[�g���܂����B"
    Me.Visible = False
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
    Me.Visible = False

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
    MsgBox Err.Description, vbInformation, LoadResString(1729)
    
End Sub

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

