VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExamineeImport 
   Caption         =   "インポート"
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
   StartUpPosition =   2  '画面の中央
   Begin VB.CheckBox chkInput 
      Caption         =   "強制登録"
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
      Caption         =   "選択"
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
      DialogTitle     =   "CSVファイルを選択"
      Filter          =   "Csv Files (*.csv)|*.csv|その他テキストファイル(*)|*.*|"
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "閉じる"
      Height          =   405
      Left            =   3690
      TabIndex        =   1
      Top             =   1020
      Width           =   855
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "インポート"
      Height          =   405
      Left            =   2520
      TabIndex        =   0
      Top             =   1020
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "ファイル"
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
    Dim col_Language As Integer '選外
    Dim col_Rika As Integer     '選理
    Dim col_MenSetu As Integer  '面接日
    Dim col_HeiGan As Integer   '併願
    Dim col_AddID As Integer    '都道府県コード
    Dim col_AddName As Integer  '都道府県名
    Dim col_Add1Name As Integer '住所１
    Dim col_Add2Name As Integer '住所2
    Dim col_Add3Name As Integer '住所3

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
        
            '行1 無視
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
                    If Trim(strLineArray(cols)) = "受験No" Then
                        col_JyukenNo = cols
                    ElseIf Trim(strLineArray(cols)) = "氏名" Then
                        col_Name = cols
                    ElseIf Trim(strLineArray(cols)) = "フリガナ" Then
                        col_NameFuri = cols
                    ElseIf Trim(strLineArray(cols)) = "生年月日" Then
                        col_BirthDay = cols
                    ElseIf Trim(strLineArray(cols)) = "性別" Then
                       col_Sex = cols
                    ElseIf Trim(strLineArray(cols)) = "郵便番号" Then
                        col_zipCode1 = cols
                    ElseIf Trim(strLineArray(cols)) = "国籍" Then
                        col_Nation = cols
                    ElseIf Trim(strLineArray(cols)) = "出身校" Then
                        col_HighSchoolID = cols
                    ElseIf Trim(strLineArray(cols)) = "高校所在地名" Then
                        col_HighSchoolAddr = cols
                    ElseIf Trim(strLineArray(cols)) = "種別" Then
                        col_HighSchoolType = cols
                    ElseIf Trim(strLineArray(cols)) = "出身校名" Then
                        col_HighSchoolName = cols
                    ElseIf Trim(strLineArray(cols)) = "課程" Then
                        col_Katei = cols
                    ElseIf Trim(strLineArray(cols)) = "学科" Then
                        col_Gaka = cols
                    ElseIf Trim(strLineArray(cols)) = "現浪１" Then
                        col_Admiss1 = cols
                    ElseIf Trim(strLineArray(cols)) = "現浪２" Then
                        col_Admiss2 = cols
                    ElseIf Trim(strLineArray(cols)) = "大学名" Then
                        col_CollageName = cols
                    ElseIf Trim(strLineArray(cols)) = "区分" Then
                        col_CollageType = cols
                    ElseIf Trim(strLineArray(cols)) = "評定" Then
                        col_Score1 = cols
                    ElseIf Trim(strLineArray(cols)) = "欠席" Then
                        col_Score2 = cols
                    ElseIf Trim(strLineArray(cols)) = "選外" Then
                        col_Language = cols
                    ElseIf Trim(strLineArray(cols)) = "選理" Then
                        col_Rika = cols
                    ElseIf Trim(strLineArray(cols)) = "面接希望日" Then
                        col_MenSetu = cols
                    ElseIf Trim(strLineArray(cols)) = "併願" Then
                        col_HeiGan = cols
                    ElseIf Trim(strLineArray(cols)) = "個人情報：都道府県コード" Then
                        col_AddID = cols
                    ElseIf Trim(strLineArray(cols)) = "個人情報：都道府県名" Then
                        col_AddName = cols
                    ElseIf Trim(strLineArray(cols)) = "個人情報：住所１" Then
                        col_Add1Name = cols
                    ElseIf Trim(strLineArray(cols)) = "個人情報：住所２" Then
                        col_Add2Name = cols
                    ElseIf Trim(strLineArray(cols)) = "個人情報：住所３" Then
                        col_Add3Name = cols
                    End If
                Next
            Else
                    objText.WriteLine "Cols不正；" & UBound(strLineArray) & Now
                    objText.Close
                    Set objText = Nothing
                    Set log = Nothing
                    
                    objTextD.WriteLine "Cols不正；" & UBound(strLineArray) & Now
                    objTextD.Close
                    Set objTextD = Nothing
                    Set logD = Nothing
                    
                    MsgBox "CSVファイルの列が少ないです。", vbInformation
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
                           
                            SQL = SQL & strNendo & ","  '年度
                            If Len(strLineArray(col_JyukenNo)) < 1 Then
                                objText.WriteLine curLine & "行の受験番号がないです。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "行の受験番号がないです。", vbInformation
'                                GoTo CsvErrHandler
                                objTextD.WriteLine curLine & "行の受験番号がないです。" & "  受験番号:" & strLineArray(col_JyukenNo)
                                errLogFlag = "1"
                            End If
                            If Not IsNumeric(strLineArray(col_JyukenNo)) Then
                                objText.WriteLine curLine & "行の受験番号に誤りがあります。" & "  受験番号:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "行の受験番号に誤りがあります。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "行の受験番号に誤りがあります。", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            SQL = SQL & strLineArray(col_JyukenNo) & ","  '受験番号
                            SQL = SQL & "'" & strLineArray(col_Name) & "'," '漢字名
                            SQL = SQL & "'" & strLineArray(col_NameFuri) & "'," 'カナ名
                            
                             If Len(strLineArray(col_BirthDay)) < 1 Then
                                objText.WriteLine curLine & "行の生年月日がないです。" & "  受験番号:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "行の生年月日がないです。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "行の生年月日がないです。", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            If Not IsNumeric(strLineArray(col_BirthDay)) And Len(strLineArray(col_BirthDay)) <> 8 Then
                                objText.WriteLine curLine & "行の生年月に誤りがあります。" & "  受験番号:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "行の生年月に誤りがあります。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "行の生年月に誤りがあります。", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            SQL = SQL & "'" & strLineArray(col_BirthDay) & "'," '生年月日
                            
                         
                            If strLineArray(col_Sex) <> "1" And strLineArray(col_Sex) <> "2" Then
                                objText.WriteLine curLine & "行の性別(1Or2)に誤りがあります。" & "  受験番号:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "行の性別(1Or2)に誤りがあります。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "行の性別(1Or2)に誤りがあります。", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                        
                            SQL = SQL & "'" & strLineArray(col_Sex) & "'," '性別
                            
                            SQL = SQL & "'" & strLineArray(col_zipCode1) & "'," '郵便番号
                            SQL = SQL & "'" & strLineArray(col_HighSchoolID) & "'," '高校コード
                            SQL = SQL & "'" & strLineArray(col_HighSchoolName) & "'," '高校名
                            SQL = SQL & "'" & strLineArray(col_Katei) & "'," '課程
                            SQL = SQL & "'" & strLineArray(col_Gaka) & "'," '学科
                            SQL = SQL & "'" & strLineArray(col_Admiss1) & "'," '現浪１
                            SQL = SQL & "'" & strLineArray(col_Admiss2) & "'," '現浪2
                            SQL = SQL & "'" & strLineArray(col_CollageName) & "'," '大学名
                   
                            SQL = SQL & "'" & strLineArray(col_CollageType) & "'," '大学区分
                            SQL = SQL & "'" & strLineArray(col_Score1) & "'," '評定
                            
                            '空　a 999 b
                            If strLineArray(col_Score2) = " " Or strLineArray(col_Score2) = "　" Then
                                SQL = SQL & "'-1'," '欠席
                            ElseIf strLineArray(col_Score2) = "999" Then
                                SQL = SQL & "'-2'," '欠席
                            Else
                                SQL = SQL & "'" & strLineArray(col_Score2) & "'," '欠席
                            End If
                            
                            SQL = SQL & "0,"                            '英語 (固定？)
                            
                            If strLineArray(col_Rika) <> "1" And strLineArray(col_Rika) <> "2" And strLineArray(col_Rika) <> "3" Then
                                objText.WriteLine curLine & "行の選理(1Or2Or3)に誤りがあります。" & "  受験番号:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "行の選理(1Or2Or3)に誤りがあります。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "行の選理(1Or2Or3)に誤りがあります。", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            SQL = SQL & "'" & strLineArray(col_Rika) & "'," '選理
                            
                            
                            If strLineArray(col_MenSetu) <> "1" And strLineArray(col_MenSetu) <> "2" And strLineArray(col_MenSetu) <> "3" Then
                                objText.WriteLine curLine & "行の面接希望日(1Or2Or3)に誤りがあります。" & "  受験番号:" & strLineArray(col_JyukenNo)
                                objTextD.WriteLine curLine & "行の面接希望日(1Or2Or3)に誤りがあります。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                                MsgBox curLine & "行の面接希望日(1Or2Or3)に誤りがあります。", vbInformation
'                                GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            SQL = SQL & "'" & strLineArray(col_MenSetu) & "'," '面接日
                            SQL = SQL & "'" & strLineArray(col_HeiGan) & "'," '併願
                            SQL = SQL & "'" & strLineArray(9) & "'" '高校所在地名
                            
                            SQL = SQL & ",'" & strLineArray(col_HighSchoolType) & "'" '高校Type
                            SQL = SQL & ",'" & strLineArray(col_Nation) & "'" '国籍
                            SQL = SQL & ",'" & strLineArray(col_AddID) & "'" '個人情報：都道府県コード
                            SQL = SQL & ",'" & strLineArray(col_AddName) & "'" '個人情報：都道府県名
                            If Len(Trim(strLineArray(col_Add1Name)) & Trim(strLineArray(col_Add2Name)) & Trim(strLineArray(col_Add2Name))) < 1 Then
                                    objText.WriteLine curLine & "行の住所がないです。" & "  受験番号:" & strLineArray(col_JyukenNo)
                                    objTextD.WriteLine curLine & "行の住所がないです。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                                    MsgBox curLine & "行の住所がないです。", vbInformation
'                                    GoTo CsvErrHandler
                                errLogFlag = "1"
                            End If
                            SQL = SQL & ",'" & strLineArray(col_Add1Name) & "'" '個人情報：住所１
                            SQL = SQL & ",'" & strLineArray(col_Add2Name) & "'" '個人情報：住所2
                            SQL = SQL & ",'" & strLineArray(col_Add3Name) & "'" '個人情報：住所3
                            
                            SQL = SQL & ",'" & chkInput.Value & "'"  '強制登録
                             objTextD.WriteLine "sql  " & SQL
'                            g_obj_Conn.Execute SQL
                             Set RS = g_obj_Conn.Execute(SQL)
                             If RS.EOF Then
                                    objText.WriteLine curLine & "行のデータをインポートする時、システムエラー発生しました。" & "  受験番号:" & strLineArray(col_JyukenNo)
                                    objTextD.WriteLine curLine & "行のデータをインポートする時、システムエラー発生しました。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                                    MsgBox curLine & "行のデータをインポートする時、システムエラー発生しました。", vbInformation
'                                    GoTo CsvErrHandler
                                errLogFlag = "1"
                             Else
                                If RS.Fields(0).Value = 0 Then
                                ElseIf RS.Fields(0).Value = 1 Then '高校コードなし
                                    objText.WriteLine curLine & "行の高校コードが存在しないです。" & "  受験番号:" & strLineArray(col_JyukenNo) & "  高校コード:" & strLineArray(col_HighSchoolID)
                                    objTextD.WriteLine curLine & "行の高校コードが存在しないです。" & "  受験番号:" & strLineArray(col_JyukenNo) & "  高校コード:" & strLineArray(col_HighSchoolID)
'                                    MsgBox curLine & "行の高校コードが存在しないです。", vbInformation
'                                    GoTo CsvErrHandler
                                errLogFlag = "1"
                                ElseIf RS.Fields(0).Value = 2 Then '郵便番号
                                    objText.WriteLine curLine & "行の郵便番号が存在しないです。" & "  受験番号:" & strLineArray(col_JyukenNo) & "  郵便番号:" & strLineArray(col_zipCode1)
                                    objTextD.WriteLine curLine & "行の郵便番号が存在しないです。" & "  受験番号:" & strLineArray(col_JyukenNo) & "  郵便番号:" & strLineArray(col_zipCode1)
'                                    MsgBox curLine & "行の郵便番号が存在しないです。", vbInformation
'                                    GoTo CsvErrHandler
                                    errLogFlag = "1"
                                ElseIf RS.Fields(0).Value = 3 Then '住所なし
                                    objText.WriteLine curLine & "行の住所がないです。" & "  受験番号:" & strLineArray(col_JyukenNo)
                                    objTextD.WriteLine curLine & "行の住所がないです。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                                    MsgBox curLine & "行の住所がないです。", vbInformation
'                                    GoTo CsvErrHandler
                                errLogFlag = "1"
                                End If
                             End If
                         Else
                           objText.WriteLine curLine & "行の列数が不一致です。" & "  受験番号:" & strLineArray(col_JyukenNo)
                           objTextD.WriteLine curLine & "行の列数が不一致です。" & "  受験番号:" & strLineArray(col_JyukenNo)
'                            MsgBox curLine & "行の列数が不一致です。", vbInformation
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
            'CSVファイルをClose
            objTextCsv.Close
            Set objTextCsv = Nothing
            Set objCsv = Nothing
    
    
        Else
            objText.WriteLine "csvfile not exist "
            objTextD.WriteLine "csvfile not exist "
            MsgBox "CSVファイルが存在していません。"
            GoTo CsvErrHandler
        End If
        
        
    Else
        objText.WriteLine "no csvfile "
        objTextD.WriteLine "no csvfile "
        MsgBox "CSVファイルが存在していません。"
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
            MsgBox "CSVファイルをインポートしました。" & Chr(10) & "CSVファイルに誤りがあります。ログを確認してください。"
'            Shell "notepad.exe " & logFile
            Me.Visible = False
            Exit Sub
        Else
            If f_bln_UpdateDatabase = True Then
                g_obj_Conn.RollbackTrans
                f_bln_UpdateDatabase = False
                 
             End If
             MsgBox "CSVファイルインポートができませんでした。" & Chr(13) & "ログを確認してください。"
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
    MsgBox "CSVファイルをインポートしました。"
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
    
'    MsgBox "CSVファイルに不正データが存在してます。"
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

