VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBrowse 
   BackColor       =   &H8000000A&
   Caption         =   "frmBrowse : Web出願データ取込み"
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
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import実行"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "閉じる"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "ファイル選択"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
         Name            =   "ＭＳ ゴシック"
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
      Caption         =   "強制登録"
      Height          =   360
      Left            =   2655
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   3210
      Value           =   1  'ﾁｪｯｸ
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   11655
      Top             =   645
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "CSVファイルを選択"
      Filter          =   "Csv Files (*.csv)|*.csv|その他テキストファイル(*)|*.*|"
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "( 簡易版ファイルのご指定はしないでください )"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3570
      TabIndex        =   11
      Top             =   2400
      Width           =   3840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "(高校コード、郵便番号のチェックを行わない)"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   3630
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "※CSVファイルには、ヘッダーありのファイルをご指定ください。"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3420
      TabIndex        =   9
      Top             =   2175
      Width           =   4890
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Import年度"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      BackStyle       =   0  '透明
      Caption         =   "lblMsg"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Web出願 Importファイル"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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

    Me.Caption = "frmBrowse : Web出願データ取込み"

''''Call g_void_SetFontProperties(Me)                 'set the font properties


    txtNendo.Text = g_int_CurrentNendo

    lblMsg.FontSize = 11                              '2021.12.21 add jhi
    lblMsg.Caption = ""

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'*******************************************************************************
'* 【Import実行】                                                              *
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
    Dim col_Language As Integer '選外
    Dim col_Rika As Integer     '選理
    Dim col_MenSetu As Integer  '面接日
    Dim col_HeiGan As Integer   '併願
    Dim col_AddID As Integer    '都道府県コード
    Dim col_AddName As Integer  '都道府県名
    Dim col_Add1Name As Integer '住所１
    Dim col_Add2Name As Integer '住所2
    Dim col_Add3Name As Integer '住所3

    Dim rinf   As Long
    Dim strMsg As String

    '---------------------------------------------------------------------------
    '実行確認問合せ
    '---------------------------------------------------------------------------
    rinf = myMsgBox("Web出願CSVデータをImportします。よろしいですか？", gTit)
    If rinf = vbCancel Then
        Exit Sub
    End If


    '---------------------------------------------------------------------------
    'ボタン属性を設定
    '---------------------------------------------------------------------------
    cmdImport.Enabled = False
    cmdClose.Enabled = False
    cmdSelect.Enabled = False

    lblMsg.ForeColor = &HFF0000 ''''Blue
    lblMsg.Caption = "Web出願CSVデータをImportしています。終了までしばらくお待ちください。"
    

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
                    objText.WriteLine "Cols不正；" & UBound(strLineArray) & " " & Now
                    objText.Close
                    Set objText = Nothing
                    Set log = Nothing
                    
                    objTextD.WriteLine "Cols不正；" & UBound(strLineArray) & " " & Now
                    objTextD.Close
                    Set objTextD = Nothing
                    Set logD = Nothing
                    
                    strMsg = ""
                    strMsg = strMsg & "CSVファイルの列が少ないです。" & vbCrLf
                    strMsg = strMsg & "(簡易版ではない、csvファイルをご指定してください)"

                    MsgBox strMsg, vbInformation

                    ''''2023.01.23 add jhi
                    lblMsg.Caption = ""      ''''ガイダンスMSGをclear
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

            lblMsg.ForeColor = &HFF& ''''red
            lblMsg.Caption = "CSVをインポートしました。(CSV内容に誤りがありましたが、強制インポートしました。ログをご確認してください。)"
            MsgBox "CSVをインポートしました。" & Chr(10) & "(CSV内容に誤りがありましたが、強制インポートしました。ログをご確認してください。)"

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
             lblMsg.Caption = "CSVをインポートができませんでした。ログをご確認してください。"
             MsgBox "CSVをインポートができませんでした。" & Chr(13) & "ログをご確認してください。"

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
    lblMsg.Caption = "Web出願CSVデータを正常にImportしました。(Import件数=" & curLine - 1 & ")"
    MsgBox "Web出願CSVデータを正常にImportしました。"
''''Me.Visible = False

    '---------------------------------------------------------------------------
    'ボタン属性を戻す
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
    
'    MsgBox "CSVファイルに不正データが存在してます。"
''''Me.Visible = False

    lblMsg.ForeColor = &HFF& ''''red
    lblMsg.Caption = "Web出願CSVデータに不正データが存在してます。ログと合わせてご確認ください。"

    '---------------------------------------------------------------------------
    'ボタン属性を戻す
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
    lblMsg.Caption = "Web出願CSVデータに取込み処理中エラーが発生しました。ログと合わせてご確認ください。"

    '---------------------------------------------------------------------------
    'ボタン属性を戻す
    '---------------------------------------------------------------------------
    cmdImport.Enabled = True
    cmdClose.Enabled = True
    cmdSelect.Enabled = True

    MsgBox Err.Description, vbInformation, "エラー"
    
End Sub

'*******************************************************************************
'* ファイル選択ボタン処理                                                      *
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

