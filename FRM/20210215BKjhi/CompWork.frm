VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCompWork 
   Caption         =   "小論文入力"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "CompWork.frx":0000
   ScaleHeight     =   9360
   ScaleWidth      =   11925
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdSendNyusi 
      Caption         =   "入試へ"
      Height          =   345
      Left            =   9120
      TabIndex        =   19
      Top             =   8520
      Width           =   1245
   End
   Begin VB.TextBox txtInput 
      Alignment       =   1  '右揃え
      Height          =   210
      Left            =   660
      MaxLength       =   2
      TabIndex        =   8
      Top             =   8610
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "登　　録"
      Height          =   345
      Left            =   7590
      TabIndex        =   9
      Top             =   8520
      Width           =   1245
   End
   Begin VB.Frame frmList 
      Height          =   5625
      Left            =   360
      TabIndex        =   17
      Top             =   2850
      Width           =   11055
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFGrid 
         Height          =   5025
         Left            =   60
         TabIndex        =   7
         Top             =   570
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8864
         _Version        =   393216
         Rows            =   152
         Cols            =   18
         FixedRows       =   2
         FixedCols       =   0
         GridLinesUnpopulated=   1
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   2
         _Band(0).Cols   =   18
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.TextBox txtRand 
         Alignment       =   1  '右揃え
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Left            =   8070
         MaxLength       =   2
         TabIndex        =   6
         Top             =   210
         Width           =   1245
      End
      Begin VB.TextBox txtTest 
         Alignment       =   1  '右揃え
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "yyyy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1530
         MaxLength       =   1
         TabIndex        =   5
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "0:第一日 1:第二日"
         Height          =   255
         Left            =   2910
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblTest 
         Caption         =   "試験日"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblRand 
         Caption         =   "乱　数"
         Height          =   255
         Left            =   7530
         TabIndex        =   18
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame frmTitle 
      Height          =   1695
      Left            =   360
      TabIndex        =   10
      Top             =   1110
      Width           =   11055
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "更　　新"
         Height          =   345
         Left            =   8970
         TabIndex        =   4
         Top             =   1230
         Width           =   1250
      End
      Begin VB.TextBox txtSecondDaySecond 
         Alignment       =   1  '右揃え
         Height          =   270
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtSecondDayFirst 
         Alignment       =   1  '右揃え
         Height          =   270
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtFirstDaySecond 
         Alignment       =   1  '右揃え
         Height          =   270
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtFirstDayFirst 
         Alignment       =   1  '右揃え
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
         Height          =   270
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl2Day2 
         Caption         =   "後半"
         Height          =   255
         Left            =   1980
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lbl2Day1 
         Caption         =   "前半"
         Height          =   255
         Left            =   1980
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lbl2Day 
         Caption         =   "2日目調整点"
         Height          =   255
         Left            =   570
         TabIndex        =   14
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lbl1Day2 
         Caption         =   "後半"
         Height          =   255
         Left            =   1980
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl1Day1 
         Caption         =   "前半"
         Height          =   195
         Left            =   1980
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl1Day 
         Caption         =   "1日目調整点"
         Height          =   255
         Left            =   570
         TabIndex        =   11
         Top             =   240
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmCompWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnUpdatedFlag As Boolean   '登録したデータだけが転送できる
Dim intPreGridRow As Long    '前の行
Dim intPreGridCol As Long    '前の列

Private Sub Grid_init()
    
    'Gridを空にする
    Me.MFGrid.Clear
    
    'Gridの行数と列数
    Me.MFGrid.Rows = 152
    Me.MFGrid.Cols = 18
    
    'Gridの固定行数と列数
    Me.MFGrid.FixedRows = 2
    Me.MFGrid.FixedCols = 0
    
    'Grid合併（一行が合併できる）
    MFGrid.MergeCells = 1
    MFGrid.MergeRow(0) = True
'    MFGrid.MergeCol(0) = True
'    MFGrid.MergeCol(1) = True
    
    'Grid列の幅
    Me.MFGrid.ColWidth(0) = 400
    Me.MFGrid.ColWidth(1) = 1000
    Me.MFGrid.ColWidth(2) = 1000
    Me.MFGrid.ColWidth(3) = 500
    Me.MFGrid.ColWidth(4) = 600
    Me.MFGrid.ColWidth(5) = 500
    Me.MFGrid.ColWidth(6) = 500
    Me.MFGrid.ColWidth(7) = 500
    Me.MFGrid.ColWidth(8) = 500
    Me.MFGrid.ColWidth(9) = 500
    Me.MFGrid.ColWidth(10) = 1000
    Me.MFGrid.ColWidth(11) = 500
    Me.MFGrid.ColWidth(12) = 600
    Me.MFGrid.ColWidth(13) = 500
    Me.MFGrid.ColWidth(14) = 500
    Me.MFGrid.ColWidth(15) = 500
    Me.MFGrid.ColWidth(16) = 500
    Me.MFGrid.ColWidth(17) = 500
    
    'Gridのタイトル文字を中央揃える
    Me.MFGrid.ColAlignmentFixed(0) = 4
    Me.MFGrid.ColAlignmentFixed(1) = 4
    Me.MFGrid.ColAlignmentFixed(2) = 4
    Me.MFGrid.ColAlignmentFixed(3) = 4
    Me.MFGrid.ColAlignmentFixed(4) = 4
    Me.MFGrid.ColAlignmentFixed(5) = 4
    Me.MFGrid.ColAlignmentFixed(6) = 4
    Me.MFGrid.ColAlignmentFixed(7) = 4
    Me.MFGrid.ColAlignmentFixed(8) = 4
    Me.MFGrid.ColAlignmentFixed(9) = 4
    Me.MFGrid.ColAlignmentFixed(10) = 4
    Me.MFGrid.ColAlignmentFixed(11) = 4
    Me.MFGrid.ColAlignmentFixed(12) = 4
    Me.MFGrid.ColAlignmentFixed(13) = 4
    Me.MFGrid.ColAlignmentFixed(14) = 4
    Me.MFGrid.ColAlignmentFixed(15) = 4
    Me.MFGrid.ColAlignmentFixed(16) = 4
    Me.MFGrid.ColAlignmentFixed(17) = 4
    
    'Gridのリスト文字を中央揃える
    Me.MFGrid.ColAlignment(0) = 4
    Me.MFGrid.ColAlignment(1) = 4
    Me.MFGrid.ColAlignment(2) = 4
    Me.MFGrid.ColAlignment(3) = 4
    Me.MFGrid.ColAlignment(4) = 4
    Me.MFGrid.ColAlignment(5) = 4
    Me.MFGrid.ColAlignment(6) = 4
    Me.MFGrid.ColAlignment(7) = 4
    Me.MFGrid.ColAlignment(8) = 4
    Me.MFGrid.ColAlignment(9) = 4
    Me.MFGrid.ColAlignment(10) = 4
    Me.MFGrid.ColAlignment(11) = 4
    Me.MFGrid.ColAlignment(12) = 4
    Me.MFGrid.ColAlignment(13) = 4
    Me.MFGrid.ColAlignment(14) = 4
    Me.MFGrid.ColAlignment(15) = 4
    Me.MFGrid.ColAlignment(16) = 4
    Me.MFGrid.ColAlignment(17) = 4
    
    
    'gridのタイトル1の設定
    Me.MFGrid.TextMatrix(0, 0) = "NO"
    Me.MFGrid.TextMatrix(0, 1) = "得点"
    Me.MFGrid.TextMatrix(0, 2) = "前半"
    Me.MFGrid.TextMatrix(0, 3) = "前半"
    Me.MFGrid.TextMatrix(0, 4) = "前半"
    Me.MFGrid.TextMatrix(0, 5) = "前半"
    Me.MFGrid.TextMatrix(0, 6) = "前半"
    Me.MFGrid.TextMatrix(0, 7) = "前半"
    Me.MFGrid.TextMatrix(0, 8) = "前半"
    Me.MFGrid.TextMatrix(0, 9) = "前半"
    Me.MFGrid.TextMatrix(0, 10) = "後半(1/2で転送)"
    Me.MFGrid.TextMatrix(0, 11) = "後半(1/2で転送)"
    Me.MFGrid.TextMatrix(0, 12) = "後半(1/2で転送)"
    Me.MFGrid.TextMatrix(0, 13) = "後半(1/2で転送)"
    Me.MFGrid.TextMatrix(0, 14) = "後半(1/2で転送)"
    Me.MFGrid.TextMatrix(0, 15) = "後半(1/2で転送)"
    Me.MFGrid.TextMatrix(0, 16) = "後半(1/2で転送)"
    Me.MFGrid.TextMatrix(0, 17) = "後半(1/2で転送)"
    
    'gridのタイトル2の設定
'    Me.MFGrid.TextMatrix(1, 0) = "NO"
'    Me.MFGrid.TextMatrix(1, 1) = "得点"
    Me.MFGrid.TextMatrix(1, 2) = "得点"
    Me.MFGrid.TextMatrix(1, 3) = "調"
    Me.MFGrid.TextMatrix(1, 4) = "平"
    Me.MFGrid.TextMatrix(1, 5) = "採1"
    Me.MFGrid.TextMatrix(1, 6) = "採2"
    Me.MFGrid.TextMatrix(1, 7) = "採3"
    Me.MFGrid.TextMatrix(1, 8) = "採4"
    Me.MFGrid.TextMatrix(1, 9) = "採5"
    Me.MFGrid.TextMatrix(1, 10) = "得点"
    Me.MFGrid.TextMatrix(1, 11) = "調"
    Me.MFGrid.TextMatrix(1, 12) = "平"
    Me.MFGrid.TextMatrix(1, 13) = "採1"
    Me.MFGrid.TextMatrix(1, 14) = "採2"
    Me.MFGrid.TextMatrix(1, 15) = "採3"
    Me.MFGrid.TextMatrix(1, 16) = "採4"
    Me.MFGrid.TextMatrix(1, 17) = "採5"
    
    'Grid行の高さを設定
    Dim I As Long
    For I = 1 To 150
        Me.MFGrid.TextMatrix(I + 1, 0) = I
        Me.MFGrid.RowHeight(I + 1) = 300
    Next



End Sub

Private Sub cmdLoad_Click()
'DBにデータを書きこみ
Dim conn                    As ADODB.Connection
Dim SQL                     As String  'sql文字列
Dim RS                      As ADODB.Recordset

Dim intDate                 As Long     '試験日
Dim intRan                  As Long  '乱数
Dim intNumberOfDateRan      As Long
Dim intTotalScore           As Double   '合計得点
Dim intTotalScore1          As Double   '一回目得点
Dim intChoScore1            As Long  '一回目調点
Dim intAveScore1            As Double   '一回目平均
Dim intP1Score1             As Long  '一回目採1
Dim intP2Score1             As Long  '一回目採1
Dim intP3Score1             As Long  '一回目採3
Dim intP4Score1             As Long  '一回目採4
Dim intP5Score1             As Long  '一回目採5
Dim intTotalScore2          As Double   '2回目得点
Dim intChoScore2            As Long  '2回目調点
Dim intAveScore2            As Double   '2回目平均
Dim intP1Score2             As Long  '2回目採1
Dim intP2Score2             As Long  '2回目採2
Dim intP3Score2             As Long  '2回目採3
Dim intP4Score2             As Long  '2回目採4
Dim intP5Score2             As Long  '2回目採5

Dim I                       As Long
Dim intInputRows            As Long  'Grid有効行数
Dim blnFlagTrans            As Boolean  'Translationのフラグ

Dim minNumberOfDateRan      As Long  '同乱数試験日の最小Number
Dim maxNumberOfDateRan      As Long  '同乱数試験日の最大Number
Dim blnUpdate               As Boolean  '更新フラグtrue:更新  false：新規

On Error GoTo ERR_HANDLE
    If Me.MFGrid.TextMatrix(2, 1) = "" Then
        Exit Sub
    End If
    
    '入力したデータをチェックする
    '試験日
    If Len(Trim(Me.txtTest.Text)) <= 0 Then
        MsgBox "試験日を入力してください"
        Me.txtTest.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Me.txtTest.Text)) <> 1 Then
        MsgBox "1桁試験日を入力してください"
        Me.txtTest.SetFocus
        Exit Sub
    End If
    
    
    If Not IsNumeric(Trim(Me.txtTest.Text)) Then
        MsgBox "数字で試験日を入力してください"
        Me.txtTest.SetFocus
        Exit Sub
    End If
    
    
    Dim strTest As String
    strTest = Trim(Me.txtTest.Text)
    strTest = StrConv(strTest, vbNarrow)
    If strTest <> Trim(Me.txtTest.Text) Then
        MsgBox "半角で試験日を入力してください"
        Me.txtTest.SetFocus
        Exit Sub
    End If
    
    If strTest <> "0" And strTest <> "1" Then
        MsgBox "0また１で試験日を入力してください"
        Me.txtTest.SetFocus
        Exit Sub
    End If
    
    '乱数
    If Len(Trim(Me.txtRand.Text)) <= 0 Then
        MsgBox "乱数を入力してください"
        Me.txtRand.SetFocus
        Exit Sub
    End If
        
    If Not IsNumeric(Trim(Me.txtRand.Text)) Then
        MsgBox "数字で乱数を入力してください"
        Me.txtRand.SetFocus
        Exit Sub
    End If
    
    Dim strRan As String
    strRan = Trim(Me.txtRand.Text)
    strRan = StrConv(strRan, vbNarrow)
    
    If strRan <> Trim(Me.txtRand.Text) Then
        MsgBox "半角で乱数を入力してください"
        Me.txtRand.SetFocus
        Exit Sub
    End If
    
    
    
    
    intDate = Val(strTest)
    intRan = Val(strRan)
    
    Dim objFS   As Object
    Dim objText As Object
    Dim DSN     As String
    Dim uid     As String
    Dim pas     As String
    
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objText = objFS.OpenTextFile("odbc.txt")
    DSN = objText.ReadLine
    uid = objText.ReadLine
    pas = objText.ReadLine
    Set objFS = Nothing
                    
    'DBの接続
    Set conn = New ADODB.Connection
    'conn.ConnectionString = "Provider=SQLOLEDB;Server=DESKPRO815;Database=STE0100;uid=sa"
    conn.Open DSN, uid, pas
    conn.BeginTrans
    blnFlagTrans = True
    
    '新規/更新のチェック
    SQL = ""
    SQL = SQL & "SELECT iNumberOfDateRan FROM tbSTECompwork"
    SQL = SQL & " WHERE iDate=" & Val(Me.txtTest.Text)
    SQL = SQL & " AND   iRan=" & Val(Me.txtRand.Text)
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn
    If RS.EOF Then
        blnUpdate = False
    Else
        blnUpdate = True
    End If
    RS.Close
    
    If blnUpdate = False Then
        intNumberOfDateRan = 0
    Else
        '最大Number数を求める
        SQL = ""
        SQL = SQL & "SELECT MAX(iNumberOfDateRan) as maxNumber,MIN(iNumberOfDateRan) as minNumber FROM tbSTECompwork"
        SQL = SQL & " WHERE iDate=" & Val(Me.txtTest.Text)
        SQL = SQL & " AND   iRan=" & Val(Me.txtRand.Text)
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn
        If Not RS.EOF Then
            RS.MoveFirst
            maxNumberOfDateRan = RS.Fields("maxNumber").Value
            minNumberOfDateRan = RS.Fields("minNumber").Value
        End If
        RS.Close
        intNumberOfDateRan = minNumberOfDateRan
    End If
    
'Debug.Print minNumberOfDateRan
'Debug.Print maxNumberOfDateRan

    '入力したデータを取得する
    For I = 2 To 151
        intTotalScore = Val(Me.MFGrid.TextMatrix(I, 1))
        intTotalScore1 = Val(Me.MFGrid.TextMatrix(I, 2))
        intChoScore1 = Val(Me.MFGrid.TextMatrix(I, 3))
        intAveScore1 = Val(Me.MFGrid.TextMatrix(I, 4))
        intP1Score1 = Val(Me.MFGrid.TextMatrix(I, 5))
        intP2Score1 = Val(Me.MFGrid.TextMatrix(I, 6))
        intP3Score1 = Val(Me.MFGrid.TextMatrix(I, 7))
        intP4Score1 = Val(Me.MFGrid.TextMatrix(I, 8))
        intP5Score1 = Val(Me.MFGrid.TextMatrix(I, 9))
        intTotalScore2 = Val(Me.MFGrid.TextMatrix(I, 10))
        intChoScore2 = Val(Me.MFGrid.TextMatrix(I, 11))
        intAveScore2 = Val(Me.MFGrid.TextMatrix(I, 12))
        intP1Score2 = Val(Me.MFGrid.TextMatrix(I, 13))
        intP2Score2 = Val(Me.MFGrid.TextMatrix(I, 14))
        intP3Score2 = Val(Me.MFGrid.TextMatrix(I, 15))
        intP4Score2 = Val(Me.MFGrid.TextMatrix(I, 16))
        intP5Score2 = Val(Me.MFGrid.TextMatrix(I, 17))
        
        '更新の時、データを消す場合
        If Len(Me.MFGrid.TextMatrix(I, 1)) <= 0 Then
            If intNumberOfDateRan <= maxNumberOfDateRan Then
                SQL = ""
                SQL = SQL & "UPDATE tbSTECompwork "
                SQL = SQL & "SET iTotalScore =null,"
                SQL = SQL & "iTotalScore1 =null,"
                SQL = SQL & "iChoScore1 =null,"
                SQL = SQL & "iAveScore1 =null,"
                SQL = SQL & "iP1Score1 =null,"
                SQL = SQL & "iP2Score1 =null,"
                SQL = SQL & "iP3Score1 =null,"
                SQL = SQL & "iP4Score1 =null,"
                SQL = SQL & "iP5Score1 =null,"
                SQL = SQL & "iTotalScore2 =null,"
                SQL = SQL & "iChoScore2 =null,"
                SQL = SQL & "iAveScore2 =null,"
                SQL = SQL & "iP1Score2 =null,"
                SQL = SQL & "iP2Score2 =null,"
                SQL = SQL & "iP3Score2 =null,"
                SQL = SQL & "iP4Score2 =null,"
                SQL = SQL & "iP5Score2 =null,"
                SQL = SQL & "dtUpdate=GETDATE()"
                SQL = SQL & " WHERE iDate=" & intDate
                SQL = SQL & " AND iRan= " & intRan
                SQL = SQL & " AND iNumberOfDateRan =" & intNumberOfDateRan
            End If
        End If
        
        'データがある場合、DB新規する
        If Len(Me.MFGrid.TextMatrix(I, 1)) > 0 Then
        
            '新規の場合
            If blnUpdate = False Then
                SQL = ""
                SQL = SQL & "INSERT INTO tbSTECompwork(iDate,iRan,iNumberOfDateRan,iTotalScore,"
                If Len(Me.MFGrid.TextMatrix(I, 2)) > 0 Then
                    SQL = SQL & "iTotalScore1,iChoScore1,iAveScore1,"
                    SQL = SQL & "iP1Score1,iP2Score1,iP3Score1,iP4Score1,iP5Score1,"
                End If
                If Len(Me.MFGrid.TextMatrix(I, 10)) > 0 Then
                    SQL = SQL & "iTotalScore2,iChoScore2,iAveScore2,"
                    SQL = SQL & "iP1Score2,iP2Score2,iP3Score2,iP4Score2,iP5Score2,"
                End If
                SQL = SQL & "dtUpdate,dtCreate) "
                SQL = SQL & "VALUES( "
                SQL = SQL & intDate & ","
                SQL = SQL & intRan & ","
                SQL = SQL & intNumberOfDateRan & ","
                SQL = SQL & intTotalScore & ","
                If Len(Me.MFGrid.TextMatrix(I, 2)) > 0 Then
                    SQL = SQL & intTotalScore1 & ","
                    SQL = SQL & intChoScore1 & ","
                    SQL = SQL & intAveScore1 & ","
                    If Len(Me.MFGrid.TextMatrix(I, 5)) > 0 Then
                        SQL = SQL & intP1Score1 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 6)) > 0 Then
                        SQL = SQL & intP2Score1 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 7)) > 0 Then
                        SQL = SQL & intP3Score1 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                                 
                    If Len(Me.MFGrid.TextMatrix(I, 8)) > 0 Then
                        SQL = SQL & intP4Score1 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 9)) > 0 Then
                        SQL = SQL & intP5Score1 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
            
                End If
                
                If Len(Me.MFGrid.TextMatrix(I, 10)) > 0 Then
                    SQL = SQL & intTotalScore2 & ","
                    SQL = SQL & intChoScore2 & ","
                    SQL = SQL & intAveScore2 & ","
                    
                    If Len(Me.MFGrid.TextMatrix(I, 13)) > 0 Then
                        SQL = SQL & intP1Score2 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 14)) > 0 Then
                        SQL = SQL & intP2Score2 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 15)) > 0 Then
                        SQL = SQL & intP3Score2 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                                 
                    If Len(Me.MFGrid.TextMatrix(I, 16)) > 0 Then
                        SQL = SQL & intP4Score2 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 17)) > 0 Then
                        SQL = SQL & intP5Score2 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                
                End If
                SQL = SQL & "GETDATE(),GETDATE()"
                SQL = SQL & ")"
  
            '更新の場合
            Else
                If intNumberOfDateRan <= maxNumberOfDateRan Then
                
                    '更新用SQL文
                    SQL = ""
                    SQL = SQL & "UPDATE tbSTECompwork "
                    SQL = SQL & "SET iTotalScore =" & intTotalScore & ","
                    If Len(Me.MFGrid.TextMatrix(I, 2)) > 0 Then
                    SQL = SQL & "iTotalScore1 =" & intTotalScore1 & ","
                    SQL = SQL & "iChoScore1 =" & intChoScore1 & ","
                    SQL = SQL & "iAveScore1 =" & intAveScore1 & ","
                    
                    If Len(Me.MFGrid.TextMatrix(I, 5)) > 0 Then
                        SQL = SQL & "iP1Score1 =" & intP1Score1 & ","
                    Else
                        SQL = SQL & "iP1Score1 =null,"
                    End If
                    If Len(Me.MFGrid.TextMatrix(I, 6)) > 0 Then
                        SQL = SQL & "iP2Score1 =" & intP2Score1 & ","
                    Else
                        SQL = SQL & "iP2Score1 =null,"
                    End If
                    If Len(Me.MFGrid.TextMatrix(I, 7)) > 0 Then
                        SQL = SQL & "iP3Score1 =" & intP3Score1 & ","
                    Else
                        SQL = SQL & "iP3Score1 =null,"
                    End If
                    If Len(Me.MFGrid.TextMatrix(I, 8)) > 0 Then
                        SQL = SQL & "iP4Score1 =" & intP4Score1 & ","
                    Else
                        SQL = SQL & "iP4Score1 =null,"
                    End If
                    If Len(Me.MFGrid.TextMatrix(I, 9)) > 0 Then
                        SQL = SQL & "iP5Score1 =" & intP5Score1 & ","
                    Else
                        SQL = SQL & "iP5Score1 =null,"
                    End If
                    
                    Else
                        SQL = SQL & "iTotalScore1 =null,"
                        SQL = SQL & "iChoScore1 =null,"
                        SQL = SQL & "iAveScore1 =null,"
                        SQL = SQL & "iP1Score1 =null,"
                        SQL = SQL & "iP2Score1 =null,"
                        SQL = SQL & "iP3Score1 =null,"
                        SQL = SQL & "iP4Score1 =null,"
                        SQL = SQL & "iP5Score1 =null,"
                    End If
                    If Len(Me.MFGrid.TextMatrix(I, 10)) > 0 Then
                    
                    
                        SQL = SQL & "iTotalScore2 =" & intTotalScore2 & ","
                        SQL = SQL & "iChoScore2 =" & intChoScore2 & ","
                        SQL = SQL & "iAveScore2 =" & intAveScore2 & ","
                        
                        If Len(Me.MFGrid.TextMatrix(I, 13)) > 0 Then
                            SQL = SQL & "iP1Score2 =" & intP1Score2 & ","
                        Else
                            SQL = SQL & "iP1Score2 =null,"
                        End If
                        If Len(Me.MFGrid.TextMatrix(I, 14)) > 0 Then
                            SQL = SQL & "iP2Score2 =" & intP2Score2 & ","
                        Else
                            SQL = SQL & "iP2Score2 =null,"
                        End If
                        If Len(Me.MFGrid.TextMatrix(I, 15)) > 0 Then
                            SQL = SQL & "iP3Score2 =" & intP3Score2 & ","
                        Else
                            SQL = SQL & "iP3Score2 =null,"
                        End If
                        If Len(Me.MFGrid.TextMatrix(I, 16)) > 0 Then
                            SQL = SQL & "iP4Score2 =" & intP4Score2 & ","
                        Else
                            SQL = SQL & "iP4Score2 =null,"
                        End If
                        If Len(Me.MFGrid.TextMatrix(I, 17)) > 0 Then
                            SQL = SQL & "iP5Score2 =" & intP5Score2 & ","
                        Else
                            SQL = SQL & "iP5Score2 =null,"
                        End If
                    
                    Else
                        SQL = SQL & "iTotalScore2 =null,"
                        SQL = SQL & "iChoScore2 =null,"
                        SQL = SQL & "iAveScore2 =null,"
                        SQL = SQL & "iP1Score2 =null,"
                        SQL = SQL & "iP2Score2 =null,"
                        SQL = SQL & "iP3Score2 =null,"
                        SQL = SQL & "iP4Score2 =null,"
                        SQL = SQL & "iP5Score2 =null,"
                    End If
                    SQL = SQL & "dtUpdate=GETDATE()"
                    SQL = SQL & " WHERE iDate=" & intDate
                    SQL = SQL & " AND iRan= " & intRan
                    SQL = SQL & " AND iNumberOfDateRan =" & intNumberOfDateRan

                '更新の場合、新規データの追加
                Else
                    SQL = ""
                    SQL = SQL & "INSERT INTO tbSTECompwork(iDate,iRan,iNumberOfDateRan,iTotalScore,"
                If Len(Me.MFGrid.TextMatrix(I, 2)) > 0 Then
                    SQL = SQL & "iTotalScore1,iChoScore1,iAveScore1,"
                    SQL = SQL & "iP1Score1,iP2Score1,iP3Score1,iP4Score1,iP5Score1,"
                End If
                If Len(Me.MFGrid.TextMatrix(I, 10)) > 0 Then
                    SQL = SQL & "iTotalScore2,iChoScore2,iAveScore2,"
                    SQL = SQL & "iP1Score2,iP2Score2,iP3Score2,iP4Score2,iP5Score2,"
                End If
                SQL = SQL & "dtUpdate,dtCreate) "
                SQL = SQL & "VALUES( "
                SQL = SQL & intDate & ","
                SQL = SQL & intRan & ","
                SQL = SQL & intNumberOfDateRan & ","
                SQL = SQL & intTotalScore & ","
                If Len(Me.MFGrid.TextMatrix(I, 2)) > 0 Then
                    SQL = SQL & intTotalScore1 & ","
                    SQL = SQL & intChoScore1 & ","
                    SQL = SQL & intAveScore1 & ","
                    If Len(Me.MFGrid.TextMatrix(I, 5)) > 0 Then
                        SQL = SQL & intP1Score1 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 6)) > 0 Then
                        SQL = SQL & intP2Score1 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 7)) > 0 Then
                        SQL = SQL & intP3Score1 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                                 
                    If Len(Me.MFGrid.TextMatrix(I, 8)) > 0 Then
                        SQL = SQL & intP4Score1 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 9)) > 0 Then
                        SQL = SQL & intP5Score1 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
            
                End If
                
                If Len(Me.MFGrid.TextMatrix(I, 10)) > 0 Then
                    SQL = SQL & intTotalScore2 & ","
                    SQL = SQL & intChoScore2 & ","
                    SQL = SQL & intAveScore2 & ","
                    
                    If Len(Me.MFGrid.TextMatrix(I, 13)) > 0 Then
                        SQL = SQL & intP1Score2 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 14)) > 0 Then
                        SQL = SQL & intP2Score2 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 15)) > 0 Then
                        SQL = SQL & intP3Score2 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                                 
                    If Len(Me.MFGrid.TextMatrix(I, 16)) > 0 Then
                        SQL = SQL & intP4Score2 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                    
                    If Len(Me.MFGrid.TextMatrix(I, 17)) > 0 Then
                        SQL = SQL & intP5Score2 & ","
                    Else
                        SQL = SQL & "null,"
                    End If
                
                End If
                SQL = SQL & "GETDATE(),GETDATE()"
                SQL = SQL & ")"

                End If
            End If
        End If
        If SQL <> "" Then
'Debug.Print sql
            conn.Execute (SQL)
            intNumberOfDateRan = intNumberOfDateRan + 1
        End If
        SQL = ""
    Next
  
    'DBのコミット
    conn.CommitTrans
    blnFlagTrans = False
    blnUpdatedFlag = True
    'DBの切断
    conn.Close
    
Exit Sub
ERR_HANDLE:
    MsgBox Err.Description
    If blnFlagTrans = True Then
        conn.RollbackTrans
        blnFlagTrans = False
    End If
    Set RS = Nothing
    If conn Is Nothing Then
    Else
        conn.Close
    End If
    
        
End Sub

Private Sub cmdSendNyusi_Click()
'DBにデータを書きこみ
Dim conn                    As ADODB.Connection
Dim SQL                     As String  'sql文字列
Dim RS                      As ADODB.Recordset
Dim intDate                 As Long     '試験日
Dim intRan                  As Long  '乱数
Dim intNumberOfDateRan      As Long
Dim intTotalScore           As Double   '合計得点
Dim strRan As String
Dim objFS   As Object
Dim objText As Object
Dim DSN     As String
Dim uid     As String
Dim pas     As String
Dim blnFlagTrans As Boolean
Dim strTestdate As String
Dim iRoomProfileid As Long
Dim iNendo As Long
Dim iSubjectProfileID  As Long

'新規・更新フラグ true:新規；false:更新
Dim blnFlag                 As Boolean
On Error GoTo ERR_HANDLE
    '1.初期処理
    '小論文
    iSubjectProfileID = 30
    
    '転送確認メッセージを出る
    If MsgBox("小論文素点を入試システムへ転送します。よろしいですか？", vbOKCancel + vbInformation) = vbCancel Then
        Exit Sub
    End If
    '入力したデータをチェックする
    '試験日
    If Len(Trim(Me.txtTest.Text)) <= 0 Then
        MsgBox "試験日を入力してください"
        Me.txtTest.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Me.txtTest.Text)) <> 1 Then
        MsgBox "1桁試験日を入力してください"
        Me.txtTest.SetFocus
        Exit Sub
    End If
    
    
    If Not IsNumeric(Trim(Me.txtTest.Text)) Then
        MsgBox "数字で試験日を入力してください"
        Me.txtTest.SetFocus
        Exit Sub
    End If
    
    
    Dim strTest As String
    strTest = Trim(Me.txtTest.Text)
    strTest = StrConv(strTest, vbNarrow)
    If strTest <> Trim(Me.txtTest.Text) Then
        MsgBox "半角で試験日を入力してください"
        Me.txtTest.SetFocus
        Exit Sub
    End If
    
    If strTest <> "0" And strTest <> "1" Then
        MsgBox "0また１で試験日を入力してください"
        Me.txtTest.SetFocus
        Exit Sub
    End If
    
    '乱数
    If Len(Trim(Me.txtRand.Text)) <= 0 Then
        MsgBox "乱数を入力してください"
        Me.txtRand.SetFocus
        Exit Sub
    End If
        
    If Not IsNumeric(Trim(Me.txtRand.Text)) Then
        MsgBox "数字で乱数を入力してください"
        Me.txtRand.SetFocus
        Exit Sub
    End If
    
    
    strRan = Trim(Me.txtRand.Text)
    strRan = StrConv(strRan, vbNarrow)
    
    If strRan <> Trim(Me.txtRand.Text) Then
        MsgBox "半角で乱数を入力してください"
        Me.txtRand.SetFocus
        Exit Sub
    End If
    intDate = Val(strTest)
    intRan = Val(strRan)
    
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objText = objFS.OpenTextFile("odbc.txt")
    DSN = objText.ReadLine
    uid = objText.ReadLine
    pas = objText.ReadLine
    Set objFS = Nothing
                    
    'DBの接続
    Set conn = New ADODB.Connection
    'conn.ConnectionString = "Provider=SQLOLEDB;Server=DESKPRO815;Database=STE0100;uid=sa"
    conn.Open DSN, uid, pas
    conn.CursorLocation = adUseClient
    
    
    '2.データ転送処理
    '①試験日と乱数より、受験番号と採点者を検索する
    SQL = ""
    SQL = SQL & " SELECT  sep.dtSecondExamDay1, sep.dtSecondExamDay2,sp.iNendo"
    SQL = SQL & "  FROM tbSTESecondExamProfile sep,tbSTESystemProfile sp"
    SQL = SQL & "  WHERE sep.iSystemProfileId = sp.iSystemProfileId"
    SQL = SQL & "  AND  sp.iActiveFlag = 1"
    Set RS = conn.Execute(SQL)
    If RS.EOF Then
        Exit Sub
    End If
    RS.MoveFirst
    
    If intDate = 0 Then
        strTestdate = Format(RS.Fields(0).Value & "", "MM/DD/YYYY")
    Else
        strTestdate = Format(RS.Fields(1).Value & "", "MM/DD/YYYY")
    End If
    iNendo = RS.Fields(2).Value
    RS.Close
    Set RS = Nothing
    
    SQL = ""
    SQL = SQL & " SELECT iRoomProfileId FROM tbSTERoomProfile "
    SQL = SQL & " WHERE iRandom=" & intRan
    'update,xzg,2008/02/12,S--------
    'SQL = SQL & " AND iInterviewRoomFlag=0"
    SQL = SQL & " AND iInterviewRoomFlag=1"
    'update,xzg,2008/02/12,E--------
    Set RS = conn.Execute(SQL)
    If RS.EOF Then
        Exit Sub
    End If
    RS.MoveFirst
    iRoomProfileid = RS.Fields(0).Value
    RS.Close
    Set RS = Nothing
    
    Dim iScoreProfileIDMax As Long
    Dim iScoreDetailMax As Long
    Dim rst As ADODB.Recordset
    Dim rst1 As ADODB.Recordset
    SQL = ""
    SQL = SQL & " SELECT ISNULL(MAX(iScoreProfileId),0) FROM tbSTEScoreProfile "
    Set rst1 = New ADODB.Recordset
    rst1.Open SQL, conn, adOpenDynamic
    rst1.MoveFirst
    iScoreProfileIDMax = rst1.Fields(0).Value + 1
    rst1.Close
    Set rst1 = Nothing
    SQL = ""
    SQL = SQL & " SELECT ISNULL(MAX(iScoreDetailId),0) FROM tbSTEScoreDetail "
    Set rst1 = conn.Execute(SQL)
    rst1.MoveFirst
    iScoreDetailMax = rst1.Fields(0).Value + 1
    rst1.Close
    Set rst1 = Nothing
    Dim iSubjectQuestionId(50) As Long
    Dim iInterviewerProfileId As Long
    SQL = ""
    SQL = SQL & " SELECT c.iInterviewerProfileId as iSubjectQuestionId "
    SQL = SQL & " FROM tbSTESubjectQuestionProfile a , tbSTEInterviewRoomProfile c ,"
    SQL = SQL & "      tbSTEInterviewerProfile d"
    SQL = SQL & " Where a.iSubjectProfileId =" & iSubjectProfileID
    SQL = SQL & " And a.iSubjectProfileId = c.iSubjectProfileId"
'    sql = sql & " AND  c.iRoomProfileId =" & iRoomProfileid
    SQL = SQL & " AND  c.iRandomNo=" & intRan
    SQL = SQL & " AND  c.iNendo =" & iNendo
    SQL = SQL & " AND  c.iDayFlag =" & intDate
    SQL = SQL & " AND  d.iInterviewerProfileId = c.iInterviewerProfileId"
    SQL = SQL & " ORDER BY c.iInterviewerProfileId"
    Set rst1 = conn.Execute(SQL)
    If rst1.EOF Then
        Exit Sub
    End If
    rst1.MoveFirst
    iInterviewerProfileId = rst1.Fields(0).Value
    Dim J As Long
    While Not rst1.EOF
        iSubjectQuestionId(J) = rst1.Fields(0).Value
        rst1.MoveNext
        J = J + 1
    Wend
    rst1.Close
    Set rst1 = Nothing
    
    SQL = ""
    SQL = SQL & " Select iExamineeProfileId, dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber, vExamineeName"
    SQL = SQL & " From tbSTEExamineeProfile"
    SQL = SQL & " Where iNendo =" & iNendo
    SQL = SQL & " And iExamineeStatus = 1"
    SQL = SQL & " And iShoronbunRandomNo =" & intRan '
'    sql = sql & " AND iExamineeProfileId IN"
'    sql = sql & " (SELECT iExamineeProfileId"
'    sql = sql & "     From tbSteExamineeRoomProfile"
'    sql = sql & "   WHERE iSubjectProfileid =" & iSubjectProfileID
'    'sql = sql & " AND iRoomProfileid = " & iRoomProfileid & ")"
'    sql = sql & "   AND iRandomNo = " & intRan & ")"
'    sql = sql & " AND iInterviewerProfileId=" & iInterviewerProfileId
'    sql = sql & " AND dtSecondExamDay = '" & strTestdate & "'"
    SQL = SQL & " AND NOT EXISTS"
    SQL = SQL & " ( SELECT 1"
    SQL = SQL & " FROM tbSTEScoreProfile as s , tbSTESubjectProfile as su"
    SQL = SQL & " Where s.iExamineeProfileID = tbSTEExamineeProfile.iExamineeProfileID"
    SQL = SQL & " AND su.iSubjectProfileID = s.iSubjectProfileID"
    SQL = SQL & " AND su.iSubjectProfileid =" & iSubjectProfileID
    SQL = SQL & " AND s.iAbsentFlag = 1 )  ORDER BY iJukenNumber"
    Set RS = conn.Execute(SQL)
    If RS.EOF Then
        Exit Sub
    End If
    Dim iExamineeId As Long
    Dim I As Long

    RS.MoveFirst
      
    conn.BeginTrans
    While Not RS.EOF
        iExamineeId = RS.Fields(0).Value
        SQL = ""
        SQL = SQL & " SELECT iScoreProfileId FROM tbSTEScoreProfile"
        SQL = SQL & " WHERE iSubjectProfileId=" & iSubjectProfileID
        SQL = SQL & " AND iExamineeProfileId=" & iExamineeId
        Set rst = conn.Execute(SQL)
        '新規・更新のチェック
        If rst.EOF Then
            SQL = ""
            SQL = SQL & " INSERT INTO tbSTEScoreProfile( "
            SQL = SQL & " iScoreProfileId,iSubjectProfileId,"
            SQL = SQL & " iExamineeProfileId,fRawScore,"
            SQL = SQL & " iAbsentFlag,dtCreate,dtUpdate) "
            SQL = SQL & " SELECT " & iScoreProfileIDMax
            SQL = SQL & " , " & iSubjectProfileID & "," & iExamineeId
            'update,xzg,2008/02/12,S----------------
            'Total2の二分の一＋Total1
            'SQL = SQL & " ,iTotalScore"
            SQL = SQL & " ,(iTotalScore1 + ROUND(iTotalScore2/2,2))"
            'update,xzg,2008/02/12,E----------------
            SQL = SQL & " ,0,GETDATE(),GETDATE() "
            SQL = SQL & " FROM tbSTECompwork "
            SQL = SQL & " WHERE iDate=" & intDate
            SQL = SQL & " AND iRan=" & intRan
            SQL = SQL & " AND iNumberOfDateRan=" & I
            conn.Execute (SQL)
            blnFlagTrans = True
            '詳細を更新
            SQL = ""
            SQL = SQL & " INSERT INTO tbSTEScoreDetail( "
            SQL = SQL & " iScoreDetailId,iScoreProfileId,"
            SQL = SQL & " iSubjectQuestionId,siSerialNo,"
            SQL = SQL & " fDetailScore,dtCreate,dtUpdate) "
            SQL = SQL & " SELECT " & iScoreDetailMax & "," & iScoreProfileIDMax
            SQL = SQL & " , " & iInterviewerProfileId
            
            'update,xzg,2008/02/12,S----------------
            'Total2の二分の一＋Total1
            'SQL = SQL & " ,iTotalScore"
            SQL = SQL & " ,1,(iTotalScore1 + ROUND(iTotalScore2/2,2))"
            'update,xzg,2008/02/12,E----------------
            SQL = SQL & " ,GETDATE(),GETDATE() "
            SQL = SQL & " FROM tbSTECompwork "
            SQL = SQL & " WHERE iDate=" & intDate
            SQL = SQL & " AND iRan=" & intRan
            SQL = SQL & " AND iNumberOfDateRan=" & I
            conn.Execute (SQL)
            iScoreDetailMax = iScoreDetailMax + 1
            iScoreProfileIDMax = iScoreProfileIDMax + 1
        Else
            SQL = ""
            SQL = SQL & " UPDATE tbSTEScoreProfile "
            'update,xzg,2008/02/12,S----------------
            'Total2の二分の一＋Total1
            'SQL = SQL & " SET fRawScore=cp.iTotalScore "
            SQL = SQL & " SET fRawScore=(cp.iTotalScore1 + ROUND(cp.iTotalScore2/2,2)) "
            'update,xzg,2008/02/12,E----------------
            SQL = SQL & " ,dtUpdate=GETDATE() "
            SQL = SQL & " FROM tbSTEScoreProfile sp,tbSTECompwork cp"
            SQL = SQL & " WHERE sp.iScoreProfileId=" & rst.Fields(0).Value
            SQL = SQL & " AND cp.iDate=" & intDate
            SQL = SQL & " AND cp.iRan=" & intRan
            SQL = SQL & " AND cp.iNumberOfDateRan=" & I
            conn.Execute (SQL)
            blnFlagTrans = True
            '詳細を更新
            SQL = ""
            SQL = SQL & " UPDATE tbSTEScoreDetail "
            'update,xzg,2008/02/12,S----------------
            'Total2の二分の一＋Total1
            'SQL = SQL & " SET fDetailScore=cp.iTotalScore "
            SQL = SQL & " SET fDetailScore=(cp.iTotalScore1 + ROUND(cp.iTotalScore2/2,2)) "
            'update,xzg,2008/02/12,E----------------
            SQL = SQL & " ,dtUpdate=GETDATE() "
            SQL = SQL & " FROM tbSTEScoreDetail sp,tbSTECompwork cp"
            SQL = SQL & " WHERE sp.iScoreProfileId=" & rst.Fields(0).Value
'            sql = sql & " AND sp.siSerialNo=1"
            SQL = SQL & " AND sp.iSubjectQuestionId=" & iInterviewerProfileId
            SQL = SQL & " AND cp.iDate=" & intDate
            SQL = SQL & " AND cp.iRan=" & intRan
            SQL = SQL & " AND cp.iNumberOfDateRan=" & I
            conn.Execute (SQL)
        End If
        
        rst.Close
        Set rst = Nothing
        RS.MoveNext
        I = I + 1
    Wend
    RS.Close
    Set RS = Nothing
    
    '3.終了処理
    Set RS = Nothing
    conn.CommitTrans
    blnFlagTrans = False
    conn.Close
    blnUpdatedFlag = False
    
    
Exit Sub
ERR_HANDLE:
    MsgBox Err.Description
    If blnFlagTrans = True Then
        conn.RollbackTrans
        blnFlagTrans = False
    End If
    Set RS = Nothing
    If conn Is Nothing Then
    Else
'        conn.Close
    End If
    
End Sub

Private Sub cmdUpdate_Click()

Dim intAddition1            As Long  '1回目調整数
Dim intAddition2            As Long  '2回目調整数
Dim I                       As Long
On Error GoTo ERR_HANDLE
    For I = 2 To 151
    
        '1回目調整数設定
        If Trim(Me.txtTest.Text) = "0" Then
            
            intAddition1 = Val(Me.txtFirstDayFirst.Text)
        ElseIf Trim(Me.txtTest.Text) = "1" Then
            intAddition1 = Val(Me.txtSecondDayFirst.Text)
        End If
            
        
        If Len(Me.MFGrid.TextMatrix(I, 4)) > 0 Then
            Me.MFGrid.TextMatrix(I, 3) = intAddition1
            '1回目得点数設定
            Me.MFGrid.TextMatrix(I, 2) = Val(Me.MFGrid.TextMatrix(I, 3)) + Val(Me.MFGrid.TextMatrix(I, 4))
        End If
        
        
        
        '2目調整数設定
        If Trim(Me.txtTest.Text) = "0" Then
            
            intAddition2 = Val(Me.txtFirstDaySecond.Text)
        ElseIf Trim(Me.txtTest.Text) = "1" Then
            intAddition2 = Val(Me.txtSecondDaySecond.Text)
        End If
        
        If Len(Me.MFGrid.TextMatrix(I, 12)) > 0 Then
            Me.MFGrid.TextMatrix(I, 11) = intAddition2
            '2回目得点数設定
            Me.MFGrid.TextMatrix(I, 10) = Val(Me.MFGrid.TextMatrix(I, 11)) + Val(Me.MFGrid.TextMatrix(I, 12))
        End If
        
        '得点数設定
        If Len(Me.MFGrid.TextMatrix(I, 2)) > 0 Or Len(Me.MFGrid.TextMatrix(I, 10)) > 0 Then
            Me.MFGrid.TextMatrix(I, 1) = Val(Me.MFGrid.TextMatrix(I, 2)) + Val(Me.MFGrid.TextMatrix(I, 10))
        End If
    Next
Exit Sub
ERR_HANDLE:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Call Grid_init
End Sub


Private Sub MFGrid_Click()

Dim intCurrentRow As Long   'current row
Dim intCurrentCol As Long   'current col
Dim intGridLeft   As Long   'grid  left
Dim intGridTop    As Long   'grid top

On Error GoTo ERR_HANDLE
    '編集用のテキストボックスを初期化する
    Me.txtInput.Visible = False
    Me.txtInput.Text = ""
        
    intCurrentRow = Me.MFGrid.Row   'get current row
    intCurrentCol = Me.MFGrid.Col   'get current col
    intPreGridRow = intCurrentRow   'get preview row
    intPreGridCol = intCurrentCol   'get preview col

    '2<=row <=151,(5<=col<=9 or 13<=col<=17)編集可
    If intCurrentRow > 1 And intCurrentRow < 152 And _
        ((intCurrentCol > 4 And intCurrentCol < 10) _
        Or (intCurrentCol > 12 And intCurrentCol < 18)) Then
        'gridの(x,y)を取得する
        intGridLeft = Me.frmList.Left + Me.MFGrid.Left
        intGridTop = Me.frmList.Top + Me.MFGrid.Top
    
        '編集用のテキストボックスが見える
        '前の採点あって、前の行にデータある
        If intCurrentRow = 2 Then
            If intCurrentCol = 5 Or intCurrentCol = 13 Then
            
                Me.txtInput.Visible = True
            Else
                If Len(Me.MFGrid.TextMatrix(intCurrentRow, intCurrentCol - 1)) > 0 Then
                    Me.txtInput.Visible = True
                End If
            End If
        ElseIf intCurrentRow > 2 Then
            If Len(Me.MFGrid.TextMatrix(intCurrentRow - 1, 1)) > 0 Then
                If intCurrentCol = 5 Or intCurrentCol = 13 Then
                    Me.txtInput.Visible = True
                Else
                    If Len(Me.MFGrid.TextMatrix(intCurrentRow, intCurrentCol - 1)) > 0 Then
                        Me.txtInput.Visible = True
                    End If
                End If
            End If
            
        End If
        '編集用のテキストボックスを当前Itemに置く
        If Me.txtInput.Visible = True Then
            Me.txtInput.Left = intGridLeft + Me.MFGrid.CellLeft
            Me.txtInput.Top = intGridTop + Me.MFGrid.CellTop
            Me.txtInput.Width = Me.MFGrid.CellWidth - 20
            Me.txtInput.Height = Me.MFGrid.CellHeight - 20
            Me.txtInput.Text = Me.MFGrid.TextMatrix(intCurrentRow, intCurrentCol)
            Me.txtInput.Visible = True
            Me.txtInput.SetFocus
        End If
        
    End If
    
    
Exit Sub
ERR_HANDLE:
    MsgBox Err.Description
End Sub

Private Sub MFGrid_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        
        VBA.SendKeys "{Down}"
        
        If intPreGridRow > 1 And intPreGridRow < 151 Then
            Me.MFGrid.Row = intPreGridRow + 1
            Me.MFGrid.Col = intPreGridCol
            intPreGridRow = Me.MFGrid.Row
        End If
    Else
        intPreGridRow = Me.MFGrid.Row
        intPreGridCol = Me.MFGrid.Col
    End If
    
End Sub

Private Sub MFGrid_Scroll()
    Me.txtInput.Visible = False
    Me.MFGrid.SetFocus
End Sub

Private Sub txtFirstDayFirst_GotFocus()
    Me.txtFirstDayFirst.SelStart = 0
    Me.txtFirstDayFirst.SelLength = Me.txtFirstDayFirst.MaxLength
End Sub

Private Sub txtFirstDaySecond_GotFocus()
    Me.txtFirstDaySecond.SelStart = 0
    Me.txtFirstDaySecond.SelLength = Me.txtFirstDaySecond.MaxLength
End Sub

Private Sub txtInput_GotFocus()
    Me.txtInput.SelStart = 0
    Me.txtInput.SelLength = Me.txtInput.MaxLength
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
       
    If KeyCode = 13 Then
        VBA.SendKeys "{Down}"
        If intPreGridRow > 1 And intPreGridRow < 151 Then
            Call txtInput_LostFocus
            DoEvents
            Me.MFGrid.Row = intPreGridRow + 1
            Me.MFGrid.Col = intPreGridCol
            intPreGridRow = intPreGridRow + 1
            Me.MFGrid.SetFocus
            Call MFGrid_Click
        End If
    End If
End Sub

Private Sub txtInput_LostFocus()

Dim intGetDot1              As Double   '1回目得点
Dim intAddition1            As Long  '1回目調整数
Dim intGetAverage1          As Double   '1回目平均数

Dim intGetDot2              As Double   '2回目得点
Dim intAddition2            As Long  '2回目調整数
Dim intGetAverage2          As Double   '2回目平均数

Dim intDeno                 As Long  '分母
Dim I                       As Long
On Error GoTo ERR_HANDLE
    If Me.txtInput.Visible = False Then
        Exit Sub
    End If
  
    If Val(Me.txtInput.Text) < 0 Then
        Me.txtInput.Visible = False
        Exit Sub
    End If
    

    
    
    'GridのCellのデータ設定
    
    Me.MFGrid.TextMatrix(intPreGridRow, intPreGridCol) = Me.txtInput.Text
    
    If (intPreGridCol > 4 And intPreGridCol < 10) Then

        intDeno = 0
        For I = 5 To 9
            If Len(Me.MFGrid.TextMatrix(intPreGridRow, I)) > 0 Then
                intDeno = intDeno + 1
            End If
        Next
        
        If intDeno > 0 Then
            '1回目平均数
            intGetAverage1 = (Val(Me.MFGrid.TextMatrix(intPreGridRow, 5)) + _
                            Val(Me.MFGrid.TextMatrix(intPreGridRow, 6)) + _
                            Val(Me.MFGrid.TextMatrix(intPreGridRow, 7)) + _
                            Val(Me.MFGrid.TextMatrix(intPreGridRow, 8)) + _
                            Val(Me.MFGrid.TextMatrix(intPreGridRow, 9))) / intDeno
            
             intGetAverage1 = Round(intGetAverage1, 1)
            '1回目調整数
            
            If Trim(Me.txtTest.Text) = "0" Then
            
                intAddition1 = Val(Me.txtFirstDayFirst.Text)
            ElseIf Trim(Me.txtTest.Text) = "1" Then
                intAddition1 = Val(Me.txtSecondDayFirst.Text)
            End If
            
            '1回目得点
            intGetDot1 = intAddition1 + intGetAverage1
            
            '1回目平均数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 4) = intGetAverage1
            '1回目調整数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 3) = intAddition1
            '1回目得点数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 2) = intGetDot1
        Else
            '1回目平均数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 4) = ""
            '1回目調整数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 3) = ""
            '1回目得点数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 2) = ""
            Me.MFGrid.TextMatrix(intPreGridRow, 1) = Val(Me.MFGrid.TextMatrix(intPreGridRow, 2)) + Val(Me.MFGrid.TextMatrix(intPreGridRow, 10))
        End If
    ElseIf intPreGridCol > 12 And intPreGridCol < 18 Then
        
        intDeno = 0
        For I = 13 To 17
            If Len(Me.MFGrid.TextMatrix(intPreGridRow, I)) > 0 Then
                intDeno = intDeno + 1
            End If
        Next
        
        If intDeno > 0 Then
            '2回目平均数
            intGetAverage2 = (Val(Me.MFGrid.TextMatrix(intPreGridRow, 13)) + _
                            Val(Me.MFGrid.TextMatrix(intPreGridRow, 14)) + _
                            Val(Me.MFGrid.TextMatrix(intPreGridRow, 15)) + _
                            Val(Me.MFGrid.TextMatrix(intPreGridRow, 16)) + _
                            Val(Me.MFGrid.TextMatrix(intPreGridRow, 17))) / intDeno
        
            
            intGetAverage2 = Round(intGetAverage2, 1)
            '2回目調整数
            If Trim(Me.txtTest.Text) = "0" Then
            
                intAddition2 = Val(Me.txtFirstDaySecond.Text)
            ElseIf Trim(Me.txtTest.Text) = "1" Then
                intAddition2 = Val(Me.txtSecondDaySecond.Text)
            End If
            '2回目得点
            intGetDot2 = intAddition2 + intGetAverage2
            
            '2回目平均数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 12) = intGetAverage2
            '2回目調整数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 11) = intAddition2
            '2回目得点数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 10) = intGetDot2
        Else
            '2回目平均数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 12) = ""
            '2回目調整数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 11) = ""
            '2回目得点数設定
            Me.MFGrid.TextMatrix(intPreGridRow, 10) = ""
            
            Me.MFGrid.TextMatrix(intPreGridRow, 1) = Val(Me.MFGrid.TextMatrix(intPreGridRow, 2)) + Val(Me.MFGrid.TextMatrix(intPreGridRow, 10))
            
        End If
    End If
    If Len(Me.MFGrid.TextMatrix(intPreGridRow, 2)) > 0 Or Len(Me.MFGrid.TextMatrix(intPreGridRow, 10)) > 0 Then
        Me.MFGrid.TextMatrix(intPreGridRow, 1) = Val(Me.MFGrid.TextMatrix(intPreGridRow, 2)) + Val(Me.MFGrid.TextMatrix(intPreGridRow, 10))
    Else
        If Len(Me.MFGrid.TextMatrix(intPreGridRow, 2)) <= 0 And Len(Me.MFGrid.TextMatrix(intPreGridRow, 10)) <= 0 Then
            Me.MFGrid.TextMatrix(intPreGridRow, 1) = ""
        End If
    End If
    '編集用のテキストボックスを初期化する
    Me.txtInput.Visible = False
    Me.txtInput.Text = ""
Exit Sub
ERR_HANDLE:
    MsgBox Err.Description
    
End Sub

Private Sub txtRand_GotFocus()
    Me.txtRand.SelStart = 0
    Me.txtRand.SelLength = Me.txtRand.MaxLength
End Sub

Private Sub txtRand_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim conn    As ADODB.Connection
Dim RS      As ADODB.Recordset
Dim DSN     As String
Dim uid     As String
Dim pas     As String
Dim objFS   As Object
Dim objText As Object
Dim SQL     As String
Dim intRan  As Long
Dim intTest As Long
Dim I       As Long
Dim J       As Long
Dim intTemp As Double
On Error GoTo ERR_HANDLE
    If KeyCode <> 13 Then
        Exit Sub
    End If
    
    
    If Len(Trim(Me.txtTest.Text)) > 0 And Len(Trim(Me.txtRand.Text)) > 0 Then
        If Trim(Me.txtTest.Text) <> "0" And Trim(Me.txtTest.Text) <> "1" Then
            Exit Sub
        End If
        If Not IsNumeric(Trim(Me.txtRand.Text)) Then
            MsgBox "数字を入力してください"
            Me.txtTest.SetFocus
            Exit Sub
        End If
        
        intRan = Trim(Me.txtRand.Text)
        intTest = Trim(Me.txtTest.Text)
        'Grid初期化
        Call Grid_init
        
        Set objFS = CreateObject("Scripting.FileSystemObject")
        Set objText = objFS.OpenTextFile("odbc.txt")
        DSN = objText.ReadLine
        uid = objText.ReadLine
        pas = objText.ReadLine
        Set objFS = Nothing
                        
        'DBの接続
        Set conn = New ADODB.Connection
        conn.Open DSN, uid, pas
        SQL = ""
        SQL = SQL & "SELECT * FROM tbSTECompwork"
        SQL = SQL & " WHERE iDate=" & intTest
        SQL = SQL & " AND   iRan=" & intRan
        SQL = SQL & " ORDER BY   iNumberOfDateRan"
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn
        If Not RS.EOF Then
            I = 1
            RS.MoveFirst
            
            While Not RS.EOF
                
                If Not IsNull(RS.Fields("iTotalScore").Value) Then
                    intTemp = RS.Fields("iTotalScore").Value
                    Me.MFGrid.TextMatrix(I + 1, 1) = intTemp
                    
                    If Not IsNull(RS.Fields("iTotalScore1").Value) Then
                        intTemp = RS.Fields("iTotalScore1").Value
                        Me.MFGrid.TextMatrix(I + 1, 2) = intTemp
                        
                        intTemp = RS.Fields("iChoScore1").Value
                        Me.MFGrid.TextMatrix(I + 1, 3) = intTemp
                                             
                        
                        intTemp = RS.Fields("iAveScore1").Value
                        Me.MFGrid.TextMatrix(I + 1, 4) = intTemp
                        
                        If Me.txtTest.Text = "0" Then
                            If Len(Trim(Me.txtFirstDayFirst.Text)) > 0 Then
                                Me.MFGrid.TextMatrix(I + 1, 3) = Trim(Me.txtFirstDayFirst.Text)
                                Me.MFGrid.TextMatrix(I + 1, 2) = Val(Me.MFGrid.TextMatrix(I + 1, 3)) + Val(intTemp)
                                Me.MFGrid.TextMatrix(I + 1, 1) = Val(Me.MFGrid.TextMatrix(I + 1, 2)) + Val(Me.MFGrid.TextMatrix(I + 1, 10))
                            End If
                        ElseIf Me.txtTest.Text = "1" Then
                            If Len(Trim(Me.txtSecondDayFirst.Text)) > 0 Then
                                Me.MFGrid.TextMatrix(I + 1, 3) = Trim(Me.txtSecondDayFirst.Text)
                                Me.MFGrid.TextMatrix(I + 1, 2) = Val(Me.MFGrid.TextMatrix(I + 1, 3)) + Val(intTemp)
                                Me.MFGrid.TextMatrix(I + 1, 1) = Val(Me.MFGrid.TextMatrix(I + 1, 2)) + Val(Me.MFGrid.TextMatrix(I + 1, 10))
                            End If
                        End If
                        
                        If Not IsNull(RS.Fields("iP1Score1").Value) Then
                            intTemp = RS.Fields("iP1Score1").Value
                            Me.MFGrid.TextMatrix(I + 1, 5) = intTemp
                        End If
                        If Not IsNull(RS.Fields("iP2Score1").Value) Then
                            intTemp = RS.Fields("iP2Score1").Value
                            Me.MFGrid.TextMatrix(I + 1, 6) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP3Score1").Value) Then
                            intTemp = RS.Fields("iP3Score1").Value
                            Me.MFGrid.TextMatrix(I + 1, 7) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP4Score1").Value) Then
                            intTemp = RS.Fields("iP4Score1").Value
                            Me.MFGrid.TextMatrix(I + 1, 8) = intTemp
                        End If
                        If Not IsNull(RS.Fields("iP5Score1").Value) Then
                            intTemp = RS.Fields("iP5Score1").Value
                            Me.MFGrid.TextMatrix(I + 1, 9) = intTemp
                        End If
                    End If
                    If Not IsNull(RS.Fields("iTotalScore2").Value) Then
                        intTemp = RS.Fields("iTotalScore2").Value
                        Me.MFGrid.TextMatrix(I + 1, 10) = intTemp
                        
                        intTemp = RS.Fields("iChoScore2").Value
                        Me.MFGrid.TextMatrix(I + 1, 11) = intTemp
                        
                                            
                        intTemp = RS.Fields("iAveScore2").Value
                        Me.MFGrid.TextMatrix(I + 1, 12) = intTemp
                        
                        If Me.txtTest.Text = "0" Then
                            If Len(Trim(Me.txtFirstDaySecond.Text)) > 0 Then
                                Me.MFGrid.TextMatrix(I + 1, 11) = Trim(Me.txtFirstDaySecond.Text)
                                Me.MFGrid.TextMatrix(I + 1, 10) = Val(Me.MFGrid.TextMatrix(I + 1, 11)) + intTemp
                                Me.MFGrid.TextMatrix(I + 1, 1) = Val(Me.MFGrid.TextMatrix(I + 1, 2)) + Val(Me.MFGrid.TextMatrix(I + 1, 10))
                            End If
                        ElseIf Me.txtTest.Text = "1" Then
                            If Len(Trim(Me.txtFirstDaySecond.Text)) > 0 Then
                                Me.MFGrid.TextMatrix(I + 1, 11) = Trim(Me.txtFirstDaySecond.Text)
                                Me.MFGrid.TextMatrix(I + 1, 10) = Val(Me.MFGrid.TextMatrix(I + 1, 11)) + intTemp
                                Me.MFGrid.TextMatrix(I + 1, 1) = Val(Me.MFGrid.TextMatrix(I + 1, 2)) + Val(Me.MFGrid.TextMatrix(I + 1, 10))
                            End If
                        End If
                        
                        If Not IsNull(RS.Fields("iP1Score2").Value) Then
                            intTemp = RS.Fields("iP1Score2").Value
                            Me.MFGrid.TextMatrix(I + 1, 13) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP2Score2").Value) Then
                            intTemp = RS.Fields("iP2Score2").Value
                            Me.MFGrid.TextMatrix(I + 1, 14) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP3Score2").Value) Then
                            intTemp = RS.Fields("iP3Score2").Value
                            Me.MFGrid.TextMatrix(I + 1, 15) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP4Score2").Value) Then
                            intTemp = RS.Fields("iP4Score2").Value
                            Me.MFGrid.TextMatrix(I + 1, 16) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP5Score2").Value) Then
                            intTemp = RS.Fields("iP5Score2").Value
                            Me.MFGrid.TextMatrix(I + 1, 17) = intTemp
                        End If
                    End If
                    
                End If
                RS.MoveNext
                I = I + 1
                intTemp = 0
            Wend
        End If
        Set RS = Nothing
        Set conn = Nothing
    End If
Exit Sub
ERR_HANDLE:
    Set RS = Nothing
    Set conn = Nothing
    MsgBox Err.Description
End Sub

Private Sub txtSecondDayFirst_GotFocus()
    Me.txtSecondDayFirst.SelStart = 0
    Me.txtSecondDayFirst.SelLength = Me.txtSecondDayFirst.MaxLength
End Sub

Private Sub txtSecondDaySecond_GotFocus()
    Me.txtSecondDaySecond.SelStart = 0
    Me.txtSecondDaySecond.SelLength = Me.txtSecondDaySecond.MaxLength
End Sub

Private Sub txtTest_GotFocus()
    Me.txtTest.SelStart = 0
    Me.txtTest.SelLength = Me.txtTest.MaxLength
End Sub

Private Sub txtTest_KeyDown(KeyCode As Integer, Shift As Integer)
Dim conn    As ADODB.Connection
Dim RS      As ADODB.Recordset
Dim DSN     As String
Dim uid     As String
Dim pas     As String
Dim objFS   As Object
Dim objText As Object
Dim SQL     As String
Dim intRan  As Long
Dim intTest As Long
Dim I       As Long
Dim J       As Long
Dim intTemp As Double
On Error GoTo ERR_HANDLE
    If KeyCode <> 13 Then
        Exit Sub
    End If
    
    If Len(Trim(Me.txtTest.Text)) > 0 And Len(Trim(Me.txtRand.Text)) > 0 Then
        If Trim(Me.txtTest.Text) <> "0" And Trim(Me.txtTest.Text) <> "1" Then
            Exit Sub
        End If
    
        If Not IsNumeric(Trim(Me.txtTest.Text)) Then
            MsgBox "数字を入力してください"
            Me.txtTest.SetFocus
            Exit Sub
        End If
        
        intRan = Trim(Me.txtRand.Text)
        intTest = Trim(Me.txtTest.Text)
        'Grid初期化
        Call Grid_init
        
        Set objFS = CreateObject("Scripting.FileSystemObject")
        Set objText = objFS.OpenTextFile("odbc.txt")
        DSN = objText.ReadLine
        uid = objText.ReadLine
        pas = objText.ReadLine
        Set objFS = Nothing
                        
        'DBの接続
        Set conn = New ADODB.Connection
        conn.Open DSN, uid, pas
        SQL = ""
        SQL = SQL & "SELECT * FROM tbSTECompwork"
        SQL = SQL & " WHERE iDate=" & intTest
        SQL = SQL & " AND   iRan=" & intRan
        SQL = SQL & " ORDER BY   iNumberOfDateRan"
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn
        If Not RS.EOF Then
            I = 1
            RS.MoveFirst
            
            While Not RS.EOF
                
                If Not IsNull(RS.Fields("iTotalScore").Value) Then
                    intTemp = RS.Fields("iTotalScore").Value
                    Me.MFGrid.TextMatrix(I + 1, 1) = intTemp
                    
                    If Not IsNull(RS.Fields("iTotalScore1").Value) Then
                        intTemp = RS.Fields("iTotalScore1").Value
                        Me.MFGrid.TextMatrix(I + 1, 2) = intTemp
                        
                        intTemp = RS.Fields("iChoScore1").Value
                        Me.MFGrid.TextMatrix(I + 1, 3) = intTemp
                        
                        intTemp = RS.Fields("iAveScore1").Value
                        Me.MFGrid.TextMatrix(I + 1, 4) = intTemp
                        
                        If Me.txtTest.Text = "0" Then
                            If Len(Trim(Me.txtFirstDayFirst.Text)) > 0 Then
                                Me.MFGrid.TextMatrix(I + 1, 3) = Trim(Me.txtFirstDayFirst.Text)
                                Me.MFGrid.TextMatrix(I + 1, 2) = Val(Me.MFGrid.TextMatrix(I + 1, 3)) + Val(intTemp)
                                Me.MFGrid.TextMatrix(I + 1, 1) = Val(Me.MFGrid.TextMatrix(I + 1, 2)) + Val(Me.MFGrid.TextMatrix(I + 1, 10))
                            End If
                        ElseIf Me.txtTest.Text = "1" Then
                            If Len(Trim(Me.txtSecondDayFirst.Text)) > 0 Then
                                Me.MFGrid.TextMatrix(I + 1, 3) = Trim(Me.txtSecondDayFirst.Text)
                                Me.MFGrid.TextMatrix(I + 1, 2) = Val(Me.MFGrid.TextMatrix(I + 1, 3)) + Val(intTemp)
                                Me.MFGrid.TextMatrix(I + 1, 1) = Val(Me.MFGrid.TextMatrix(I + 1, 2)) + Val(Me.MFGrid.TextMatrix(I + 1, 10))
                            End If
                        End If
                        
                        If Not IsNull(RS.Fields("iP1Score1").Value) Then
                            intTemp = RS.Fields("iP1Score1").Value
                            Me.MFGrid.TextMatrix(I + 1, 5) = intTemp
                        End If
                        If Not IsNull(RS.Fields("iP2Score1").Value) Then
                            intTemp = RS.Fields("iP2Score1").Value
                            Me.MFGrid.TextMatrix(I + 1, 6) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP3Score1").Value) Then
                            intTemp = RS.Fields("iP3Score1").Value
                            Me.MFGrid.TextMatrix(I + 1, 7) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP4Score1").Value) Then
                            intTemp = RS.Fields("iP4Score1").Value
                            Me.MFGrid.TextMatrix(I + 1, 8) = intTemp
                        End If
                        If Not IsNull(RS.Fields("iP5Score1").Value) Then
                            intTemp = RS.Fields("iP5Score1").Value
                            Me.MFGrid.TextMatrix(I + 1, 9) = intTemp
                        End If
                    End If
                    If Not IsNull(RS.Fields("iTotalScore2").Value) Then
                        intTemp = RS.Fields("iTotalScore2").Value
                        Me.MFGrid.TextMatrix(I + 1, 10) = intTemp
                        
                        intTemp = RS.Fields("iChoScore2").Value
                        Me.MFGrid.TextMatrix(I + 1, 11) = intTemp
                        
                        intTemp = RS.Fields("iAveScore2").Value
                        Me.MFGrid.TextMatrix(I + 1, 12) = intTemp
                        
                        If Me.txtTest.Text = "0" Then
                            If Len(Trim(Me.txtFirstDaySecond.Text)) > 0 Then
                                Me.MFGrid.TextMatrix(I + 1, 11) = Trim(Me.txtFirstDaySecond.Text)
                                Me.MFGrid.TextMatrix(I + 1, 10) = Val(Me.MFGrid.TextMatrix(I + 1, 11)) + intTemp
                                Me.MFGrid.TextMatrix(I + 1, 1) = Val(Me.MFGrid.TextMatrix(I + 1, 2)) + Val(Me.MFGrid.TextMatrix(I + 1, 10))
                            End If
                        ElseIf Me.txtTest.Text = "1" Then
                            If Len(Trim(Me.txtFirstDaySecond.Text)) > 0 Then
                                Me.MFGrid.TextMatrix(I + 1, 11) = Trim(Me.txtFirstDaySecond.Text)
                                Me.MFGrid.TextMatrix(I + 1, 10) = Val(Me.MFGrid.TextMatrix(I + 1, 11)) + intTemp
                                Me.MFGrid.TextMatrix(I + 1, 1) = Val(Me.MFGrid.TextMatrix(I + 1, 2)) + Val(Me.MFGrid.TextMatrix(I + 1, 10))
                            End If
                        End If
                        
                        If Not IsNull(RS.Fields("iP1Score2").Value) Then
                            intTemp = RS.Fields("iP1Score2").Value
                            Me.MFGrid.TextMatrix(I + 1, 13) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP2Score2").Value) Then
                            intTemp = RS.Fields("iP2Score2").Value
                            Me.MFGrid.TextMatrix(I + 1, 14) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP3Score2").Value) Then
                            intTemp = RS.Fields("iP3Score2").Value
                            Me.MFGrid.TextMatrix(I + 1, 15) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP4Score2").Value) Then
                            intTemp = RS.Fields("iP4Score2").Value
                            Me.MFGrid.TextMatrix(I + 1, 16) = intTemp
                        End If
                        
                        If Not IsNull(RS.Fields("iP5Score2").Value) Then
                            intTemp = RS.Fields("iP5Score2").Value
                            Me.MFGrid.TextMatrix(I + 1, 17) = intTemp
                        End If
                    End If
                    
                End If
                RS.MoveNext
                I = I + 1
                intTemp = 0
            Wend
        End If
        Set RS = Nothing
        Set conn = Nothing
    End If
Exit Sub
ERR_HANDLE:
    Set RS = Nothing
    Set conn = Nothing
    MsgBox Err.Description
End Sub
