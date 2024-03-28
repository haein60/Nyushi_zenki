VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBrowse 
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmBrowse.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   12960
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "キャンセル"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6630
      TabIndex        =   3
      Top             =   2505
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "データをアップロード"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3405
      TabIndex        =   2
      Top             =   2505
      Width           =   2775
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   1320
      Width           =   450
   End
   Begin VB.TextBox txtCSVPath 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog cdlCSVPath 
      Left            =   240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select CSV File"
      Filter          =   "TEXT Files (*.txt)|*.txt|その他テキストファイル(*)|*.*|"
   End
   Begin VB.Label lblErrorDetails 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   8010
   End
   Begin VB.Label lblCSVPath 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "Web出願データファイルを選択"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   525
      TabIndex        =   5
      Top             =   1365
      Width           =   3495
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmBrowse
'Author         :   Vishal Kamath
'Created On     :   17/9/2001
'Description    :   This form allows user to pick up the data file from which examinee data has to me inserted into the database table
'Reference      :   Functional Specs Of Read From OCR Data Ver 1.0
'**************************************************************************************************

'Dim f_obj_DummyDll As New UpdateExaminee.clsUpdateExaminee  'DLL which updates exmainee details
Dim f_bln_ReturnVal As Boolean  ' to check the status of DLL operation
Private Declare Function dcvConvert Lib "dataconv.dll" _
                (ByVal iniFile As String, ByVal params As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Const prvcsProfileName         As String = "D:\Comdesign\Spp\StMari\Nyushi\Etc\STMari.Ini"
Private Const prvcsOCRDTConvIniFileDef As String = "D:\Comdesign\Spp\StMari\Nyushi\Etc\STEOCRCNV.Ini"
Private prvsOCRDTConvIniFile           As String

''''Private Const prvsOCRFilePathDef   As String = "R:\"
Private Const prvsOCRFilePathDef       As String = "C:\"      'iniファイルを取得できなかった場合defaultで使用する設定 2021.12.03
Private prvsOCRFilePath                As String

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    LoadResStrings Me

    Me.Caption = "frmBrowse : Web出願データ取込み"    ''''LoadResString(1731)

    Call g_void_SetFontProperties(Me)                 'set the font properties

    Call getIniData
    txtCSVPath.Text = prvsOCRFilePath

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

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

Private Sub getIniData()

    Dim sProfileName As String
    Dim sRtn         As String
    Dim lRtn         As Long


    '初期化ファイルを復号
    If Right(App.Path, 1) = "\" Then
        sProfileName = App.Path & prvcsProfileName
    Else
        sProfileName = App.Path & "\" & prvcsProfileName
    End If

    '初期化ファイルを読取
    sRtn = Space(255): lRtn = GetPrivateProfileString("OCRDATA", "INIFILEPATH_A_NAME", prvcsOCRDTConvIniFileDef, sRtn, 255, sProfileName)
    If lRtn > 0 Then prvsOCRDTConvIniFile = lf_StrNullCut(sRtn)

    sRtn = Space(255)
    lRtn = GetPrivateProfileString("OCRDATA", "OCRFILEPATH", prvsOCRFilePathDef, sRtn, 255, sProfileName)

    If lRtn > 0 Then
        prvsOCRFilePath = lf_StrNullCut(sRtn)
    End If

End Sub

Private Sub cmdBrowse_Click()

    On Error GoTo ErrorHandler
    Err.Clear
    cdlCSVPath.ShowOpen

    ' check for cancel error
    If Err.Number = 0 Then
        txtCSVPath.Text = Left(cdlCSVPath.FileName, InStrRev(cdlCSVPath.FileName, "\"))
    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub cmdCancel_Click()

    On Error GoTo ErrorHandler

    txtCSVPath.Text = ""
    txtCSVPath.SetFocus
    lblErrorDetails.Visible = False

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorHandler

    If txtCSVPath.Text = "" Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "ファイルを指定してください。" ''''LoadResString(1730)
    Else
        lblErrorDetails.Visible = False
        lblErrorDetails.Caption = ""
        cmdOK.Enabled = False
        Call f_void_SendData(Trim(txtCSVPath.Text))
        cmdOK.Enabled = True
    End If

    Exit Sub

ErrorHandler:
    cmdOK.Enabled = True
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler

    fMainForm.mnuTools.Enabled = False  ' disable tools menu

    Dim Index
    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub


Private Sub f_void_SendData(ByVal psFilePath As String)

    On Error GoTo ErrorHandler
    
    Dim iniFile         As String
    Dim param           As String
    Dim f_bln_ReturnVal As Integer
    Dim yy              As Integer
    Dim sFilePath       As String
    Dim sFileName       As String
    Dim sCopyFile       As String


    'XXXXX残件
    'tbSETSystemProfileのiActive=1のiNendoを設定

     yy = g_int_CurrentNendo

    iniFile = prvsOCRDTConvIniFile

    If Right(psFilePath, 1) <> "\" Then
        sFilePath = psFilePath & "\"
    Else
        sFilePath = psFilePath
    End If

    f_bln_ReturnVal = 1

    sFileName = Dir(sFilePath, vbNormal)   ' 最初のフォルダ名を返します。

    Do While sFileName <> ""   ' ループを開始します。

        ' 現在のフォルダと親フォルダは無視します。
        If sFileName <> "." And sFileName <> ".." Then
            ' ビット単位の比較を行い、MyName がフォルダかどうかを調べます。
            sCopyFile = sFilePath & "bak\" & sFileName
            sFileName = sFilePath & sFileName
            If (GetAttr(sFileName) And vbNormal) = vbNormal Then
                param = ";YY='" & yy & "';FILE='" & sFileName & "'"
                f_bln_ReturnVal = dcvConvert(iniFile, param)
                If f_bln_ReturnVal <> 0 Then GoTo RetData
                FileCopy sFileName, sCopyFile
                Kill sFileName
            End If
        End If
        sFileName = Dir               ' 次のフォルダ名を返します。

    Loop

RetData:

    If f_bln_ReturnVal = 0 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = LoadResString(1726)

    ElseIf f_bln_ReturnVal = -1 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "定義ファイル(" & iniFile & ")の読み込みに失敗しました。"

    ElseIf f_bln_ReturnVal = -2 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "ODBCへの接続に失敗しました。"

    ElseIf f_bln_ReturnVal = -3 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "受験者プロファイルテーブルにデータを挿入中エラーが発生しました。もう一度やってください。" ''''LoadResString(1727)

    ElseIf f_bln_ReturnVal = -4 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "入力ファイル(" & sFileName & ")の読み込みに失敗しました。"

    ElseIf f_bln_ReturnVal = -5 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "データベースのコマンド呼び出しに失敗しました。"

    ElseIf f_bln_ReturnVal = 1 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "指定のパスにファイルがありません。"
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call g_void_CloseChildForm

End Sub
