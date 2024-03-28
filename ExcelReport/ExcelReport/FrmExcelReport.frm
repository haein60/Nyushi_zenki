VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExcelReport 
   Caption         =   "frmExcelReport : Excel帳票"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmExcelReport.frx":0000
   ScaleHeight     =   10155
   ScaleWidth      =   12600
   WindowState     =   2  '最大化
   Begin MSComDlg.CommonDialog dlgFileRef 
      Left            =   11940
      Top             =   3945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOutFileRef 
      Caption         =   "..."
      Height          =   375
      Left            =   12090
      TabIndex        =   2
      Top             =   4530
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtOutFile 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6660
      TabIndex        =   1
      Text            =   "txtOutFile"
      Top             =   4545
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.CheckBox chkOpen 
      Caption         =   "ファイル作成後、エクセルで開く"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6720
      TabIndex        =   3
      Top             =   4995
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "終了"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4680
      TabIndex        =   5
      Top             =   8640
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Excel 表示"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   960
      TabIndex        =   4
      Top             =   8640
      Width           =   1815
   End
   Begin VB.ListBox lstTemplate 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   720
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1740
      Width           =   5835
   End
   Begin VB.Label Label2 
      Caption         =   "出力ファイル名"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      TabIndex        =   7
      Top             =   4860
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  '透明
      Caption         =   "印刷テンプレートを選択して下さい。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   705
      TabIndex        =   6
      Top             =   1455
      Width           =   5835
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   705
      TabIndex        =   8
      Top             =   1395
      Width           =   5835
   End
End
Attribute VB_Name = "frmExcelReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'///////////////////////////////////////////////////////////////////////////////
'// Form_Load
'///////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()

' フォームを中央に配置
'    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    ' テンプレートファイル列挙
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objFolder     As Scripting.Folder
    Dim objFile       As Scripting.File
    Dim sName         As String

    Set objFolder = objFileSystem.GetFolder(App.Path & "\Template")

    For Each objFile In objFolder.Files
        sName = objFile.Name
        If LCase(Right$(sName, 4)) = ".xls" Then lstTemplate.AddItem Left$(sName, Len(sName) - 4)
    Next

    lstTemplate.Selected(0) = True
    
    ' Init chkOpen, txtOutFile
    chkOpen.Value = 1
    txtOutFile = App.Path & "\Output.xls"

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim i As Integer

    fMainForm.mnuTools.Enabled = False                        ' disable tools menu

    For i = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'///////////////////////////////////////////////////////////////////////////////
'// cmdOK_Click : ExcelReportMain
'///////////////////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()

    ' エラーハンドル登録
    On Error GoTo ERROR_HANDLE

    Dim sTemplateFile As String
    Dim sOutputFile   As String


    cmdOK.Enabled = False
    

    sTemplateFile = App.Path & "\Template\" & lstTemplate.Text & ".xls"

    If StrConv(Right(txtOutFile, 4), vbLowerCase) = ".xls" Then
        sOutputFile = StrReverse(Mid(StrReverse(txtOutFile), 5))
    Else
        sOutputFile = txtOutFile
    End If

    ExcelReportMain sTemplateFile, sOutputFile, gsUserPwd
    
    ' Excel起動
    If chkOpen.Value = 1 Then
        DoEvents
        ShellExecute Me.hwnd, "open", sOutputFile, vbNullString, vbNullString, SW_SHOWNORMAL
    End If
    
    cmdOK.Enabled = True

    Exit Sub
    
ERROR_HANDLE:
    MsgBox Err.Description, vbCritical, "Error"
    cmdOK.Enabled = True

End Sub

Private Sub cmdOutFileRef_Click()

    dlgFileRef.FileName = txtOutFile.Text
    dlgFileRef.Flags = cdlOFNPathMustExist
    dlgFileRef.Filter = "xls file(*.xls)|*.xls|all file(*.*)|*.*||)"

    ' dlgFileRef.CancelError = True
    dlgFileRef.ShowSave
    txtOutFile.Text = dlgFileRef.FileName

End Sub

'///////////////////////////////////////////////////////////////////////////////
'// 終了 cmdCancel_Click
'///////////////////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()

    ''''End
    Unload Me

End Sub

