VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExcelReport 
   Caption         =   "ExcelReport"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   Picture         =   "FrmExcelReport.frx":0000
   ScaleHeight     =   11115
   ScaleWidth      =   12750
   WindowState     =   2  '最大化
   Begin MSComDlg.CommonDialog dlgFileRef 
      Left            =   5520
      Top             =   5700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOutFileRef 
      Caption         =   "..."
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   5220
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
      Left            =   180
      TabIndex        =   1
      Text            =   "txtOutFile"
      Top             =   5220
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
      Left            =   180
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "終了"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4080
      TabIndex        =   5
      Top             =   8640
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   4
      Top             =   8640
      Width           =   1515
   End
   Begin VB.ListBox lstTemplate 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
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
      Left            =   180
      TabIndex        =   7
      Top             =   4860
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "テンプレートを選択して下さい。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   4395
   End
End
Attribute VB_Name = "frmExcelReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    On Error GoTo ErrorHandler
'    fMainForm.mnuTools.Enabled = False  ' disable tools menu
    Dim Index
'    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
'       fMainForm.Toolbar1.Buttons(Index).Enabled = False
'    Next
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// Form_Load
Private Sub Form_Load()
    ' フォームを中央に配置
'    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    ' テンプレートファイル列挙
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objFolder As Scripting.Folder, objFile As Scripting.File
    Dim sName As String
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

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// cmdCancel_Click
Private Sub cmdCancel_Click()
    End
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// cmdOK_Click
Private Sub cmdOK_Click()
    ' エラーハンドル登録
    On Error GoTo ERROR_HANDLE
    cmdOK.Enabled = False
    
    ' ExcelReportMain
    Dim sTemplateFile As String, sOutputFile As String
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

