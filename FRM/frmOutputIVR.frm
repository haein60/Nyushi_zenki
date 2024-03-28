VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOutputIVR 
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmOutputIVR.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   12960
   WindowState     =   2  '最大化
   Begin MSComCtl2.UpDown udCount 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2880
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtCount 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "1061"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "データを転送する"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2280
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
      Height          =   390
      Left            =   7920
      TabIndex        =   1
      Top             =   1065
      Visible         =   0   'False
      Width           =   375
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
      Top             =   1080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog cdlCSVPath 
      Left            =   240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select CSV File"
      Filter          =   "CSV Files (*.*)|*.csv|"
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  '透明
      Caption         =   "繰上実施回数"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
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
      Top             =   1680
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.Label lblCSVPath 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "データ転送するファイルを選択"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmOutputIVR"
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

Private Const prvcsProfileName As String = "D:\Comdesign\Spp\StMari\Nyushi\Etc\STMari.Ini"
Private Const prvcsOCRDTConvIniFileDef As String = "D:\Comdesign\Spp\StMari\Nyushi\Etc\NYUSHI2IVR.Ini"
Private prvsOCRDTConvIniFile As String

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
Dim sRtn As String
Dim lRtn As Long

    '初期化ファイルを復号
    If Right(App.Path, 1) = "\" Then
        sProfileName = App.Path & prvcsProfileName
    Else
        sProfileName = App.Path & "\" & prvcsProfileName
    End If

    '初期化ファイルを読取
    sRtn = Space(255): lRtn = GetPrivateProfileString("IVRDATA", "INIFILEPATH_A_NAME", prvcsOCRDTConvIniFileDef, sRtn, 255, sProfileName)
    If lRtn > 0 Then prvsOCRDTConvIniFile = lf_StrNullCut(sRtn)

End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo ErrorHandler
    Err.Clear
    cdlCSVPath.ShowOpen
    ' check for cancel error
    If Err.Number = 0 Then
        txtCSVPath.Text = cdlCSVPath.FileName
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
    lblErrorDetails.Visible = False
    lblErrorDetails.Caption = ""
    cmdOK.Enabled = False
    Call f_void_SendData
    cmdOK.Enabled = True
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

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Dim sSQL   As String
    Dim oRs    As ADODB.Recordset
    Dim iCount As Integer



    LoadResStrings Me
    Me.Caption = "ＩＶＲシステムへのデータ転送"
    Call g_void_SetFontProperties(Me)     ' set the font properties
    Call getIniData

''''If fMainForm.f_int_CurrentPhase = gclPhase_WaitPass Then


        sSQL = "SELECT top 1 iWaitPassCount + 1 FROM tbSTESystemProfile where iActiveFlag = 1 order by iSystemProfileID desc "

        Set oRs = g_obj_Conn.Execute(sSQL)

        If oRs.EOF Then
            iCount = 1
        Else
            iCount = oRs.Fields(0)
            oRs.Close
        End If

        Set oRs = Nothing

        txtCount.Visible = True
        lblCount.Visible = True
        udCount.Visible = True
        txtCount.Text = Trim(str(iCount))

    Else

        txtCount.Visible = False
        lblCount.Visible = False
        udCount.Visible = False

    End If

Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_SendData()
    On Error GoTo ErrorHandler
    
    Dim iniFile As String
    Dim param As String
    Dim f_bln_ReturnVal As Integer
    Dim yy As Integer
    Dim iCount As Integer
    Dim sWk As String
    Dim sSQL As String
    Dim iStatus  As Integer

'XXXXX残件
'tbSETSystemProfileのiActive=1のiNendoを設定
    yy = g_int_CurrentNendo

    If fMainForm.f_int_CurrentPhase = gclPhase_WaitPass Then
        sWk = Trim(txtCount.Text)
        If gf_IntCheck(sWk) Then
            iCount = CInt(sWk)
        Else
            iCount = 1
        End If
        sSQL = "update tbSTESystemProfile set iWaitPassCount = " & Trim(str(iCount)) & " where iActiveFlag = 1 "
        Call g_obj_Conn.Execute(sSQL)
        iStatus = iCount + gclPhase_WaitPass - 1
    Else
        iStatus = fMainForm.f_int_CurrentPhase
    End If

    iniFile = prvsOCRDTConvIniFile
    param = ";YY='" & yy & "';iStatus='" & Trim(str(iStatus)) & "'"
    f_bln_ReturnVal = dcvConvert(iniFile, param)
    If f_bln_ReturnVal = 0 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "転送は正常に完了しました。"
    ElseIf f_bln_ReturnVal = -1 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "定義ファイル(" & iniFile & ")の読み込みに失敗しました。"
    ElseIf f_bln_ReturnVal = -2 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "ODBCへの接続に失敗しました。"
    ElseIf f_bln_ReturnVal = -3 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = LoadResString(1727)
    ElseIf f_bln_ReturnVal = -4 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "入力ファイルの読み込みに失敗しました。"
    ElseIf f_bln_ReturnVal = -5 Then
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "データベースのコマンド呼び出しに失敗しました。"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call g_void_CloseChildForm
End Sub

Private Sub udCount_DownClick()

Dim iCount As Integer
Dim sWk As String

    sWk = Trim(txtCount.Text)
    If gf_IntCheck(sWk) Then
        iCount = CInt(sWk) - 1
        If iCount = 0 Then iCount = 1
    Else
        iCount = 1
    End If

    txtCount.Text = Trim(str(iCount))

End Sub

Private Sub udCount_UpClick()

Dim iCount As Integer
Dim sWk As String

    sWk = Trim(txtCount.Text)
    If gf_IntCheck(sWk) Then
        iCount = CInt(sWk) + 1
    Else
        iCount = 1
    End If

    txtCount.Text = Trim(str(iCount))

End Sub
