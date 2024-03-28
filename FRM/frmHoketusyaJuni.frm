VERSION 5.00
Begin VB.Form frmHoketusyaJuni 
   AutoRedraw      =   -1  'True
   Caption         =   "frmHoketusyaJuni : ï‚åáé“èáà "
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmHoketusyaJuni.frx":0000
   ScaleHeight     =   10305
   ScaleWidth      =   13425
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.CommandButton cmdOK 
      Caption         =   "ämíË"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   645
      TabIndex        =   8
      Top             =   8650
      Width           =   1300
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ï¬Ç∂ÇÈ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4710
      TabIndex        =   7
      Top             =   8650
      Width           =   1300
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "èoóÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2610
      TabIndex        =   6
      Top             =   8650
      Width           =   1300
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Å™"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6210
      TabIndex        =   5
      Top             =   3870
      Width           =   480
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Å´"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6195
      TabIndex        =   4
      Top             =   5265
      Width           =   480
   End
   Begin VB.ListBox lstKuriage 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6585
      Left            =   660
      TabIndex        =   3
      Top             =   1695
      Width           =   5325
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "ï\é¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4680
      TabIndex        =   2
      Top             =   1125
      Width           =   1300
   End
   Begin VB.TextBox txtNendo 
      Alignment       =   2  'íÜâõëµÇ¶
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      IMEMode         =   3  'µÃå≈íË
      Left            =   1875
      MaxLength       =   4
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "[iNendo]"
      Top             =   795
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblNendo 
      BackStyle       =   0  'ìßñæ
      Caption         =   "YYYY"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1905
      TabIndex        =   19
      Top             =   1245
      Width           =   1170
   End
   Begin VB.Line Line4 
      X1              =   7920
      X2              =   11050
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Line Line3 
      X1              =   11040
      X2              =   11040
      Y1              =   1710
      Y2              =   4810
   End
   Begin VB.Line Line2 
      X1              =   7920
      X2              =   11050
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      X1              =   7920
      X2              =   7920
      Y1              =   1710
      Y2              =   4810
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'ìßñæ
      Caption         =   "çáåvìæì_ÇÃçÇÇ¢èáÇ…ï\é¶ÇµÇ‹Ç∑ÅB"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8145
      TabIndex        =   18
      Top             =   2040
      Width           =   3030
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'ìßñæ
      Caption         =   $"frmHoketusyaJuni.frx":3AD3
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8280
      TabIndex        =   17
      Top             =   4455
      Width           =   1275
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'ìßñæ
      Caption         =   "áFñ ê⁄áT"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8295
      TabIndex        =   16
      Top             =   4155
      Width           =   1470
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'ìßñæ
      Caption         =   "áEè¨ò_ï∂"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8295
      TabIndex        =   15
      Top             =   3855
      Width           =   1590
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'ìßñæ
      Caption         =   "áDê∂ï®"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8295
      TabIndex        =   14
      Top             =   3555
      Width           =   1590
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'ìßñæ
      Caption         =   "áCâªäw"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8295
      TabIndex        =   13
      Top             =   3255
      Width           =   1590
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'ìßñæ
      Caption         =   "áBï®óù"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8295
      TabIndex        =   12
      Top             =   2955
      Width           =   1590
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'ìßñæ
      Caption         =   "áAêîäw"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8295
      TabIndex        =   11
      Top             =   2655
      Width           =   1590
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'ìßñæ
      Caption         =   "á@âpåÍ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8295
      TabIndex        =   10
      Top             =   2355
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'ìßñæ
      Caption         =   "ï\é¶èáÇÕÅAà»â∫ÇÃâ»ñ⁄ÇÃ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8145
      TabIndex        =   9
      Top             =   1830
      Width           =   2535
   End
   Begin VB.Label lblTit 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "èàóùîNìx"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   735
      TabIndex        =   1
      Tag             =   "1804"
      Top             =   1245
      Width           =   1080
   End
End
Attribute VB_Name = "frmHoketusyaJuni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public g_obj_Conn               As ADODB.Connection   'connection object
Public g_void_OpenConnection    As Boolean


'*******************************************************************************
'* 3.10 ï‚åáé“èáà (sub-systemÇÊÇËìùçá)                                         *
'*-----------------------------------------------------------------------------*
'* Form Load                                                                   *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Dim sConn        As String
    Dim sPwd         As String
    Dim sUser        As String
    Dim sDatabase    As String
    Dim sMachine     As String


''''txtNendo.Text = Year(Date)

    lblNendo.Caption = g_int_CurrentNendo & "îN"

    cmdUp.Enabled = False
    cmdDown.Enabled = False

    g_void_OpenConnection = False
   
   
    sPwd = GetSetting("Nyushi", "Settings", "DatabasePassword", "")
    sUser = GetSetting("Nyushi", "Settings", "DatabaseUser", "")
    sDatabase = GetSetting("Nyushi", "Settings", "DatabaseName", "")
    sMachine = GetSetting("Nyushi", "Settings", "MachineName", "")

    If Trim(sUser) = "" Or Trim(sDatabase) = "" Or Trim(sMachine) = "" Then
        MsgBox "ÉfÅ[É^ÉxÅ[ÉXÇÃê›íËÇ…åÎÇËÇ™Ç†ÇËÇ‹Ç∑ÅB", vbInformation, "ï‚åáé“èáà "
        Exit Sub
    End If

    sConn = ";DSN=" & sMachine & ";UID=" & sUser & ";PWD=" & sPwd & ";Database=" & sDatabase

    Set g_obj_Conn = New ADODB.Connection
    g_obj_Conn.CursorLocation = adUseClient
    g_obj_Conn.Open sConn ''''Database Open
  
    If Err.Number <> 0 Then
        g_void_OpenConnection = False
    Else
        g_void_OpenConnection = True
    End If
    
    Exit Sub

ErrorHandler:
        MsgBox Err.Description, vbInformation, "ï‚åáé“èáà "

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If g_void_OpenConnection = True Then
        g_void_OpenConnection = False
        g_obj_Conn.Close
        Set g_obj_Conn = Nothing
    End If

End Sub

'*******************************************************************************
'* Åyï\é¶ÅzÉ{É^Éìèàóù                                                          *
'*******************************************************************************
Private Sub cmdShow_Click()
   
    On Error GoTo ErrorHandler

'   Dim blnOpenDB              As Boolean
    Dim strNendo               As String
    Dim l_obj_Rst              As ADODB.Recordset    'recordset object
    Dim l_str_Sql              As String             'The SQL string
    Dim l_str_DisplayString    As String             'to form the display string in the list box

    
    'check nendo
    cmdUp.Enabled = False
    cmdDown.Enabled = False
        
''''    strNendo = txtNendo.Text
    strNendo = g_int_CurrentNendo

''''    If strNendo = "" Then
''''        MsgBox "îNìxÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", vbInformation, "ï‚åáé“èáà "
''''        Exit Sub
''''    End If
''''
''''    strNendo = Trim(strNendo)
''''    If strNendo = "" Then
''''        MsgBox "îNìxÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", vbInformation, "ï‚åáé“èáà "
''''        Exit Sub
''''    End If
''''
''''    If strNendo >= "2101" Or strNendo < "2010" Then
''''        MsgBox "îNìxì¸óÕÇ…åÎÇËÇ™Ç†ÇËÇ‹Ç∑ÅB(2010Å`2100îNÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢)", vbInformation, "ï‚åáé“èáà "
''''        Exit Sub
''''    End If
 
    
    
    Me.lstKuriage.Clear


    cmdShow.Enabled = False    '2021.11.17 add jhi

    'getData

'     l_str_Sql = "SELECT iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName,iSex FROM tbSTEExamineeProfile WHERE" & _
'    " iNendo = " & strNendo & _
'    " AND iAbsentFlag = 0"
'    l_str_Sql = l_str_Sql & " AND iExamineeStatus = 3"

    l_str_Sql = "Exec uspSTEGetExamineeOrder " & strNendo
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If l_obj_Rst.EOF Then
        Set l_obj_Rst = Nothing
        MsgBox "éwíËîNìxÇ…äYìñÇ∑ÇÈÉfÅ[É^ÇÕÇ†ÇËÇ‹ÇπÇÒÅB", vbInformation, "ï‚åáé“èáà " '2021.11.17 add jhi
        cmdShow.Enabled = True
        Exit Sub
    End If


    Do While Not l_obj_Rst.EOF

        l_str_DisplayString = g_str_LPad(l_obj_Rst.Fields("iJukenNumber").Value, Len(l_obj_Rst.Fields("iJukenNumber").Value)) & _
            " - " & l_obj_Rst.Fields("vExamineeName").Value

'        If l_obj_Rst.Fields("iSex").Value = 0 Then
'            l_str_DisplayString = l_str_DisplayString & " - (*)"
'        End If
        
        lstKuriage.AddItem l_str_DisplayString
        l_obj_Rst.MoveNext

    Loop

    l_obj_Rst.Close
    Set l_obj_Rst = Nothing


    cmdShow.Enabled = True    '2021.11.17 add jhi
  
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ï‚åáé“èáà "

End Sub

Private Sub cmdDown_Click()

    On Error GoTo ErrorHandler

    Dim l_int_Count As Integer


    If lstKuriage.ListCount < 1 Then
       Exit Sub
    End If

    If lstKuriage.SelCount < 0 Then
        Exit Sub
    End If

    'lstKuriage
    For l_int_Count = 0 To lstKuriage.ListCount - 1
        If lstKuriage.Selected(l_int_Count) Then
            lstKuriage.AddItem lstKuriage.List(l_int_Count), l_int_Count + 2
            lstKuriage.RemoveItem l_int_Count
            lstKuriage.Selected(l_int_Count + 1) = True
            Exit Sub
        End If
    Next
    
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ï‚åáé“èáà "

End Sub

Private Sub cmdExcel_Click()

    On Error GoTo ErrorHandler

    Dim fso                   As Object
    Dim objText               As Object
    Dim strFile               As String
    Dim blnOpenFile           As Boolean
    Dim l_str_JukenNo         As String
    Dim l_str_ExamineeName    As String
    Dim l_int_Count           As Integer
    Dim strLine               As String


    If lstKuriage.ListCount < 1 Then
        Exit Sub
    End If


    blnOpenFile = False

    'FSOÉIÉuÉWÉFÉNÉbÉgÇèâä˙âª
    strFile = App.Path & "\Report\ï‚åáé“àÍóó" & Format(Now(), "YYYYMMDDHHmmSS") & ".csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.CreateTextFile(strFile, True, False)

    blnOpenFile = True

    l_str_JukenNo = ""
    l_str_ExamineeName = ""


    'ÉtÉ@ÉCÉãÇèoóÕ
    For l_int_Count = 0 To lstKuriage.ListCount - 1
        l_str_JukenNo = Left(lstKuriage.List(l_int_Count), 4)

        l_str_ExamineeName = Mid(lstKuriage.List(l_int_Count), 7)
        l_str_ExamineeName = Trim(l_str_ExamineeName)
        strLine = l_int_Count + 1 & "," & l_str_JukenNo & "," & l_str_ExamineeName
'       strLine = lstKuriage.List(l_int_Count)
        objText.WriteLine (strLine)
    Next

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If

    ShellExecute Me.hwnd, "open", strFile, vbNullString, vbNullString, 1

    Exit Sub

ErrorHandler:
    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If
    MsgBox Err.Description, vbInformation, "ï‚åáé“èáà "

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorHandler

    Dim strNendo              As String
    Dim l_str_Sql             As String             ' The SQL string
    
    Dim l_int_TempJuken       As Integer            ' to store the juken number
    Dim l_str_JukenNo         As String             ' to store all the lstThisTimeSelected juken numbers as a string
    Dim l_str_ExamineeName    As String
    Dim blnTrans              As Boolean
    Dim l_int_Count           As Integer
    
    
    'check nendo
    cmdUp.Enabled = False
    cmdDown.Enabled = False
        
    strNendo = txtNendo.Text
    If strNendo = "" Then
        MsgBox "îNìxÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", vbInformation, "ï‚åáé“èáà "
        Exit Sub
    End If

    strNendo = Trim(strNendo)
    If strNendo = "" Then
        MsgBox "îNìxÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", vbInformation, "ï‚åáé“èáà "
        Exit Sub
    End If

    If strNendo >= "9999" And strNendo < "2010" Then
        MsgBox "îNìxì¸óÕóìÇ…åÎÇËÇ™Ç†ÇËÇ‹Ç∑ÅB", vbInformation, "ï‚åáé“èáà "
        Exit Sub
    End If

    'getData

    
    blnTrans = False

'   l_str_Sql = "Exec uspSTESetExamineeOrder " & strNendo

    g_obj_Conn.BeginTrans

    l_str_Sql = "DELETE FROM tbSTEExamineeOrder WHERE  iNendo=" & strNendo
    g_obj_Conn.Execute (l_str_Sql)

    blnTrans = True

    For l_int_Count = 0 To lstKuriage.ListCount - 1
        l_int_TempJuken = Left(lstKuriage.List(l_int_Count), 4)
        l_str_JukenNo = l_int_TempJuken
        
        l_str_ExamineeName = Mid(lstKuriage.List(l_int_Count), 7)
        l_str_ExamineeName = Trim(l_str_ExamineeName)
        l_str_Sql = "INSERT INTO tbSTEExamineeOrder(iJukenNumber,iNendo,vExamineeName)"
        l_str_Sql = l_str_Sql & "VALUES(" & l_str_JukenNo & "," & strNendo
        l_str_Sql = l_str_Sql & ",'" & l_str_ExamineeName & "'"
        l_str_Sql = l_str_Sql & ")"
        g_obj_Conn.Execute (l_str_Sql)
    Next

    g_obj_Conn.CommitTrans
    blnTrans = False
    
    MsgBox "ï‚åáé“èáî‘ÇçXêVÇµÇ‹ÇµÇΩÅB", vbInformation, "ï‚åáé“èáà "

    Exit Sub

ErrorHandler:
    If blnTrans = True Then
        blnTrans = False
        g_obj_Conn.RollbackTrans
    End If

    MsgBox Err.Description, vbInformation, "ï‚åáé“èáà "

End Sub


Public Function g_str_LPad(ByVal str As String, ByVal iLen As Integer) As String

    '-------------------------------------------------------------
    'Left pads a string with 0 up to iLen.
    '-------------------------------------------------------------
    Select Case iLen
    Case 1
        g_str_LPad = "000" & str
    Case 2
        g_str_LPad = "00" & str
    Case 3
        g_str_LPad = "0" & str
    Case 4
        g_str_LPad = str
    End Select

End Function

Private Sub cmdUp_Click()

    On Error GoTo ErrorHandler

    Dim l_int_Count As Integer



    If lstKuriage.ListCount < 1 Then
        Exit Sub
    End If

    If lstKuriage.SelCount < 0 Then
        Exit Sub
    End If


    'lstKuriage

    For l_int_Count = 0 To lstKuriage.ListCount - 1
        If lstKuriage.Selected(l_int_Count) Then
            lstKuriage.AddItem lstKuriage.List(l_int_Count), l_int_Count - 1
            lstKuriage.RemoveItem l_int_Count + 1
            lstKuriage.Selected(l_int_Count - 1) = True
            Exit Sub
        End If
    Next
    
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ï‚åáé“èáà "

End Sub

Private Sub cmdClose_Click()

    On Error GoTo ErrorHandler

    If g_void_OpenConnection = True Then
        g_void_OpenConnection = False
        g_obj_Conn.Close
        Set g_obj_Conn = Nothing
    End If

''''End          '2021.11.17 del jhi
    Unload Me    '2021.11.17 add jhi
    Exit Sub     '2021.11.17 add jhi

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ï‚åáé“èáà "

End Sub


Private Sub lstKuriage_Click()

    On Error GoTo ErrorHandler

    If lstKuriage.ListCount < 1 Then
        Exit Sub
    End If
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    If lstKuriage.Selected(0) Then
        cmdUp.Enabled = False
'        Exit Sub
    End If
    If lstKuriage.Selected(lstKuriage.ListCount - 1) Then
        cmdDown.Enabled = False
'        Exit Sub
    End If
    Exit Sub
ErrorHandler:
        MsgBox Err.Description, vbInformation, "ï‚åáé“èáà "
End Sub


