VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmZipCode 
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12285
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmZipCode.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   12285
   WindowState     =   2  'ç≈ëÂâª
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgsearchgrid 
      Height          =   3855
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6800
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtPrefectureName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      Left            =   3120
      MaxLength       =   60
      TabIndex        =   9
      Tag             =   "[vPrefectureName]"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtAddress2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      Left            =   8760
      MaxLength       =   50
      TabIndex        =   11
      Tag             =   "[vAddress2]"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtCityName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      Left            =   8760
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "[vCityName]"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtAddress1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      Left            =   8760
      MaxLength       =   50
      TabIndex        =   7
      Tag             =   "[vAddress1]"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtZipCodeName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      Left            =   3120
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "[vZipCodeName]"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtZipCodeId 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      Left            =   3105
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   1
      Tag             =   "[iZipCodeId]"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblPrefectureName 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1304"
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
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblAddress2 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1307"
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
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblCityName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1305"
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
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblAddress1 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1306"
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
      Height          =   255
      Left            =   6360
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   8400
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   8400
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   8400
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrorMsg 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   10650
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblZipCodeName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1303"
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
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblZipCodeId 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1302"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "frmZipCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmZipCode
'Author         :   Vishal Kamath
'Created On     :   10/8/01
'Description    :   This form makes a provision for master maintenance of tbSTRZipCode Table.
'Reference      :   Functional Specs Of MasterMaintenance Ver 1.0
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History   -   04/04/2002  -   Dileep Cherian
'On activation of the form, only the "new" and "query" toolbar icons should be enabled
'**************************************************************************************************
Dim m_bload As Boolean
Public m_TableName As String
Public m_colFieldDetails As New FieldDetails
Public m_bDirty As Boolean
Public m_bChangeOn As Boolean
Public m_lngSearchGridTop As Long
Public m_bMode As String
Public m_lngCurrentRow As Long
Public f_int_PrevRow As Long

Private Sub Form_Activate()
On Error GoTo ErrHandler
    Dim lngRow As Integer
    Dim index As Integer
    For index = 1 To fMainForm.Toolbar1.Buttons.Count
       If index = 1 Or index = 6 Then
            fMainForm.Toolbar1.Buttons(index).Enabled = True
        Else
            fMainForm.Toolbar1.Buttons(index).Enabled = False
        End If
    Next
    If m_bload = False Then
        m_lngSearchGridTop = 2000
        'initialize the search grid
        InitializeSearchGrid
        'populate the grid with empty row and column headings
        SearchRecords True
        fMainForm.mnuToolsSave.Enabled = False
        fMainForm.mnuToolsCancel.Enabled = False
        
        fMainForm.Toolbar1.Buttons("Save").Enabled = False
        fMainForm.Toolbar1.Buttons("Cancel").Enabled = False
    End If
    m_bload = True
    fMainForm.mnuTools.Enabled = True
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    LoadResStrings Me
    Me.Caption = LoadResString(1301)
    Call g_void_SetFontProperties(Me)     ' set the font properties
    ' set the table name
    m_TableName = "tbSTEZipCodeMaster"
    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained
    With m_colFieldDetails
        .Add "[iZipCodeId]", "[" & LoadResString(1302) & "]", 1, True, True, "", "INTEGER", 1500, "", "[iZipCodeId]"
        .Add "[vZipCodeName]", "[" & LoadResString(1303) & "]", 2, True, False, "", "STRING", 1800, "", "[vZipCodeName]"
        .Add "[vPrefectureName]", "[" & LoadResString(1304) & "]", 3, True, False, "", "STRING", 2000, "", "[vPrefectureName]"
        .Add "[vCityName]", "[" & LoadResString(1305) & "]", 4, True, False, "", "STRING", 1500, "", "[vCityName]"
        .Add "[vAddress1]", "[" & LoadResString(1306) & "]", 5, True, False, "", "STRING", 1500, "", "[vAddress1]"
        .Add "[vAddress2]", "[" & LoadResString(1307) & "]", 6, False, False, "", "STRING", 1500, "", "[vAddress2]"
    End With
    m_bload = False
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_colFieldDetails = Nothing
    m_bDirty = False
    Call g_void_CloseChildForm
End Sub

Private Sub hfgSearchGrid_DblClick()
    'check  if existing data is dirty if yes then prompt for saving
    ' else
    'populate the new row in the text boxes
    AssignValues True
    Call g_void_HighlightRow(hfgsearchgrid.Row, f_int_PrevRow)
    f_int_PrevRow = hfgsearchgrid.Row
End Sub

Private Sub hfgSearchGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        hfgSearchGrid_DblClick
    End If
End Sub

Private Sub SetChange()
    lblErrorMsg.Caption = ""
    lblErrorMsg.Visible = False
    If m_bMode = "SEARCH" Then Exit Sub
    If m_bChangeOn = False Then m_bDirty = True
    If m_bDirty = True And fMainForm.mnuToolsSave.Enabled = False Then
        fMainForm.mnuToolsSave.Enabled = True
        fMainForm.mnuToolsCancel.Enabled = True
        
        fMainForm.Toolbar1.Buttons("Save").Enabled = True
        fMainForm.Toolbar1.Buttons("Cancel").Enabled = True
    End If
End Sub

Public Function ExtraValidation() As Boolean
    On Error GoTo ErrorHandler
    ' if goes through then ExtraValidation is true if fails then false
    ExtraValidation = True
    Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Private Sub txtAddress1_Change()
    SetChange
End Sub

Private Sub txtAddress2_Change()
    SetChange
End Sub

Private Sub txtCityName_Change()
    SetChange
End Sub

Private Sub txtPrefectureName_Change()
    SetChange
End Sub

Private Sub txtZipCodeId_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtZipCodeName_Change()
    SetChange
End Sub
