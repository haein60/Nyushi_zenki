VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSubjectProfile 
   Caption         =   "Form1"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
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
   Picture         =   "frmSubjectProfile.frx":0000
   ScaleHeight     =   9645
   ScaleWidth      =   10545
   WindowState     =   2  'ç≈ëÂâª
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSearchGrid 
      Height          =   3855
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6800
      _Version        =   393216
      FocusRect       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cboExamType 
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
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3720
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   5
      Tag             =   "[iExamType]"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtSubjectName 
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
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3720
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "[vSubjectName]"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtSubjectProfileId 
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
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   1
      Tag             =   "[iSubjectProfileId]"
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblExamType 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1404"
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
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   7
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
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   6015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSubjectName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1403"
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
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label lblSubjectProfileId 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "1402"
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
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
End
Attribute VB_Name = "frmSubjectProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmSubjectProfile
'Author         :   Vishal Kamath
'Created On     :   10/8/01
'Description    :   This form makes a provision for master maintenance of tbSTRSubjectProfile Table.
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
Public m_ComboDetails As New ComboCollection
Public f_int_PrevRow As Long

Private Sub cboExamType_Click()
    SetChange
End Sub

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
    Me.Caption = LoadResString(1401)
    Call g_void_SetFontProperties(Me)     ' set the font properties
    ' set the table name
    m_TableName = "tbSTESubjectProfile"
    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained
    With m_colFieldDetails
        .Add "[iSubjectProfileId]", "[" & LoadResString(1402) & "]", 1, True, True, "", "INTEGER", 2000, "", "[iSubjectProfileId]"
        .Add "[vSubjectName]", "[" & LoadResString(1403) & "]", 2, True, False, "", "STRING", 3000, "", "[vSubjectName]"
        .Add "[iExamType]", "[" & LoadResString(1404) & "]", 3, True, False, "", "COMBO", 2000, "", "[iExamType]"
    End With
    
   Call f_void_AddExamType ' Added by team
   ' EstablishConnection
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

Private Sub txtExamType_Change()
    SetChange
End Sub

Private Sub txtSubjectName_Change()
    SetChange
End Sub

Private Sub f_void_AddExamType()
    With cboExamType
        .AddItem "", 0
        .AddItem LoadResString(1005), 1
        .AddItem LoadResString(1008), 2
        .AddItem LoadResString(2434), 3
        .AddItem LoadResString(2431), 4
        .AddItem LoadResString(2432), 5
        .AddItem LoadResString(2433), 6
        
        m_ComboDetails.Add 0, LoadResString(1005), "[iExamType]"
        m_ComboDetails.Add 1, LoadResString(1008), "[iExamType]"
        m_ComboDetails.Add 2, LoadResString(2434), "[iExamType]"
        m_ComboDetails.Add 3, LoadResString(2431), "[iExamType]"
        m_ComboDetails.Add 4, LoadResString(2432), "[iExamType]"
        m_ComboDetails.Add 5, LoadResString(2433), "[iExamType]"
    End With
End Sub

Private Sub txtSubjectProfileId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 46 And KeyCode <> 8 Then
        KeyCode = 0
    End If
End Sub
