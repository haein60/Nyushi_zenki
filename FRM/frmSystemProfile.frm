VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSystemProfile 
   BackColor       =   &H80000018&
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14010
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
   ScaleHeight     =   9660
   ScaleWidth      =   14010
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtLoginName 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   855
      Left            =   9360
      MaxLength       =   60
      TabIndex        =   25
      Tag             =   "[vLoginName]"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtSourceName 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   855
      Left            =   3360
      MaxLength       =   60
      TabIndex        =   24
      Tag             =   "[vSourceName]"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CheckBox chkVisibleFlag 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   3360
      TabIndex        =   23
      Tag             =   "[iVisibleFlag]"
      Top             =   1500
      Width           =   2175
   End
   Begin VB.ComboBox cboCurrentPhase 
      Height          =   315
      Left            =   9360
      TabIndex        =   22
      Tag             =   "[iCurrentPhase]"
      Text            =   "Combo1"
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtExamDate 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   3360
      MaxLength       =   60
      TabIndex        =   11
      Tag             =   "[dtExamDate]"
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtActiveFlag 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   9360
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "[iActiveFlag]"
      Top             =   1440
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSearchGrid 
      Height          =   3975
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7011
      _Version        =   393216
      BackColorSel    =   8388608
      BackColorBkg    =   12632256
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtNendo 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   9360
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "[iNendo]"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtSystemProfileId 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   3360
      MaxLength       =   5
      TabIndex        =   1
      Tag             =   "[iSystemProfileId]"
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   9120
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblExamDate 
      BackColor       =   &H80000018&
      Caption         =   "Exam Date"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblActiveFlag 
      BackColor       =   &H80000018&
      Caption         =   "Active Flag"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   18
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblCurrentPhase 
      BackColor       =   &H80000018&
      Caption         =   "Current Phase"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblSourceName 
      BackColor       =   &H80000018&
      Caption         =   "Source Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblVisibleFlag 
      BackColor       =   &H80000018&
      Caption         =   "Visible Flag"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   9120
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblLoginName 
      BackColor       =   &H80000018&
      Caption         =   "Login Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   9120
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   9120
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrorMsg 
      BackColor       =   &H80000018&
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   13455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNendo 
      BackColor       =   &H80000018&
      Caption         =   "Nendo"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblSystemProfileId 
      BackColor       =   &H80000018&
      Caption         =   "System Profile Id"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmSystemProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_bload As Boolean
Public m_TableName As String
Public m_colFieldDetails As New FieldDetails
Public m_bDirty As Boolean
Public m_bChangeOn As Boolean
Public m_lngSearchGridTop As Long
Public m_bMode As String
Public m_lngCurrentRow As Long

Private Sub Form_Activate()
On Error GoTo ErrHandler
    Dim lngRow As Integer
    If mbload = False Then
        m_lngSearchGridTop = 3600
        'initialize the search grid
        InitializeSearchGrid
        'populate the grid with empty row and column headings
        SearchRecords True
        fMainForm.mnuToolsSave.Enabled = False
        fMainForm.mnuToolsCancel.Enabled = False
    End If
    m_bload = True
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    ' set the table name
    m_TableName = "tbSTESystemProfile"
    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained
    With m_colFieldDetails
        .Add "[iSystemProfileId]", "[SystemProfile Id]", 1, True, True, "", "INT", 1500, "", "[iSystemProfileId]"
        .Add "[iNendo]", "[Nendo]", 2, True, False, "", "INT", 1500, "", "[iNendo]"
        .Add "[dtExamDate]", "[Exam Date]", 3, True, False, "", "DATETIME", 1500, "", "[dtExamDate]"
        .Add "[iCurrentPhase]", "[Current Phase]", 4, True, False, "", "INT", 1500, "", "[iCurrentPhase]"
        .Add "[iVisibleFlag]", "[Visible Flag]", 5, True, False, "", "INT", 1500, "", "[iVisibleFlag]"
        .Add "[iActiveFlag]", "[Active Flag]", 6, True, False, "", "INT", 1500, "", "[iActiveFlag]"
        .Add "[vSourceName]", "[Source Name]", 7, True, False, "", "VARCHAR", 1500, "", "[vSourceName]"
        .Add "[vLoginName]", "[Login Name]", 8, True, False, "", "VARCHAR", 1500, "", "[vLoginName]"
    End With
    
    With cboCurrentPhase
    .AddItem "ApplyPhase", 0
    .List(0) = 0
    .AddItem "1stExam", 1
    .List(1) = 1
    .AddItem "2ndExam", 2
    .List(2) = 2
    .AddItem "enter", 3
    .List (3)
    End With
    
    EstablishConnection
    m_bload = False
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub hfgSearchGrid_DblClick()
    'check  if existing data is dirty if yes then prompt for saving
    ' else
    'populate the new row in the text boxes
    AssignValues True
End Sub

Private Sub hfgSearchGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        hfgSearchGrid_DblClick
    End If
End Sub

Private Sub txtAddress_Change()
    SetChange
End Sub

Private Sub txtCity_Change()
    SetChange
End Sub

Private Sub txtContactName_Change()
    SetChange
End Sub

Private Sub txtContactTitle_Change()
    SetChange
End Sub

Private Sub txtCountry_Change()
    SetChange
End Sub

Private Sub txtCustomerId_Change()
    SetChange
End Sub

Private Sub SetChange()
    If m_bChangeOn = False Then m_bDirty = True
    If m_bDirty = True And fMainForm.mnuToolsSave.Enabled = False Then
        fMainForm.mnuToolsSave.Enabled = True
        fMainForm.mnuToolsCancel.Enabled = True
    End If
End Sub

Private Sub txtFax_Change()
    SetChange
End Sub

Private Sub txtlCompanyName_Change()
    SetChange
End Sub

Private Sub txtPhone_Change()
    SetChange
End Sub

Private Sub txtPostalCode_Change()
    SetChange
End Sub

Private Sub txtRegion_Change()
    SetChange
End Sub

Public Function ExtraValidation() As Boolean
    On Error GoTo ErrorHandler
    ' if goes through then ExtraValidation is true if fails then false
    
    ExtraValidation = True
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description
    
End Function

