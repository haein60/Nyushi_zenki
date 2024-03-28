VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmInterviewGroupProfile 
   Caption         =   "frmInterviewGroupProfile : 所属プロファイル"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmInterviewGroupProfile.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   9885
   WindowState     =   2  '最大化
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgsearchgrid 
      Height          =   3855
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   6800
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtInterviewGroupProfileId 
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
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   1
      Tag             =   "[iInterviewGroupProfileId]"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtInterviewGroupName 
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
      Left            =   3480
      MaxLength       =   40
      TabIndex        =   0
      Tag             =   "[vInterviewGroupName]"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lblInterviewerProfileId 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "2468"
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
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblInterviewerName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "2469"
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
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lblErrorMsg 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   6255
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  '透明
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  '透明
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmInterviewGroupProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmInterviewGroupProfile
'Author         :   Mahesh Deshpande
'Created On     :   09/05/2002 (9th May 2002)
'Description    :   This form makes a provision for master maintenance of tbSTRInterviewGroupProfile Table.
'Reference      :   Functional Specs Of MasterMaintenance Ver 1.0
'***************************************************************************************************
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
    Dim Index As Integer

    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       If Index = 1 Or Index = 6 Then
            fMainForm.Toolbar1.Buttons(Index).Enabled = True
        Else
            fMainForm.Toolbar1.Buttons(Index).Enabled = False
        End If
    Next

    If m_bload = False Then
        m_lngSearchGridTop = 1700
        'initialize the search grid
        InitializeSearchGrid
        'populate the grid with empty row and column headings
        SearchRecords True
        fMainForm.mnuToolsSave.Enabled = False
        fMainForm.mnuToolsCancel.Enabled = False
        
        fMainForm.Toolbar1.Buttons("Save").Enabled = False
        fMainForm.Toolbar1.Buttons("Cancel").Enabled = False
'        cmdSave.Enabled = False
'        cmdCancel.Enabled = False
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
    Me.Caption = "frmInterviewGroupProfile : 所属プロファイル" ''''LoadResString(2466)

    Call g_void_SetFontProperties(Me)     ' set the font properties

    ' set the table name
    m_TableName = "tbSTEInterviewGroupProfile"

    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained
    With m_colFieldDetails
        .Add "[iInterviewGroupProfileId]", "[" & LoadResString(2468) & "]", 1, True, True, "", "INT", 2800, "", "[iInterviewGroupProfileId]"
        .Add "[vInterviewGroupName]", "[" & LoadResString(2469) & "]", 2, True, False, "", "STRING", 2600, "", "[vInterviewGroupName]"
    End With
    
    'EstablishConnection
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
'        cmdSave.Enabled = True
'        cmdCancel.Enabled = True
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

Private Sub txtInterviewGroupName_Change()

    SetChange

End Sub

Private Sub txtInterviewGroupProfileId_Change()

'    SetChange

End Sub

Private Sub txtInterviewGroupProfileId_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub
