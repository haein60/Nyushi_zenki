VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInterviewRoomProfile 
   Caption         =   "InterviewRoomProfile"
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
   Begin VB.ComboBox cboInterviewerProfileId 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "[iInterviewerProfileId]"
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox cboRoomProfileId 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "[iRoomProfileId]"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtInterviewRoomProfileId 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4080
      MaxLength       =   60
      TabIndex        =   1
      Tag             =   "[iInterviewRoomProfileId]"
      Top             =   360
      Width           =   2175
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSearchGrid 
      Height          =   6615
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11668
      _Version        =   393216
      BackColorSel    =   8388608
      BackColorBkg    =   12632256
      ScrollBars      =   2
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
   Begin VB.Label lblInterviewerProfileId 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1603"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label lblRoomProfileId 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1604"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblInterviewRoomProfileId 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1602"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
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
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblErrorMsg 
      BackColor       =   &H00F4DBC4&
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   9120
      Visible         =   0   'False
      Width           =   7575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInterviewRoomProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*************************************************************************************************
'Form Name      :   frmInterviewerRoomProfile
'Author         :   Vishal Kamath
'Created On     :
'Description    :   This form makes a provision for master maintenance of tbSTRInterviewerRoomProfile Table.
'Reference      :   Functional Specs Of MasterMaintenance Ver 1.0
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


Private Sub cboInterviewerProfileId_Click()
    SetChange
End Sub

Private Sub cboRoomProfileId_Click()
    SetChange
End Sub

Private Sub Form_Activate()
On Error GoTo ErrHandler
    Dim lngRow As Integer
    Dim index
    For index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(index).Enabled = True
    Next
    
    If m_bload = False Then
        m_lngSearchGridTop = 3600
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
    Me.Caption = LoadResString(1601)
    ' set the table name
    m_TableName = "tbSTEInterviewRoomProfile"
    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained
    With m_colFieldDetails
        .Add "[iInterviewRoomProfileId]", "[" & LoadResString(1602) & "]", 1, True, True, "", "INTEGER", 3000, "", "[iInterviewRoomProfileId]"
        .Add "[iInterviewerProfileId]", "[" & LoadResString(1603) & "]", 2, True, False, "", "COMBO", 2700, "", "[iInterviewerProfileId]"
        .Add "[iRoomProfileId]", "[" & LoadResString(1604) & "]", 3, True, False, "", "COMBO", 2200, "", "[iRoomProfileId]"
    End With
           
    Call f_void_AddInterviewerProfileID
    Call f_void_AddRoomProfileID
    
    
    m_bload = False
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_colFieldDetails = Nothing
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

Private Sub SetChange()
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

Private Sub tblcboInterviewerProfileId_Click()
    SetChange
End Sub

Private Sub tblcboRoomProfileId_Click()
    SetChange
End Sub


Private Sub f_void_AddInterviewerProfileID()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_sql = "Select iInterviewerProfileId, vInterviewerName from tbSTEInterviewerProfile"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
         
         cboInterviewerProfileId.AddItem "", 0
         
         If Not l_obj_Rst.EOF Then
            l_obj_Rst.MoveFirst
         Else
            Exit Sub
         End If
         
         With m_ComboDetails
            While Not l_obj_Rst.EOF
               cboInterviewerProfileId.AddItem l_obj_Rst.Fields("vInterviewerName")
               m_ComboDetails.Add l_obj_Rst.Fields("iInterviewerProfileId"), l_obj_Rst.Fields("vInterviewerName"), cboInterviewerProfileId.Tag
               l_obj_Rst.MoveNext
            Wend
         End With
         
     l_obj_Rst.Close
     Set l_obj_Rst = Nothing
     
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
    Resume
End Sub

Private Sub f_void_AddRoomProfileID()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_sql = "Select iRoomProfileId, vRoomName from tbSTERoomProfile"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
    
        cboRoomProfileId.AddItem "", 0
    
         If Not l_obj_Rst.EOF Then
            l_obj_Rst.MoveFirst
         Else
            Exit Sub
         End If
         
         With m_ComboDetails
            While Not l_obj_Rst.EOF
               cboRoomProfileId.AddItem l_obj_Rst.Fields("vRoomName")
               m_ComboDetails.Add l_obj_Rst.Fields("iRoomProfileId"), l_obj_Rst.Fields("vRoomName"), cboRoomProfileId.Tag
               l_obj_Rst.MoveNext
            Wend
         End With
         
     l_obj_Rst.Close
     Set l_obj_Rst = Nothing
     
     Exit Sub
ErrorHandler:
    MsgBox Err.Description
    Resume
End Sub

Private Sub txtInterviewRoomProfileId_Change()
'    SetChange
End Sub


Private Sub txtInterviewRoomProfileId_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
