VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSubjectQuestionProfile 
   Caption         =   "2421"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   14700
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtQuestion 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Tag             =   "[vQuestionName]"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtQuestionNumber 
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Tag             =   "[iQuestionNo]"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtSubjectQuestionId 
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Tag             =   "[iSubjectQuestionId]"
      Top             =   240
      Width           =   2415
   End
   Begin VB.ComboBox cboInterViewerProfileId 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "[iInterviewerProfileID]"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.ComboBox cboSubjectProfileID 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "[iSubjectProfileId]"
      Top             =   840
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSearchGrid 
      Height          =   4095
      Left            =   240
      TabIndex        =   15
      Top             =   3720
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   7223
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
   Begin VB.Label lblErrorMsg 
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   8880
      Visible         =   0   'False
      Width           =   7575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblErrIndicator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Index           =   4
      Left            =   3120
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "2440"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "2439"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "2438"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "2437"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblCode 
      Caption         =   "2436"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmSubjectQuestionProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*************************************************************************************************
'Form Name      :   frmSubjectQuestionProfile
'Author         :   Sankha Mitra
'Created On     :
'Description    :   This form makes a provision for master maintenance of tbSTESubjectQuestionProfile Table.
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
Public f_int_ExamType As Integer

Private Sub cboInterViewerProfileId_Click()
SetChange
End Sub

'Private Sub cboSubjectProfileID_Change()
'    cboSubjectProfileID_Click
'End Sub
'
Private Sub cboSubjectProfileID_Click()
  Dim l_obj_subject As New ADODB.Recordset
  Dim l_str_sql As String
  'Dim l_int_exam As Integer
    SetChange
On Error GoTo errorhandler

  If cboSubjectProfileID.Text <> "" Then
      l_str_sql = "SELECT iExamType FROM tbSTESubjectProfile WHERE vSubjectName = " & "'" & cboSubjectProfileID.Text & "'"
      l_obj_subject.Open l_str_sql, g_obj_Conn, adOpenStatic, adLockReadOnly
    
      f_int_ExamType = l_obj_subject.Fields("iExamType").Value
      If f_int_ExamType = 0 Or f_int_ExamType = 1 Then
            If cboInterViewerProfileId.Text <> "" Then
                cboInterViewerProfileId.ListIndex = 0
                fMainForm.mnuToolsSave.Enabled = False
                fMainForm.Toolbar1.Buttons("Save").Enabled = False
'            Else
'                txtQuestion.Text = ""
'                txtQuestionNumber = ""
            End If
                        
            cboInterViewerProfileId.Enabled = False
            txtQuestionNumber.Enabled = True
            txtQuestion.Enabled = True
      Else
            cboInterViewerProfileId.Enabled = True
            txtQuestionNumber.Enabled = False
            txtQuestion.Enabled = False
      End If
  End If
  Set l_obj_subject = Nothing
  Exit Sub
errorhandler:
        MsgBox Err.Description
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
    Me.Caption = LoadResString(2435)
    ' set the table name
    m_TableName = "tbSTESubjectQuestionProfile"
    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained
    With m_colFieldDetails
        .Add "[iSubjectQuestionId]", "[" & LoadResString(2436) & "]", 1, True, True, "", "INTEGER", 2500, "", "[iSubjectQuestionId]"
        .Add "[iSubjectProfileId]", "[" & LoadResString(2437) & "]", 2, True, False, "", "COMBO", 2500, "", "[iSubjectProfileId]"
        .Add "[iInterviewerProfileID]", "[" & LoadResString(2438) & "]", 3, False, False, "", "COMBO", 2200, "", "[iInterviewerProfileID]"
        .Add "[iQuestionNo]", "[" & LoadResString(2439) & "]", 4, False, False, "", "INTEGER", 2000, "", "[iQuestionNo]"
        .Add "[vQuestionName]", "[" & LoadResString(2440) & "]", 5, False, False, "", "STRING", 2200, "", "[vQuestionName]"
        
    End With
     
    Call f_void_AddSubjectProfileID
    Call f_void_AddInterviewerProfileId
    
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
        
        'Checking for the enablity/disablity of the save toolbar button or menu option
        If f_int_ExamType = 0 Or f_int_ExamType = 1 Then
                      
            If cboSubjectProfileID.Text <> "" And (txtQuestion.Text = "" Or txtQuestionNumber.Text = "") Then
                fMainForm.mnuToolsSave.Enabled = False
                fMainForm.Toolbar1.Buttons("Save").Enabled = False
            Else
                fMainForm.mnuToolsSave.Enabled = True
                fMainForm.Toolbar1.Buttons("Save").Enabled = True
            End If
       Else
            If cboSubjectProfileID.Text <> "" And (cboInterViewerProfileId.Text = "") Then
                fMainForm.mnuToolsSave.Enabled = False
                fMainForm.Toolbar1.Buttons("Save").Enabled = False
            Else
                fMainForm.mnuToolsSave.Enabled = True
                fMainForm.Toolbar1.Buttons("Save").Enabled = True
            End If
       End If
        
''        fMainForm.Toolbar1.Buttons("Save").Enabled = True
        fMainForm.Toolbar1.Buttons("Cancel").Enabled = True

    End If
End Sub

Public Function ExtraValidation() As Boolean
    On Error GoTo errorhandler
    ' if goes through then ExtraValidation is true if fails then false
    
    ExtraValidation = True
    
    Exit Function
errorhandler:
    MsgBox Err.Description
    
End Function

Private Sub tblcboInterviewerProfileId_Click()
    SetChange
End Sub

Private Sub tblcboRoomProfileId_Click()
    SetChange
End Sub


Private Sub f_void_AddSubjectProfileID()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo errorhandler
    
    l_str_sql = "Select iSubjectProfileId, vSubjectName from tbSTESubjectProfile"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
         
         cboSubjectProfileID.AddItem "", 0
         
         If Not l_obj_Rst.EOF Then
            l_obj_Rst.MoveFirst
         Else
            Exit Sub
         End If
         
         With m_ComboDetails
            While Not l_obj_Rst.EOF
               cboSubjectProfileID.AddItem l_obj_Rst.Fields("vSubjectName")
               m_ComboDetails.Add l_obj_Rst.Fields("iSubjectProfileId"), l_obj_Rst.Fields("vSubjectName"), cboSubjectProfileID.Tag
               l_obj_Rst.MoveNext
            Wend
         End With
         
     l_obj_Rst.Close
     Set l_obj_Rst = Nothing
     
    Exit Sub
errorhandler:
    MsgBox Err.Description
'    Resume
End Sub

Private Sub f_void_AddInterviewerProfileId()
    Dim l_str_sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo errorhandler
    
    l_str_sql = "Select iInterviewerProfileId, vInterviewerName from tbSTEInterviewerProfile"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)
    
        cboInterViewerProfileId.AddItem "", 0
        m_ComboDetails.Add 0, "", cboInterViewerProfileId.Tag
         If Not l_obj_Rst.EOF Then
            l_obj_Rst.MoveFirst
         Else
            Exit Sub
         End If
         
         With m_ComboDetails
            While Not l_obj_Rst.EOF
               cboInterViewerProfileId.AddItem l_obj_Rst.Fields("vInterviewerName")
               m_ComboDetails.Add l_obj_Rst.Fields("iInterviewerProfileId"), l_obj_Rst.Fields("vInterviewerName"), cboInterViewerProfileId.Tag
               l_obj_Rst.MoveNext
            Wend
         End With
         
     l_obj_Rst.Close
     Set l_obj_Rst = Nothing
     
     Exit Sub
errorhandler:
    MsgBox Err.Description
   ' Resume
End Sub

Private Sub txtInterviewRoomProfileId_Change()
'    SetChange
End Sub


Private Sub txtInterviewRoomProfileId_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtQuestion_Change()
SetChange
End Sub

Private Sub txtQuestionNumber_Change()
SetChange
End Sub
