VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSubjectQuestionProfile 
   Caption         =   "2458"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmSubjectQuestionProfile.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   14700
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.TextBox txtQuestion 
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
      Left            =   3165
      TabIndex        =   7
      Tag             =   "[vQuestionName]"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtQuestionNumber 
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
      Left            =   3165
      TabIndex        =   5
      Tag             =   "[iQuestionNo]"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtSubjectQuestionId 
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
      Left            =   3165
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "[iSubjectQuestionId]"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox cboSubjectProfileID 
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
      Left            =   3165
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   3
      Tag             =   "[iSubjectProfileId]"
      Top             =   1560
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgsearchgrid 
      Height          =   3855
      Left            =   255
      TabIndex        =   13
      Top             =   3480
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   6800
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblErrorMsg 
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
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   7650
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblErrIndicator 
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
      Height          =   135
      Index           =   3
      Left            =   2880
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
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
      Height          =   225
      Index           =   4
      Left            =   2880
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
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
      Height          =   150
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
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
      Height          =   210
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "2457"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   2490
   End
   Begin VB.Label Label4 
      Caption         =   "2456"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2490
   End
   Begin VB.Label Label2 
      Caption         =   "2454"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2490
   End
   Begin VB.Label lblCode 
      Caption         =   "2453"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
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
      Top             =   1080
      Width           =   2490
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
'Created On     :   10/8/01
'Description    :   This form makes a provision for master maintenance of tbSTESubjectQuestionProfile Table.
'Reference      :   Functional Specs Of MasterMaintenance Ver 1.0
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History   -   04/04/2002  -   Dileep Cherian
'On activation of the form, only the "new" and "query" toolbar icons should be enabled
'21/5/2002 Mahesh
'Removed Interviewer Id combo, only Examtype 0 and one subjects 2 b listed in Subject combo
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
Public f_int_PrevRow As Long

Private Sub cboInterViewerProfileId_Click()
    SetChange
End Sub

Private Sub cboSubjectProfileID_Click()
    Dim l_obj_subject As New ADODB.Recordset
    Dim l_str_Sql As String
    
    On Error GoTo ErrorHandler
    SetChange
        'Commented by Mahesh 21/5/2002
'    If cboSubjectProfileID.Text <> "" Then
'        l_str_Sql = "SELECT iExamType FROM tbSTESubjectProfile WHERE vSubjectName = " & "'" & cboSubjectProfileID.Text & "'"
'        l_obj_subject.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
'        f_int_ExamType = l_obj_subject.Fields("iExamType").Value
'        If f_int_ExamType = 0 Or f_int_ExamType = 1 Then
'            If cboInterViewerProfileId.Text <> "" Then
'                cboInterViewerProfileId.ListIndex = 0
'       Comments end
'                fMainForm.mnuToolsSave.Enabled = False
'                fMainForm.Toolbar1.Buttons("Save").Enabled = False
        'Commented by Mahesh 21/5/2002
'            End If
'            cboInterViewerProfileId.Enabled = False
'            txtQuestionNumber.Enabled = True
'            txtQuestion.Enabled = True
'        Else
'            cboInterViewerProfileId.Enabled = True
'            txtQuestionNumber.Enabled = False
'            txtQuestion.Enabled = False
'        End If
'Comments end
'    End If
    Set l_obj_subject = Nothing
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
    Dim lngRow As Integer
    Dim index As Integer
    On Error GoTo ErrHandler
'
    For index = 1 To fMainForm.Toolbar1.Buttons.Count
       If index = 1 Or index = 6 Then
            fMainForm.Toolbar1.Buttons(index).Enabled = True
        Else
            fMainForm.Toolbar1.Buttons(index).Enabled = False
        End If
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
    Me.Caption = LoadResString(2458)
    Call g_void_SetFontProperties(Me)     ' set the font properties
    ' set the table name
    m_TableName = "tbSTESubjectQuestionProfile"
    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained
    With m_colFieldDetails
        .Add "[iSubjectQuestionId]", "[" & LoadResString(2453) & "]", 1, True, True, "", "INTEGER", 2500, "", "[iSubjectQuestionId]"
        .Add "[iSubjectProfileId]", "[" & LoadResString(2454) & "]", 2, True, False, "", "COMBO", 2500, "", "[iSubjectProfileId]"
        'Changes by Mahesh 21/5/2002 - No interviewer in the master now
        '.Add "[iInterviewerProfileID]", "[" & LoadResString(2455) & "]", 3, False, False, "", "COMBO", 2200, "", "[iInterviewerProfileID]"
        'changes end
        .Add "[iQuestionNo]", "[" & LoadResString(2456) & "]", 3, True, False, "", "INTEGER", 2000, "", "[iQuestionNo]"
        .Add "[vQuestionName]", "[" & LoadResString(2457) & "]", 4, True, False, "", "STRING", 2200, "", "[vQuestionName]"
    End With
     
    Call f_void_AddSubjectProfileID
    'Changes by Mahesh 21/5/2002 - No interviewer in the master now
    'Call f_void_AddInterviewerProfileId
    'End
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
    'Changes by Mahesh 21/5/2002. Disable question no and question desc text boxes if examtype is 0,1
    Dim l_str_Sql As String
    Dim l_obj_subject As New ADODB.Recordset
    Dim l_int_oldRow As Integer
    Dim l_int_OldCol As Integer
    With hfgsearchgrid
        l_int_oldRow = .Row
        l_int_OldCol = .Col
        .Col = 2
        l_str_Sql = "SELECT iExamType FROM tbSTESubjectProfile WHERE iSubjectProfileId= " & "'" & Trim(.Text) & "'"
        l_obj_subject.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
        f_int_ExamType = l_obj_subject.Fields("iExamType").Value
        If f_int_ExamType = 0 Or f_int_ExamType = 1 Then
            txtQuestionNumber.Enabled = True
            txtQuestion.Enabled = True
        Else
            txtQuestionNumber.Enabled = False
            txtQuestion.Enabled = False
        End If
        .Row = l_int_oldRow
        .Col = l_int_OldCol
    
    End With
    'Changes end
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
        
        'Checking for the enablity/disablity of the save toolbar button or menu option
        If f_int_ExamType = 0 Or f_int_ExamType = 1 Then
                      
            If cboSubjectProfileID.Text <> "" And (txtQuestion.Text = "" Or txtQuestionNumber.Text = "") Then
                fMainForm.mnuToolsSave.Enabled = False
                fMainForm.Toolbar1.Buttons("Save").Enabled = False
            Else
                fMainForm.mnuToolsSave.Enabled = True
                fMainForm.Toolbar1.Buttons("Save").Enabled = True
            End If
'       Else
'            If cboSubjectProfileID.Text <> "" And (cboInterViewerProfileId.Text = "") Then
'                fMainForm.mnuToolsSave.Enabled = False
'                fMainForm.Toolbar1.Buttons("Save").Enabled = False
'            Else
'                fMainForm.mnuToolsSave.Enabled = True
'                fMainForm.Toolbar1.Buttons("Save").Enabled = True
'            End If
       End If
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

Private Sub tblcboInterviewerProfileId_Click()
    SetChange
End Sub

Private Sub tblcboRoomProfileId_Click()
    SetChange
End Sub

Private Sub f_void_AddSubjectProfileID()
    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    l_str_Sql = "Select iSubjectProfileId, vSubjectName from tbSTESubjectProfile where iExamType=0 or iExamType =1"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
         
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
ErrorHandler:
    MsgBox Err.Description
End Sub

'Private Sub f_void_AddInterviewerProfileId()
'    Dim l_str_Sql As String
'    Dim l_obj_Rst As ADODB.Recordset
'
'    On Error GoTo ErrorHandler
'
'    l_str_Sql = "Select iInterviewerProfileId, vInterviewerName from tbSTEInterviewerProfile"
'    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
'
'        cboInterViewerProfileId.AddItem "", 0
'        m_ComboDetails.Add 0, "", cboInterViewerProfileId.Tag
'         If Not l_obj_Rst.EOF Then
'            l_obj_Rst.MoveFirst
'         Else
'            Exit Sub
'         End If
'
'         With m_ComboDetails
'            While Not l_obj_Rst.EOF
'               cboInterViewerProfileId.AddItem l_obj_Rst.Fields("vInterviewerName")
'               m_ComboDetails.Add l_obj_Rst.Fields("iInterviewerProfileId"), l_obj_Rst.Fields("vInterviewerName"), cboInterViewerProfileId.Tag
'               l_obj_Rst.MoveNext
'            Wend
'         End With
'
'     l_obj_Rst.Close
'     Set l_obj_Rst = Nothing
'
'     Exit Sub
'ErrorHandler:
'    MsgBox Err.Description
'End Sub

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

Private Sub txtQuestionNumber_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
