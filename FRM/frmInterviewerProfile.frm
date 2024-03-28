VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmInterviewerProfile 
   Caption         =   "frmInterviewerProfile : 採点者プロファイル"
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
   Picture         =   "frmInterviewerProfile.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   14010
   WindowState     =   2  '最大化
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgsearchgrid 
      Height          =   3855
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   6275
      _ExtentX        =   11060
      _ExtentY        =   6800
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cboGroup 
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
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Tag             =   "[iInterviewGroupProfileId]"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtInterviewerName 
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
      TabIndex        =   3
      Tag             =   "[vInterviewerName]"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtInterviewerProfileId 
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
      Tag             =   "[iInterviewerProfileId]"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  '透明
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
      Left            =   3240
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblExamType 
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
      TabIndex        =   8
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  '透明
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
      Left            =   3240
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  '透明
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
      Left            =   3240
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrorMsg 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
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
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   6255
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInterviewerName 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "1203"
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
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lblInterviewerProfileId 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "1202"
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
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
End
Attribute VB_Name = "frmInterviewerProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmInterviewerProfile
'Author         :   Vishal Kamath
'Created On     :   10/8/01
'Description    :   This form makes a provision for master maintenance of tbSTRInterviewerprofile Table.
'Reference      :   Functional Specs Of MasterMaintenance Ver 1.0
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History   -
'1) 04/04/2002    -   Dileep Cherian
'Description : On activation of the form, only the "new" and "query" toolbar icons should be enabled
'2)'09/05/2002 - Mahesh Deshpande
'Description : Modified to inclde Foreign Key iInterviewGroupProfileId with table tbSTEInterviewGroupProfile
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
Public f_int_PrevRow  As Long

Private Sub cboGroup_Click()

    SetChange

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrHandler
    Dim lngRow As Integer
    Dim Index  As Integer

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
    Call g_void_SetFontProperties(Me)     ' set the font properties

    Me.Caption = "frmInterviewerProfile:  採点者プロファイル"    ''''LoadResString(1201) 採点者プロファイル

    ' set the table name
    m_TableName = "tbSTEInterviewerProfile"

    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained
    With m_colFieldDetails
        .Add "[iInterviewerProfileId]", "[" & LoadResString(1202) & "]", 1, True, True, "", "INT", 2800, "", "[iInterviewerProfileId]"
        .Add "[vInterviewerName]", "[" & LoadResString(1203) & "]", 2, True, False, "", "STRING", 2600, "", "[vInterviewerName]"
        .Add "[iInterviewGroupProfileId]", "[" & LoadResString(2468) & "]", 3, True, True, "", "COMBO", 2800, "", "[iInterviewGroupProfileId]"
    End With

    f_void_AddInterviewGroupID
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

Private Sub txtInterviewerName_Change()

    SetChange

End Sub

Private Sub txtInterviewerProfileId_Change()

'    SetChange

End Sub


Private Sub txtInterviewerProfileId_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

'Populate InterviweGroupId combo (9/5/02)
Private Sub f_void_AddInterviewGroupID()

    On Error GoTo ErrorHandler

    Dim l_str_Sql As String
    Dim l_obj_Rst As ADODB.Recordset
    
    
    l_str_Sql = "Select iInterviewGroupProfileId, vInterviewGroupName from tbSTEInterviewGroupProfile"

    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
         cboGroup.AddItem "", 0
        
         If Not l_obj_Rst.EOF Then
            l_obj_Rst.MoveFirst
         Else
            Exit Sub
         End If
         
         With m_ComboDetails
            While Not l_obj_Rst.EOF
               cboGroup.AddItem Trim(l_obj_Rst.Fields("vInterviewGroupName"))
               m_ComboDetails.Add l_obj_Rst.Fields("iInterviewGroupProfileId"), l_obj_Rst.Fields("vInterviewGroupName"), "[iInterviewGroupProfileId]"
               l_obj_Rst.MoveNext
            Wend
         End With
         
     l_obj_Rst.Close
     Set l_obj_Rst = Nothing
     
    Exit Sub

ErrorHandler:
    MsgBox Err.Description
'    Resume

End Sub

