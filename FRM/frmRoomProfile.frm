VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmRoomProfile 
   Caption         =   "Form1"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13650
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmRoomProfile.frx":0000
   ScaleHeight     =   9735
   ScaleWidth      =   13650
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.TextBox txtRoomProfileId 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   5325
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   6
      Tag             =   "[iRoomProfileId]"
      Top             =   8460
      Width           =   600
   End
   Begin VB.TextBox txtRoomName 
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
      Height          =   405
      Left            =   3525
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "[vRoomName]"
      Top             =   1410
      Width           =   1245
   End
   Begin VB.TextBox txtRandom 
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
      Height          =   405
      Left            =   5100
      MaxLength       =   15
      TabIndex        =   4
      Tag             =   "[iRandom]"
      Top             =   1410
      Width           =   930
   End
   Begin VB.TextBox txtMaxCapacity 
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
      Height          =   405
      Left            =   6435
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "[iMaxCapacity]"
      Top             =   1410
      Width           =   1365
   End
   Begin VB.ComboBox cboInterviewRoomFlag 
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
      Left            =   315
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   2
      Tag             =   "[iInterviewRoomFlag]"
      Top             =   1410
      Width           =   2685
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgsearchgrid 
      Height          =   6540
      Left            =   2550
      TabIndex        =   1
      Top             =   1890
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   11536
      _Version        =   393216
      Cols            =   4
      HighLight       =   0
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
      _Band(0).Cols   =   4
   End
   Begin VB.Label lblTit2 
      BackStyle       =   0  'ìßñæ
      Caption         =   "(É}ÉXÉ^Å[TableÇÃÉåÉRÅ[ÉhNoÇï\Ç∑ÅBämîFÇÃÇΩÇﬂï\é¶ÇµÇƒÇ¢Ç‹Ç∑)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2550
      TabIndex        =   18
      Top             =   8790
      Width           =   5430
   End
   Begin VB.Label lblTit 
      BackStyle       =   0  'ìßñæ
      Caption         =   "âÔèÍÅEñ ê⁄ÉOÉãÅ[ÉvÇÃ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2550
      TabIndex        =   17
      Top             =   8505
      Width           =   2040
   End
   Begin VB.Label lblRoomProfileId 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "ä«óù ID"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   8505
      Width           =   840
   End
   Begin VB.Label lblRoomName 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "âÔèÍñº"
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
      Left            =   3645
      TabIndex        =   15
      Top             =   1140
      Width           =   1065
   End
   Begin VB.Label lblErrIndicator 
      Alignment       =   2  'íÜâõëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   5910
      TabIndex        =   14
      Top             =   8475
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3345
      TabIndex        =   13
      Top             =   1485
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   6300
      TabIndex        =   12
      Top             =   1470
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   4950
      TabIndex        =   11
      Top             =   1485
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblRandom 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "óêêî"
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
      Left            =   5130
      TabIndex        =   10
      Top             =   1125
      Width           =   840
   End
   Begin VB.Label lblMaxCapacity 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "íËàı"
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
      Left            =   6420
      TabIndex        =   9
      Top             =   1125
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "âÔèÍ(àÍéü)/ñ ê⁄Å@ãÊï™"
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
      Left            =   255
      TabIndex        =   8
      Top             =   1110
      Width           =   2670
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  'ìßñæ
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   7
      Top             =   1470
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblErrorMsg 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  'ìßñæ
      Caption         =   "lblErrorMsg"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2550
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   9375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRoomProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmRoomProfile
'Author         :   Vishal Kamath
'Created On     :   10/8/01
'Description    :   This form makes a provision for master maintenance of tbSTRRoomProfile Table.
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

Private Sub cboInterviewRoomFlag_Click()
    SetChange
End Sub

Private Sub Form_Load()

    On Error GoTo ErrHandler

    LoadResStrings Me
    Me.Caption = "frmRooMProfile : âÔèÍÅEñ ê⁄ÉOÉãÅ[Év"  ''''LoadResString(1501)

    Call g_void_SetFontProperties(Me)     ' set the font properties


    ' set the table name
    m_TableName = "tbSTERoomProfile"

    'Fill the m_colFieldDetails with all the fields corresponding to the table to be maintained


''''----------------------------------------------------------------------------
''''2021.12.09 del jhi
''''----------------------------------------------------------------------------
''''    With m_colFieldDetails
''''        .Add "[iRoomProfileId]", "[" & LoadResString(1602) & "]", 1, True, True, "", "INTEGER", 2000, "", "[iRoomProfileId]"
''''        .Add "[vRoomName]", "[" & LoadResString(1603) & "]", 2, True, False, "", "STRING", 1500, "", "[vRoomName]"
''''        'Changed by Mahesh 21/5/2002 Random is not mandatory
''''        .Add "[iRandom]", "[" & LoadResString(1604) & "]", 3, False, False, "", "INTEGER", 1500, "", "[iRandom]"
''''        'Changes end
''''        .Add "[iMaxCapacity]", "[" & LoadResString(1605) & "]", 4, True, False, "", "INTEGER", 1600, "", "[iMaxCapacity]"
''''        .Add "[iInterviewRoomFlag]", "[" & LoadResString(2506) & "]", 5, True, False, "", "COMBO", 1600, "", "[iInterviewRoomFlag]"     ' change made on 31/07/02
''''    End With


''''----------------------------------------------------------------------------
''''2021.12.09 add jhi
''''----------------------------------------------------------------------------
    With m_colFieldDetails
        .Add "[iRoomProfileId]", "[" & "âÔèÍÅEñ ê⁄ÉOÉãÅ[ÉvID" & "]", 1, True, True, "", "INTEGER", 0, "", "[iRoomProfileId]"
        .Add "[vRoomName]", "[" & "Å@Å@âÔèÍñº" & "]", 2, True, False, "", "STRING", 1500, "", "[vRoomName]"

        'Changed by Mahesh 21/5/2002 Random is not mandatory
        .Add "[iRandom]", "[" & "óêêî        " & "]", 3, False, False, "", "INTEGER", 1500, "", "[iRandom]"
        'Changes end
        .Add "[iMaxCapacity]", "[" & "íËàı(éÛå±ê∂)  " & "]", 4, True, False, "", "INTEGER", 1500, "", "[iMaxCapacity]"
        .Add "[iInterviewRoomFlag]", "[" & "ñ ê⁄ÉOÉãÅ[Év" & "]", 5, True, False, "", "COMBO", 1600, "", "[iInterviewRoomFlag]"     ' change made on 31/07/02
    End With


''''----------------------------------------------------------------------------
''''2021.12.09 del jhi
''''----------------------------------------------------------------------------
''''    With cboInterviewRoomFlag
''''        .AddItem ""
''''        .AddItem LoadResString(2507)
''''        .AddItem LoadResString(2508)
''''        m_ComboDetails.Add 0, LoadResString(2507), "[iInterviewRoomFlag]"
''''        m_ComboDetails.Add 1, LoadResString(2508), "[iInterviewRoomFlag]"
''''    End With

''''----------------------------------------------------------------------------
''''2021.12.09 add jhi
''''----------------------------------------------------------------------------
    With cboInterviewRoomFlag
        .AddItem ""
        .AddItem LoadResString(2508) ''''í èÌééå±âÔèÍ <---èáèòÇïœçXÇµÇΩÅB
        .AddItem LoadResString(2507) ''''ñ ê⁄ÉOÉãÅ[Év

        m_ComboDetails.Add 0, LoadResString(2507), "[iInterviewRoomFlag]" ''''ñ ê⁄ÉOÉãÅ[Év
        m_ComboDetails.Add 1, LoadResString(2508), "[iInterviewRoomFlag]" ''''í èÌééå±âÔèÍ
    End With

    m_bload = False

    lblTit.FontSize = 10
    lblTit2.FontSize = 10
    lblRoomProfileId.FontSize = 10
    lblRoomProfileId.ForeColor = &HC000&

    txtRoomProfileId.FontSize = 10
    txtRoomProfileId.Height = 270
    txtRoomProfileId.BackColor = &H80FFFF

    Exit Sub

ErrHandler:
    MsgBox Err.Description

End Sub

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
        m_lngSearchGridTop = 1600

        'initialize the search grid
        Call InitializeSearchGrid

        'populate the grid with empty row and column headings
        Call SearchRecords(True)

        fMainForm.mnuToolsSave.Enabled = False
        fMainForm.mnuToolsCancel.Enabled = False
        fMainForm.Toolbar1.Buttons("Save").Enabled = False
        fMainForm.Toolbar1.Buttons("Cancel").Enabled = False

    End If

    m_bload = True
    fMainForm.mnuTools.Enabled = True


    '---------------------------------------------------------------------------
    'èâä˙ÉfÅ[É^ï\é¶
    '2022.01.05 add jhi
    '---------------------------------------------------------------------------
    Call ClearData     'sougannkyo
    Call SearchRecords 'New icon
    Call NewData

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
    Call g_void_HighlightRow(hfgSearchGrid.Row, f_int_PrevRow)
    f_int_PrevRow = hfgSearchGrid.Row
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

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub txtMaxCapacity_Change()
    SetChange
End Sub

Private Sub txtMaxCapacity_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRandom_Change()
    SetChange
End Sub

Private Sub txtRandom_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRoomName_Change()
    SetChange
End Sub

Private Sub txtRoomProfileId_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
