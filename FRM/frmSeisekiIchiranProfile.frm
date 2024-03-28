VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSeisekiIchiranProfile 
   Caption         =   "4577"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmSeisekiIchiranProfile.frx":0000
   ScaleHeight     =   9960
   ScaleWidth      =   12345
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdGetId 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   14
      Top             =   1020
      Width           =   495
   End
   Begin VB.TextBox txtSrNo 
      Height          =   405
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   12
      Top             =   1680
      Width           =   495
   End
   Begin MSComCtl2.UpDown udCmdBtnPage 
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   8760
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "4408"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   9
      Top             =   1050
      Width           =   1470
   End
   Begin VB.TextBox txtReportId 
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   9000
      TabIndex        =   7
      Top             =   1020
      Width           =   1005
   End
   Begin VB.CommandButton f_cmd_ButtonArray 
      BackColor       =   &H0080C0FF&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   3195
      TabIndex        =   2
      Top             =   2730
      Width           =   1290
   End
   Begin VB.CommandButton f_cmd_ButtonArray 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3195
      TabIndex        =   1
      Top             =   2205
      Width           =   1290
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfTeiki 
      Height          =   7410
      Left            =   240
      TabIndex        =   0
      Top             =   1965
      Width           =   2835
      _cx             =   5001
      _cy             =   13070
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfG2 
      Height          =   7830
      Left            =   4575
      TabIndex        =   3
      Top             =   1530
      Width           =   3615
      _cx             =   6376
      _cy             =   13811
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfG3 
      Height          =   7830
      Left            =   8445
      TabIndex        =   4
      Top             =   1545
      Width           =   3615
      _cx             =   6376
      _cy             =   13811
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "SrNo"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   1740
      Width           =   615
   End
   Begin VB.Label lblCmdBtnPage 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      BorderStyle     =   1  '実線
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label lblReportId 
      Caption         =   "4586"
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
      Height          =   405
      Left            =   6960
      TabIndex        =   8
      Tag             =   "4586"
      Top             =   1020
      Width           =   1800
   End
   Begin VB.Label lblError 
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
      Height          =   360
      Left            =   225
      TabIndex        =   6
      Top             =   1065
      Width           =   6660
   End
   Begin VB.Label lblTeiki 
      Caption         =   "4415"
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
      Width           =   2835
   End
   Begin VB.Line Line4 
      X1              =   11400
      X2              =   11400
      Y1              =   1200
      Y2              =   2040
   End
End
Attribute VB_Name = "frmSeisekiIchiranProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmSeisekiIchiranProfile
'Author         :   Dileep Cherian
'Created On     :   16/08/2002
'Description    :   To Print Reports of SeisekiIchiranProfile
'Reference      :   Functional Specs Of SeisekiIchiranProfile Ver 1.0
'*************************************************************************************************

Private Const prviShowButtons As Integer = 10
Private Const prvlBtnBCR_DEF As Long = &H8000000F
Private Const prvlBtnBCR_NeedSub As Long = &H80C0FF

Private Sub cmdGetId_Click()
    
Dim sID As String
    
    Call dlgSeisekiChohyoIchiran.getPrintCommandId(sID)

    txtReportId.Text = sID

End Sub

Private Sub cmdSearch_Click()
    If Len(txtReportId.Text) = 0 Then
'        lblError.Caption = LoadResString(4587)
    Else
        lblError.Caption = ""
        Call f_void_PopulateG2(CLng(txtReportId.Text))
    End If
End Sub

Private Sub f_cmd_ButtonArray_Click(Index As Integer)
    If txtReportId.Text = "" Then
        Exit Sub
    Else
        If Trim(txtReportId.Text) = "" Then
            Exit Sub
        End If
    End If
    Select Case Index
    Case 0      ' add resords from teiki/kakushu grids to the grid G2, on click of ">" button
        Call f_void_ForwardArrow
    Case 1      ' remove records from G2 and G3 grids, based on selection in grid G2
        Call f_void_DeleteRecords
    Case Else   ' add records from teiki/kakushu grids to G2, on click of dynamic buttons
        Call f_void_ProcessDynamicButtons(f_cmd_ButtonArray(Index))
    End Select

    Call f_void_PopulateG2(CLng(Me.txtReportId.Text))
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    fMainForm.mnuTools.Enabled = False  ' disable tools menu
    Dim Index
    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Deactivate()

    fMainForm.mnuHelp.Visible = False

End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Call LoadResStrings(Me)
    Call g_void_SetFontProperties(Me)
    
    txtReportId.Text = 1
    lblTeiki.Alignment = 0

    Call f_void_InitGrid(vsfTeiki)
    Call f_void_InitGrid(vsfG2)
    Call f_void_InitGrid(vsfG3)
    
    ' initialize the grids
    Call f_void_InitSubjectGrids(vsfTeiki)
    Call f_void_InitSubjectGrids(vsfG2)
    Call f_void_InitSubjectGrids(vsfG3)
    
    vsfG2.Rows = 1
    vsfG3.Rows = 1
    
    ' display the dynamic buttons
    Call f_void_DisplayButtons

    Call f_void_ShowButtons

    Call f_void_PopulateSubjectGrids

    Call f_void_SetReportNo

    fMainForm.mnuPrint.Enabled = True
    fMainForm.Toolbar1.Buttons("Print").Enabled = True

    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)

    fMainForm.mnuHelp.Visible = False

End Sub

Private Sub txtReportId_KeyPress(KeyAscii As Integer)
    ' allow only only integers in the nendo textbox
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub f_void_InitSubjectGrids(ByRef l_grd_MyGrid As VSFlexGrid)
    With l_grd_MyGrid
        .Rows = 2
        .Cols = 6           ' last 3 are hidden columns
                            ' col 0 - serial no
                            ' col 1 - subject name
                            ' col 2 - selected subject names/subject grade profile ids (hidden)
                            ' col 3 - seiseki ichiran id (hidden)
                            ' col 4 - seiseki special subject profile id (hidden)
                            ' col 5 - subject question profile id (hidden)
        .FixedCols = 1
        .FixedRows = 1
        
        .Row = 0
        .Col = 0
        .ColWidth(0) = 800
'        .Text = LoadResString(350)

        .Col = 1
        If .Name = "vsfG2" Then
            .ColWidth(1) = 2500
'            .Text = LoadResString(4585)
        ElseIf .Name = "vsfG3" Then
            .ColWidth(1) = 2500
'            .Text = LoadResString(552)
        Else
            .ColWidth(1) = 1700
'            .Text = LoadResString(552)
        End If
        .Col = 2    ' (hidden)
        .ColWidth(2) = 0
        .Col = 3    ' (hidden)
        .ColWidth(3) = 0
        .Col = 4    ' (hidden)
        .ColWidth(4) = 0
        .Col = 5    ' (hidden)
        .ColWidth(5) = 0
    End With
End Sub

Private Sub f_void_PopulateSubjectGrids()
    Dim sSQL As String                 ' sql string to get the subjects
    Dim l_obj_rsSubjects As ADODB.Recordset     ' recordset object to get the subjects
    Dim l_int_Counter As Integer                    ' counter variable
    
    On Error GoTo ErrorHandler
    
    lblError.Caption = ""
    
    'enable the buttons
    For l_int_Counter = 0 To f_cmd_ButtonArray.Count - 1
        f_cmd_ButtonArray(l_int_Counter).Enabled = True
    Next l_int_Counter
    
    ' clear G2 and G3 when the selection changes
    vsfG2.Rows = 1
    vsfG3.Rows = 1

    ' select subjects based on input parameters - nendo, gakunen and gakki
    sSQL = "SELECT a.iSubjectProfileId "
    sSQL = sSQL & " , case when b.iSubjectQuestionId is null then a.vSubjectName "
    sSQL = sSQL & "        else b.vQuestionName end as vSubjectName "
    sSQL = sSQL & " , b.iSubjectQuestionId "
    sSQL = sSQL & " FROM tbSTESubjectProfile a "
    sSQL = sSQL & " LEFT OUTER JOIN tbSTESubjectQuestionProfile b "
    sSQL = sSQL & " ON b.iSubjectProfileId = a.iSubjectProfileId "
    sSQL = sSQL & " ORDER BY a.iExamType , a.iDispOrder , b.iQuestionNo "

    Set l_obj_rsSubjects = g_obj_Conn.Execute("select 1 ")

    l_obj_rsSubjects.Close
    Set l_obj_rsSubjects = Nothing

    Set l_obj_rsSubjects = g_obj_Conn.Execute(sSQL)

    vsfTeiki.Rows = 1
    Do While Not l_obj_rsSubjects.EOF

        l_int_Counter = l_int_Counter + 1
        With vsfTeiki
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
            .Text = .Rows - 1
            .Col = 1
            .Text = l_obj_rsSubjects.Fields("vSubjectName").Value
            .Col = 2
            .Text = l_obj_rsSubjects.Fields("iSubjectProfileId").Value
            .Col = 5
            If IsNull(l_obj_rsSubjects.Fields("iSubjectQuestionId").Value) Then
                .Text = ""
            Else
                .Text = l_obj_rsSubjects.Fields("iSubjectQuestionId").Value
            End If
        End With
                
        l_obj_rsSubjects.MoveNext
    Loop
    
    l_obj_rsSubjects.Close
    Set l_obj_rsSubjects = Nothing

    vsfTeiki.Row = 0

    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub f_void_ShowButtons()

Dim iPage As Integer
Dim iLoopCnt As Integer
Dim oCmd As Object

    iPage = CInt(lblCmdBtnPage.Caption)

    For Each oCmd In f_cmd_ButtonArray
        oCmd.Visible = False
    Next

    f_cmd_ButtonArray(0).Visible = True
    f_cmd_ButtonArray(1).Visible = True

    For iLoopCnt = 0 To prviShowButtons - 1
        If (iPage - 1) * prviShowButtons + iLoopCnt + 2 >= f_cmd_ButtonArray.Count Then Exit For
        f_cmd_ButtonArray((iPage - 1) * prviShowButtons + iLoopCnt + 2).Visible = True
    Next

End Sub

Private Sub f_void_DisplayButtons()
' pick up and display dynamic buttons from tbSTRSeisekiSpecialProfile table
Dim l_int_Counter As Integer                ' counter variable
Dim sSQL As String              ' sql query
Dim l_obj_rsButtons As ADODB.Recordset  ' recordset variable
Dim oCn As New ADODB.Connection

    sSQL = "Select "
    sSQL = sSQL & "  iSpecialProfileId "
    sSQL = sSQL & ", vButtonName "
    sSQL = sSQL & ", isnull( iDispOrder , 99999 ) as iDispOrder "
    sSQL = sSQL & ", isnull(siNeedSubFlag , 0 ) as siNeedSubFlag "
    sSQL = sSQL & " FROM tbSTESeisekiSpecialProfile "
    sSQL = sSQL & " where iDispOrder <> -1 "
    sSQL = sSQL & " order by "
    sSQL = sSQL & "   isnull( iDispOrder , 99999 ) "
    sSQL = sSQL & " , iSpecialProfileId"

'    Call f_void_OpenConnection(oCn)
'    Set l_obj_rsButtons = oCn.Execute(sSQL)
    Set l_obj_rsButtons = g_obj_Conn.Execute(sSQL)

    Do While Not l_obj_rsButtons.EOF
        l_int_Counter = f_cmd_ButtonArray.Count
        Load f_cmd_ButtonArray(l_int_Counter)
        f_cmd_ButtonArray(l_int_Counter).Caption = l_obj_rsButtons.Fields("vButtonName").Value
        If ((l_int_Counter - 2) Mod prviShowButtons) = 0 Then
            f_cmd_ButtonArray(l_int_Counter).Top = f_cmd_ButtonArray(1).Top + 525
        Else
            f_cmd_ButtonArray(l_int_Counter).Top = f_cmd_ButtonArray(l_int_Counter - 1).Top + 525
        End If
        f_cmd_ButtonArray(l_int_Counter).Width = f_cmd_ButtonArray(l_int_Counter - 1).Width
        f_cmd_ButtonArray(l_int_Counter).Height = f_cmd_ButtonArray(l_int_Counter - 1).Height
        f_cmd_ButtonArray(l_int_Counter).Left = f_cmd_ButtonArray(l_int_Counter - 1).Left
        f_cmd_ButtonArray(l_int_Counter).Tag = l_obj_rsButtons.Fields("iSpecialProfileId").Value
        If l_obj_rsButtons.Fields(3).Value = 1 Then
            f_cmd_ButtonArray(l_int_Counter).BackColor = prvlBtnBCR_NeedSub
        Else
            f_cmd_ButtonArray(l_int_Counter).BackColor = prvlBtnBCR_DEF
        End If
        f_cmd_ButtonArray(l_int_Counter).Visible = True
        l_obj_rsButtons.MoveNext
    Loop
    
    l_obj_rsButtons.Close
    Set l_obj_rsButtons = Nothing

'    oCn.Close
'    Set oCn = Nothing

End Sub

Private Sub f_void_ForwardArrow()
    Dim l_int_Counter As Integer                    ' counter variable for outer loop
    Dim l_int_InnerCounter As Integer               ' counter variable for inner loop
    Dim l_int_SesekiIchiranId As Long            ' to get the latest Seiseki Ichiran Id
    Dim l_str_sqlGetId As String                    ' sql string to get the latest Seiseki Ichiran Id
    Dim l_obj_rsGetId As New ADODB.Recordset        ' recordset object to get the latest Seiseki Ichiran Id
    Dim l_str_sqlInsert As String                   ' insert statement
    Dim l_obj_rsTableMapping As New ADODB.Recordset ' recordset object to get the Seiseki Ichiran Id from tableidmapping table
    Dim l_str_sqlTableMapping As String             ' qsl string to get the Seiseki Ichiran Id from tableidmapping table
    Dim l_int_SelectedTeiki() As Integer            ' array to store the selected teiki rows
    Dim l_int_SelectedKakushu() As Integer          ' array to store the selected Kakushu rows
    Dim l_int_TeikiCounter As Integer               ' teiki counter
    Dim l_int_KakushuCounter As Integer             ' kakushu counter
    Dim l_int_SrNo As Integer                       ' serial number
    
    
    On Error GoTo ErrorHandler
    
    ' initialize the counter variables
    l_int_TeikiCounter = 0
    l_int_KakushuCounter = 0
    
    ' loop thru the teiki grid and store the selected rows into array
    ' this is done becasue later while programatically fetching the values for inserting,
    ' the orininally selected rows gets lost - so store it initially and then laster retrieve it
    With vsfTeiki

    For l_int_Counter = 1 To .Rows - 1
        If .IsSelected(l_int_Counter) Then
            ReDim Preserve l_int_SelectedTeiki(l_int_TeikiCounter)
            l_int_SelectedTeiki(l_int_TeikiCounter) = l_int_Counter
            l_int_TeikiCounter = l_int_TeikiCounter + 1
        End If
    Next l_int_Counter
    
    End With
    
    ' if no subjects are selected in both the subject grids then display error
    If l_int_TeikiCounter = 0 Then
        lblError.Caption = LoadResString(4078)
        Exit Sub
    End If
    
    lblError.Caption = ""
'
'    ' get the latest SeisekiIchiranId
'    l_str_sqlGetId = "SELECT ISNULL( Max( iSeisekiIchiranId ) + 1 , -1 ) as iSeisekiIchiranId FROM tbSTESeisekiIchiranProfile"
'    l_obj_rsGetId.Open l_str_sqlGetId, g_obj_Conn
'
'    If l_obj_rsGetId.Fields("iSeisekiIchiranId").Value <> -1 Then
'        l_int_SesekiIchiranId = l_obj_rsGetId.Fields("iSeisekiIchiranId").Value
'    Else
'        ' no records - extract fiorst id from tableidmapping table
'        l_str_sqlTableMapping = "Select iTableCounterIdMapping from tbSTETableIdMapping where vTableName = 'tbSTESeisekiIchiranProfile'"
'        l_obj_rsTableMapping.Open l_str_sqlTableMapping, g_obj_Conn
'        If Not l_obj_rsTableMapping.EOF Then
'            l_int_SesekiIchiranId = l_obj_rsTableMapping.Fields("iTableCounterIdMapping").Value
'        Else
'            ' no records - initialize to 1
'            l_int_SesekiIchiranId = 1
'        End If
'        l_obj_rsTableMapping.Close
'        Set l_obj_rsTableMapping = Nothing
'    End If
'
'    l_obj_rsGetId.Close
'    Set l_obj_rsGetId = Nothing
Dim bRtn As Boolean
    bRtn = getNewId("tbSTESeisekiIchiranProfile", "iSeisekiIchiranId", l_int_SesekiIchiranId)

    ' get the latest SeisekiIchiranId
    l_str_sqlGetId = "SELECT ISNULL( Max(iSerialNo) + 1 , 1 ) as iSerialNo FROM tbSTESeisekiIchiranProfile"
    l_str_sqlGetId = l_str_sqlGetId & " WHERE iReportNo = " & txtReportId.Text
    l_obj_rsGetId.Open l_str_sqlGetId, g_obj_Conn

    l_int_SrNo = l_obj_rsGetId.Fields("iSerialNo").Value
    If Trim(Me.txtSrNo.Text) <> "" Then
        If CInt(Me.txtSrNo.Text) < l_int_SrNo Then
            '入力シリアル番号より大きいシリアル番号のレコードを自レコードのシリアル番号＋１にて更新
            l_str_sqlGetId = "Update tbSTESeisekiIchiranProfile SET iSerialNo = iSerialNo + 1 "
            l_str_sqlGetId = l_str_sqlGetId & " WHERE iReportNo = " & txtReportId.Text
            l_str_sqlGetId = l_str_sqlGetId & " AND iSerialNo >= " & Me.txtSrNo.Text
            g_obj_Conn.Execute l_str_sqlGetId
            l_int_SrNo = CInt(Me.txtSrNo.Text)
        End If
    End If

    l_obj_rsGetId.Close
    Set l_obj_rsGetId = Nothing

    g_obj_Conn.BeginTrans
    
    ' loop thru the teiki grid and insert data of selected rows into tbSTRSeisekiIchiranProfile table
    With vsfTeiki
            
    For l_int_Counter = 1 To .Rows - 1
        For l_int_InnerCounter = 0 To l_int_TeikiCounter - 1
            If l_int_Counter = l_int_SelectedTeiki(l_int_InnerCounter) Then
                .Row = l_int_Counter
                ' special profile id will go as null
                l_str_sqlInsert = " INSERT INTO tbSTESeisekiIchiranProfile VALUES("
                ' primary key
                l_str_sqlInsert = l_str_sqlInsert & l_int_SesekiIchiranId & ","
                ' report id
                If Len(txtReportId.Text) > 0 Then
                    l_str_sqlInsert = l_str_sqlInsert & txtReportId.Text & ","
                Else    ' insert default value as 1
                    l_str_sqlInsert = l_str_sqlInsert & "1,"
                End If
                ' serial id 画面入力ｏｒ最大、入力＞最大は最大にする
                'とりあえず、最大のみ
                l_str_sqlInsert = l_str_sqlInsert & l_int_SrNo & ","
                ' special id
                l_str_sqlInsert = l_str_sqlInsert & "null,"
                .Col = 2    ' subject profile id
                l_str_sqlInsert = l_str_sqlInsert & (.Text) & ","
                .Col = 5    ' subject question profile id
                l_str_sqlInsert = l_str_sqlInsert & IIf(.Text = "", "null", .Text) & ","

                l_str_sqlInsert = l_str_sqlInsert & "'" & Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
                
                g_obj_Conn.Execute l_str_sqlInsert
                                
                .Col = 1
                vsfG2.Rows = vsfG2.Rows + 1
                vsfG2.Row = vsfG2.Rows - 1
                vsfG2.Col = 0
                vsfG2.Text = vsfG2.Rows - 1
                vsfG2.Col = 1
                vsfG2.Text = .Text
                vsfG2.Col = 3
                vsfG2.Text = l_int_SesekiIchiranId

                l_int_SesekiIchiranId = l_int_SesekiIchiranId + 1
                l_int_SrNo = l_int_SrNo + 1

                Exit For
            End If
        Next l_int_InnerCounter
    Next l_int_Counter
    
    End With
    
    g_obj_Conn.CommitTrans
    
    Exit Sub
ErrorHandler:
    If g_obj_Conn.Errors.Count > 0 Then
        ' rollback transaction in case of any transaction error
        g_obj_Conn.RollbackTrans
    End If
    MsgBox Err.Description, vbInformation
End Sub

Private Sub f_void_ProcessDynamicButtons(ByRef l_cmd_Button As CommandButton)
    ' functionality of dynamic buttons
    ' one row should be added to tbSTRSeisekiIchiranProfile table
    ' for each of the selected subjects, one row should be added to tbSTRSeisekiSpecialSubjectProfile table
    Dim l_int_Counter As Integer                    ' counter variable for outer loop
    Dim l_int_InnerCounter As Integer               ' counter variable for inner loop
    Dim l_int_SesekiIchiranId As Long            ' to get the latest Seiseki Ichiran Id
    Dim l_int_SpecialSubjectProfileId As Long    ' to get the latest SpecialSubjectProfileId
    Dim l_str_sqlGetId As String                    ' sql string to get the latest Seiseki Ichiran Id
    Dim l_obj_rsGetId As New ADODB.Recordset        ' recordset object to get the latest Seiseki Ichiran Id
    Dim l_str_sqlInsert As String                   ' insert statement
    Dim l_obj_rsTableMapping As New ADODB.Recordset ' recordset object to get the Seiseki Ichiran Id from tableidmapping table
    Dim l_str_sqlTableMapping As String             ' qsl string to get the Seiseki Ichiran Id from tableidmapping table
    Dim l_int_SelectedTeiki() As Integer            ' array to store the selected teiki rows
    Dim l_int_SelectedKakushu() As Integer          ' array to store the selected Kakushu rows
    Dim l_int_TeikiCounter As Integer               ' teiki counter
    Dim l_int_KakushuCounter As Integer             ' kakushu counter
    Dim l_int_SrNo As Integer                       ' serial number
    Dim l_str_SubjectNames As String                ' store the selected subject names
    Dim l_str_SpecialSubjectIdList                  ' List of special subject profile ids
    Dim l_bln_TeikiSelected As Boolean
    Dim l_bln_KakushuSelected As Boolean
    
    On Error GoTo ErrorHandler
    
    l_str_SpecialSubjectIdList = ""                 ' to store the list of special subject profile id's getting & _
                                                        added for the selected dynamic button

    ' check whether any subjcets are selected in the teiki grid
    If txtReportId.Text = "" Then
        Exit Sub
    Else
        If Trim(txtReportId.Text) = "" Then
            Exit Sub
        End If
    End If
    If l_cmd_Button.BackColor = prvlBtnBCR_NeedSub Then
        For l_int_Counter = 1 To vsfTeiki.Rows - 1
            If vsfTeiki.IsSelected(l_int_Counter) = True Then
                l_bln_TeikiSelected = True
                Exit For
            End If
        Next l_int_Counter
    
        ' if no subjects are selected in both the grids then display error
        If Not l_bln_TeikiSelected And Not l_bln_KakushuSelected Then
    '        lblError.Caption = LoadResString(4078)
            lblError.Caption = "科目を選択してください。"
            Exit Sub
        End If
    End If

    lblError.Caption = ""

    ' get the latest SeisekiIchiranId
    l_str_sqlGetId = "SELECT ISNULL( Max( iSeisekiIchiranId ) + 1 , -1 ) as iSeisekiIchiranId FROM tbSTESeisekiIchiranProfile"
    l_obj_rsGetId.Open l_str_sqlGetId, g_obj_Conn

    If l_obj_rsGetId.Fields("iSeisekiIchiranId").Value <> -1 Then
        l_int_SesekiIchiranId = l_obj_rsGetId.Fields("iSeisekiIchiranId").Value
    Else
        ' no records - extract fiorst id from tableidmapping table
        l_str_sqlTableMapping = "Select iTableCounterIdMapping from tbSTETableIdMapping where vTableName = 'tbSTESeisekiIchiranProfile'"
        l_obj_rsTableMapping.Open l_str_sqlTableMapping, g_obj_Conn
        If Not l_obj_rsTableMapping.EOF Then
            l_int_SesekiIchiranId = l_obj_rsTableMapping.Fields("iTableCounterIdMapping").Value
        Else
            ' no records - initialize to 1
            l_int_SesekiIchiranId = 1
        End If
        l_obj_rsTableMapping.Close
        Set l_obj_rsTableMapping = Nothing
    End If

    l_obj_rsGetId.Close
    Set l_obj_rsGetId = Nothing

    ' get the latest SeisekiIchiranId
    l_str_sqlGetId = "SELECT ISNULL( Max(iSerialNo) + 1 , 1 ) as iSerialNo FROM tbSTESeisekiIchiranProfile"
    l_str_sqlGetId = l_str_sqlGetId & " WHERE iReportNo = " & txtReportId.Text
    l_obj_rsGetId.Open l_str_sqlGetId, g_obj_Conn

    l_int_SrNo = l_obj_rsGetId.Fields("iSerialNo").Value
    If Trim(Me.txtSrNo.Text) <> "" Then
        If CInt(Me.txtSrNo.Text) < l_int_SrNo Then
            '入力シリアル番号より大きいシリアル番号のレコードを自レコードのシリアル番号＋１にて更新
            l_str_sqlGetId = "Update tbSTESeisekiIchiranProfile SET iSerialNo = iSerialNo + 1 "
            l_str_sqlGetId = l_str_sqlGetId & " WHERE iReportNo = " & txtReportId.Text
            l_str_sqlGetId = l_str_sqlGetId & " AND iSerialNo >= " & Me.txtSrNo.Text
            g_obj_Conn.Execute l_str_sqlGetId
            l_int_SrNo = CInt(Me.txtSrNo.Text)
        End If
    End If

    l_obj_rsGetId.Close
    Set l_obj_rsGetId = Nothing
'
'    ' get the latest SpecialSubjectProfileId
'    l_str_sqlGetId = "SELECT iSpecialSubjectProfileId FROM tbSTESeisekiSpecialSubjectProfile"
'    l_obj_rsGetId.Open l_str_sqlGetId, g_obj_Conn
'
'    If Not l_obj_rsGetId.EOF Then
'        ' get the last SpecialSubjectProfileId and add 1 to that
'        l_obj_rsGetId.MoveLast
'        l_int_SpecialSubjectProfileId = l_obj_rsGetId.Fields("iSpecialSubjectProfileId").Value + 1
'    Else
'        ' no records - extract fiorst id from tableidmapping table
'        l_str_sqlTableMapping = "Select iTableCounterIdMapping from tbSTETableIdMapping where vTableName = 'tbSTRSeisekiSpecialSubjectProfile'"
'        l_obj_rsTableMapping.Open l_str_sqlTableMapping, g_obj_Conn
'        If Not l_obj_rsTableMapping.EOF Then
'            l_int_SpecialSubjectProfileId = l_obj_rsTableMapping.Fields("iTableCounterIdMapping").Value
'        Else
'            ' no records - initialize to 1
'            l_int_SpecialSubjectProfileId = 1
'        End If
'        l_obj_rsTableMapping.Close
'        Set l_obj_rsTableMapping = Nothing
'    End If
'
'    l_obj_rsGetId.Close
'    Set l_obj_rsGetId = Nothing
Dim bRtn As Boolean
    bRtn = getNewId("tbSTESeisekiSpecialSubjectProfile", "iSpecialSubjectProfileId", l_int_SpecialSubjectProfileId)

    g_obj_Conn.BeginTrans
                
    ' special profile id will go as null
    l_str_sqlInsert = " INSERT INTO tbSTESeisekiIchiranProfile VALUES("
    ' primary key
    l_str_sqlInsert = l_str_sqlInsert & l_int_SesekiIchiranId & ","
    ' report id
    If Len(txtReportId.Text) > 0 Then
        l_str_sqlInsert = l_str_sqlInsert & txtReportId.Text & ","
    Else    ' insert default value as 1
        l_str_sqlInsert = l_str_sqlInsert & "1,"
    End If
    ' serial id 画面入力ｏｒ最大、入力＞最大は最大にする
    'とりあえず、最大のみ
    l_str_sqlInsert = l_str_sqlInsert & l_int_SrNo & ","
    ' special id
    l_str_sqlInsert = l_str_sqlInsert & l_cmd_Button.Tag & ","
    ' subject profile id
    l_str_sqlInsert = l_str_sqlInsert & "null,"
    ' subject question profile id
    l_str_sqlInsert = l_str_sqlInsert & "null,"

    l_str_sqlInsert = l_str_sqlInsert & "'" & Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
    
    g_obj_Conn.Execute l_str_sqlInsert
            
    ' initialize the counter variables
    l_int_TeikiCounter = 0
    l_int_KakushuCounter = 0

'集計しない（Subjectを必要としない）項目は以降処理しない
    If l_cmd_Button.BackColor = prvlBtnBCR_NeedSub Then
        ' loop thru the teiki grid and store the selected rows into array
        ' this is done becasue later while programatically fetching the values for inserting,
        ' the orininally selected rows gets lost - so store it initially and then laster retrieve it
        With vsfTeiki

        For l_int_Counter = 1 To .Rows - 1
            If .IsSelected(l_int_Counter) Then
                ReDim Preserve l_int_SelectedTeiki(l_int_TeikiCounter)
                l_int_SelectedTeiki(l_int_TeikiCounter) = l_int_Counter
                l_int_TeikiCounter = l_int_TeikiCounter + 1
            End If
        Next l_int_Counter

        End With

        l_str_SubjectNames = ""

        ' loop thru the teiki grid and insert data of selected rows into tbSTRSeisekiSpecialSubjectProfile table
        With vsfTeiki

        For l_int_Counter = 1 To .Rows - 1
            For l_int_InnerCounter = 0 To l_int_TeikiCounter - 1
                If l_int_Counter = l_int_SelectedTeiki(l_int_InnerCounter) Then
                    .Row = l_int_Counter
                    ' special profile id will go as null
                    l_str_sqlInsert = " INSERT INTO tbSTESeisekiSpecialSubjectProfile VALUES(" & _
                        l_int_SpecialSubjectProfileId & ","     ' primary key
                    .Col = 2    ' subject grade profile id
                    l_str_sqlInsert = l_str_sqlInsert & (.Text) & "," & _
                        l_cmd_Button.Tag & "," & _
                        l_int_SesekiIchiranId & ",'" & _
                        Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"

                    g_obj_Conn.Execute l_str_sqlInsert

                    .Col = 1
                    l_str_SubjectNames = l_str_SubjectNames & .Text & ","
                    l_str_SpecialSubjectIdList = l_str_SpecialSubjectIdList & l_int_SpecialSubjectProfileId & ","
                    
                    l_int_SpecialSubjectProfileId = l_int_SpecialSubjectProfileId + 1
                    Exit For
                End If
            Next l_int_InnerCounter
        Next l_int_Counter

        End With

    End If
    
    g_obj_Conn.CommitTrans
    
    ' add data to grid G2
    vsfG2.Rows = vsfG2.Rows + 1                 ' add a row
    vsfG2.Row = vsfG2.Rows - 1                  ' set the row
    vsfG2.Col = 0
    vsfG2.Text = vsfG2.Rows - 1                 ' serial number
    vsfG2.Col = 1
    vsfG2.Text = l_cmd_Button.Caption           ' name of dynamic button
    vsfG2.Col = 2
    vsfG2.Text = l_str_SubjectNames             ' selected subject names -  only for forward button
    vsfG2.Col = 3
    vsfG2.Text = l_int_SesekiIchiranId          ' list of ichiran rows added for the selected button
    vsfG2.Col = 4
    vsfG2.Text = l_str_SpecialSubjectIdList     ' list of specialsubject rows added for the selected button - only for dynamic buttons
    
    Exit Sub
ErrorHandler:
    If g_obj_Conn.Errors.Count > 0 Then
        ' rollback transaction in case of any transaction error
        g_obj_Conn.RollbackTrans
    End If
    MsgBox Err.Description, vbInformation
End Sub

Private Sub udCmdBtnPage_DownClick()

    If CInt(lblCmdBtnPage.Caption) = 1 Then Exit Sub

    lblCmdBtnPage.Caption = Trim(str(CInt(lblCmdBtnPage.Caption) - 1))

    Call f_void_ShowButtons

End Sub

Private Sub udCmdBtnPage_UpClick()

    If CInt(lblCmdBtnPage.Caption) = CInt(f_cmd_ButtonArray.Count / prviShowButtons) Then Exit Sub

    lblCmdBtnPage.Caption = Trim(str(CInt(lblCmdBtnPage.Caption) + 1))

    Call f_void_ShowButtons

End Sub

Private Sub vsfG2_Click()
    Dim l_str_SubjectNames() As String      ' to get the list of subjects from grid G2
    Dim l_int_Counter As Integer            ' counter
    Dim l_int_SrNo As Integer               ' to store serial number
    
        
    With vsfG3
    '表示を消す
    .Rows = 1

    ' check whether subjects for the selected row in G2 is already added to G3 or not
    ' if already added, then exit the procedure
'    For l_int_Counter = 1 To .Rows - 1
'        .Row = l_int_Counter
'        .Col = 2
'        vsfG2.Col = 3
'        If .Text = vsfG2.Text Then Exit Sub
'    Next l_int_Counter
      
    vsfG2.Col = 2
    If Len(vsfG2.Text) > 0 Then
        l_str_SubjectNames = Split(vsfG2.Text, ",")
        
        ' loop thru the subject array and add each subject into grid G3
        For l_int_Counter = 0 To UBound(l_str_SubjectNames)
            .Rows = .Rows + 1                           ' add a row
            .Row = .Rows - 1                            ' set the row
            .Col = 0                                    ' serial no
            .Text = .Rows - 1
            .Col = 1                                    ' subject name
            .Text = l_str_SubjectNames(l_int_Counter)
            vsfG2.Col = 3
            .Col = 2
            .Text = vsfG2.Text                          ' store the seiseki ichiran id
            l_int_SrNo = l_int_SrNo + 1                 ' increment the counter
        Next l_int_Counter
    End If
    
    End With
End Sub

Private Sub f_void_DeleteRecords()
    ' deleted selected records from G2 and related records from G3
    ' also delete from tbSTRSeisekiIchiranProfile tale and related records from tbSTRSeisekiSpecialSubjectProfile table
    Dim l_int_Counter As Integer                        ' counter for grid G2
    Dim l_int_counter1 As Integer                       ' counter for grid G3
    Dim l_int_InitialRows As Integer                    ' initial no of rows in G3
    Dim l_int_InitialG2Rows As Integer                  ' initial no of rows in G2
    Dim l_str_sqlDeleteSpecialSubjectId As String       ' sql string to form the delete statement for tbSTESeisekiSpecialSubjectProfile table
    Dim l_str_sqlDeleteSeisekiIchiranId As String       ' sql string to form the delete statement for tbSTESeisekiIchiranProfile table
    Dim l_str_sqlUpdateSerialNo As String               ' update statement to set the iSerialNo field in tbSTRSeisekiIchiranProfile table, after a deletion happens
    Dim l_str_sqlSeisekiIchiran As String               ' to get the iSeisekiIchiranId
    Dim l_obj_rsSeisekiIchiran As New ADODB.Recordset   ' recordset varibale for the above
    Dim l_str_sqlSerialNo As String                     ' to get the iSerialNo of the record being deleted
    Dim l_obj_rsSerialno As New ADODB.Recordset         ' recordset varibale for the above
    Dim l_int_SerialNo As Integer                       ' to store the iSerialNo
    Dim l_int_G2Counter As Integer                      ' counter for the rgid G2
    Dim l_int_SelectedG2() As Integer                   ' to store selected rows in grid G2
    Dim l_int_InnerCounter As Integer                   ' to loop thru the selected array
    Dim l_int_DeleteCounter As Integer                  ' to update the selected rows array after a row is deleted
    Dim l_int_UpperLimit As Integer                     ' to set the upper limit of the selected array, after a row is deleted - next search will start from here onwards
    
    On Error GoTo ErrorHandler
    
    ' initialize the counter variables
    l_int_G2Counter = 0
    
    ' loop thru the G2 grid and store the selected rows into array
    ' this is done becasue later while programatically fetching the values for inserting,
    ' the orininally selected rows gets lost - so store it initially and then laster retrieve it
    With vsfG2
    
    For l_int_Counter = 1 To .Rows - 1
        If .IsSelected(l_int_Counter) Then
            ReDim Preserve l_int_SelectedG2(l_int_G2Counter)
            l_int_SelectedG2(l_int_G2Counter) = l_int_Counter
            l_int_G2Counter = l_int_G2Counter + 1
        End If
    Next l_int_Counter
                           
    ' loop thru the G2 grid and delete data of selected rows from tbSTRSeisekiIchiranProfile table
    l_int_InitialG2Rows = .Rows - 1
                    
    For l_int_Counter = 1 To .Rows - 1      ' outer loop for the grid
        For l_int_InnerCounter = 0 To l_int_G2Counter - 1   ' inner loop for the selected array
            If l_int_InnerCounter >= l_int_UpperLimit Then  ' check the upper limit of selected array
                If l_int_Counter = l_int_SelectedG2(l_int_InnerCounter) Then    ' check whether current row was selected or not
                    If l_int_Counter > l_int_InitialG2Rows Then Exit Sub
                    
                    .Row = l_int_Counter
        
                    g_obj_Conn.BeginTrans
                        
                    ' if  specialprofileid field id not null, then corresponding record exist in tbstrseisekispecialsubjectprofile table - delete it
                    .Col = 4
                    If Len(.Text) > 0 Then
                        l_str_sqlDeleteSpecialSubjectId = "DELETE FROM tbSTESeisekiSpecialSubjectProfile" & _
                            " WHERE iSpecialSubjectProfileId IN(" & .Text & ")"
                        g_obj_Conn.Execute l_str_sqlDeleteSpecialSubjectId
                    End If
                    ' delete records form tbSTRSeisekiIchiranProfile table
                    .Col = 3
                    If Len(.Text) > 0 Then
                        ' get serial number of the row to be deleted
                        l_str_sqlSerialNo = "SELECT iSerialNo FROM tbSTESeisekiIchiranProfile" & _
                            " WHERE iSeisekiIchiranId = " & .Text
                        l_obj_rsSerialno.Open l_str_sqlSerialNo, g_obj_Conn
                        If Not l_obj_rsSerialno.EOF Then
                            l_int_SerialNo = l_obj_rsSerialno.Fields("iSerialNo").Value
                        End If
                        l_obj_rsSerialno.Close
                        Set l_obj_rsSerialno = Nothing
                        
                        l_str_sqlDeleteSeisekiIchiranId = "DELETE FROM tbSTESeisekiIchiranProfile" & _
                            " WHERE iSeisekiIchiranId = " & .Text
                        g_obj_Conn.Execute l_str_sqlDeleteSeisekiIchiranId
                    End If
                    
                    ' update serial no of the remaining rows
                    l_str_sqlSeisekiIchiran = "SELECT iSeisekiIchiranId , iSerialNo FROM tbSTESeisekiIchiranProfile"
                    l_str_sqlSeisekiIchiran = l_str_sqlSeisekiIchiran & " WHERE iReportNo = " & txtReportId.Text
                    l_str_sqlSeisekiIchiran = l_str_sqlSeisekiIchiran & " AND iSerialNo > " & l_int_SerialNo
                    l_str_sqlSeisekiIchiran = l_str_sqlSeisekiIchiran & " ORDER BY iSerialNo"
                    l_obj_rsSeisekiIchiran.Open l_str_sqlSeisekiIchiran, g_obj_Conn
                    
                    Do While Not l_obj_rsSeisekiIchiran.EOF
                        l_str_sqlUpdateSerialNo = "UPDATE tbSTESeisekiIchiranProfile"
                        l_str_sqlUpdateSerialNo = l_str_sqlUpdateSerialNo & " SET iSerialNo = " & l_obj_rsSeisekiIchiran.Fields("iSerialNo").Value - 1 & ","
                        l_str_sqlUpdateSerialNo = l_str_sqlUpdateSerialNo & " dtUpdate ='" & Format(Date, "MM/DD/YYYY") & "'"
                        l_str_sqlUpdateSerialNo = l_str_sqlUpdateSerialNo & " WHERE iSeisekiIchiranId = " & l_obj_rsSeisekiIchiran.Fields("iSeisekiIchiranId").Value
                        g_obj_Conn.Execute l_str_sqlUpdateSerialNo
                        l_obj_rsSeisekiIchiran.MoveNext
                   Loop
                    l_obj_rsSeisekiIchiran.Close
                    Set l_obj_rsSeisekiIchiran = Nothing
                    
                    g_obj_Conn.CommitTrans
                                    
                    ' remove the corresponding rows from G3
                    l_int_InitialRows = vsfG3.Rows - 1
    
                    .Col = 3
                    For l_int_counter1 = 1 To vsfG3.Rows - 1
                        If l_int_counter1 > l_int_InitialRows Then Exit For
                        vsfG3.Row = l_int_counter1
                        vsfG3.Col = 2
                        If .Text = vsfG3.Text Then
                            vsfG3.RemoveItem l_int_counter1
                            l_int_counter1 = l_int_counter1 - 1
                            l_int_InitialRows = l_int_InitialRows - 1
                        End If
                    Next l_int_counter1
                    
                    ' remove the selected rows from G2
                    .RemoveItem .Row
                    l_int_UpperLimit = l_int_InnerCounter + 1   ' after this, search from this element onwards in the l_int_SelectedG2 array
                    For l_int_DeleteCounter = l_int_UpperLimit To l_int_G2Counter - 1
                        l_int_SelectedG2(l_int_DeleteCounter) = l_int_SelectedG2(l_int_DeleteCounter) - 1
                    Next l_int_DeleteCounter
                    l_int_Counter = l_int_Counter - 1
                    
                    Exit For
                End If
            End If
        Next l_int_InnerCounter
    Next l_int_Counter
        
    End With
    
    ' set the serial numbers in order after the deletions
'    Call g_void_RefreshSerialNo(vsfG2)
'    Call g_void_RefreshSerialNo(vsfG3)
    Exit Sub
ErrorHandler:
    If g_obj_Conn.Errors.Count > 0 Then
        g_obj_Conn.RollbackTrans
    End If
    MsgBox Err.Description, vbInformation
End Sub


Private Sub f_void_PopulateG2(ByVal l_int_ReportId As Long)
    ' retireve records for the input report no and then display it in G2 and G3 accordingly
    Dim l_str_sqlSeisekiIchiran As String               ' sql string to get the seisekiichiran id
    Dim l_obj_rsSeisekiIchiran As New ADODB.Recordset   ' recordset object to get the seisekiichiran id
    Dim l_str_sqlSpecialSubject As String               ' sql string to get the specialsubject profile ids
    Dim l_obj_rsSpecialSubject As New ADODB.Recordset   ' recordset object to get the specialsubject profile ids
    Dim l_str_sqlSubjectName As String                  ' sql string to get the subject names
    Dim l_obj_rsSubjectName As New ADODB.Recordset      ' recordset object to get the sunject names
    Dim l_obj_rsSubjectName2 As New ADODB.Recordset      ' recordset object to get the sunject names
    Dim l_int_SeisekiIchiranId As Integer
    Dim l_int_GradeProfileId As Integer
    Dim l_str_sqlButtonName As String                   ' sql string to get the button name
    Dim l_obj_rsButtonName As New ADODB.Recordset       ' recordset object to get the button name
    Dim l_str_SubjectList As String                     ' to store the subjects
    Dim l_str_SpecialSubjectList As String              ' to store the special subject id's
    Dim l_int_Counter As Integer
        
    ' clear the grids
'    vsfTeiki.Rows = 1
    vsfG3.Rows = 1
    vsfG2.Rows = 1

    l_str_sqlSeisekiIchiran = "SELECT * FROM tbSTESeisekiIchiranProfile" & _
        " WHERE iReportNo = " & l_int_ReportId & _
        " ORDER BY iSerialNo"
    l_obj_rsSeisekiIchiran.Open l_str_sqlSeisekiIchiran, g_obj_Conn
    
    If l_obj_rsSeisekiIchiran.EOF Then
'        lblError.Caption = LoadResString(4213)
        fMainForm.mnuPrint.Enabled = False
        fMainForm.Toolbar1.Buttons("Print").Enabled = False
        Exit Sub
    End If
    
'    'disable the buttons
'    For l_int_Counter = 0 To f_cmd_ButtonArray.Count - 1
'        f_cmd_ButtonArray(l_int_Counter).Enabled = False
'    Next l_int_Counter

    fMainForm.mnuPrint.Enabled = True
    fMainForm.Toolbar1.Buttons("Print").Enabled = True

    lblError.Caption = ""

    With vsfG2

    Do While Not l_obj_rsSeisekiIchiran.EOF

        l_int_SeisekiIchiranId = l_obj_rsSeisekiIchiran.Fields("iSeisekiIchiranId").Value

        If Len(l_obj_rsSeisekiIchiran.Fields("iSpecialProfileId").Value) > 0 Then
            l_str_SubjectList = ""
            l_str_SpecialSubjectList = ""

            l_str_sqlSpecialSubject = "SELECT iSpecialSubjectProfileId, iSubjectProfileId FROM tbSTESeisekiSpecialSubjectProfile" & _
                " WHERE iSeisekiIchiranId = " & l_int_SeisekiIchiranId
            l_obj_rsSpecialSubject.Open l_str_sqlSpecialSubject, g_obj_Conn
            Do While Not l_obj_rsSpecialSubject.EOF
                l_str_sqlSubjectName = "SELECT vSubjectName FROM tbSTESubjectProfile"
                l_str_sqlSubjectName = l_str_sqlSubjectName & " WHERE iSubjectProfileId = " & l_obj_rsSpecialSubject.Fields("iSubjectProfileId").Value
                l_obj_rsSubjectName2.Open l_str_sqlSubjectName, g_obj_Conn

                If Not l_obj_rsSubjectName2.EOF Then

'                    With vsfG3

'                    .Rows = .Rows + 1
'                    .Row = .Rows - 1
'                    .Col = 0
'                    .Text = .Rows - 1
'                    .Col = 1
'                    .Text = l_obj_rsSubjectName2.Fields("vSubjectName").Value
                    l_str_SubjectList = l_str_SubjectList & l_obj_rsSubjectName2.Fields("vSubjectName").Value & ","
'                    .Col = 2
'                    .Text = l_int_SeisekiIchiranId
                    l_str_SpecialSubjectList = l_str_SpecialSubjectList & l_obj_rsSpecialSubject.Fields("iSpecialSubjectProfileId").Value & ","
'
'                    End With
'
                End If
                l_obj_rsSubjectName2.Close
                Set l_obj_rsSubjectName2 = Nothing

                l_obj_rsSpecialSubject.MoveNext
            Loop

            If Len(l_str_SubjectList) > 0 Then
                l_str_SubjectList = Left(l_str_SubjectList, Len(l_str_SubjectList) - 1)
            End If

            If Len(l_str_SpecialSubjectList) > 0 Then
                l_str_SpecialSubjectList = Left(l_str_SpecialSubjectList, Len(l_str_SpecialSubjectList) - 1)
            End If

            l_obj_rsSpecialSubject.Close
            Set l_obj_rsSpecialSubject = Nothing
        End If
        
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .Text = .Rows - 1
        If Len(l_obj_rsSeisekiIchiran.Fields("iSpecialProfileId").Value) > 0 Then
            l_str_sqlButtonName = "SELECT vButtonName FROM tbSTESeisekiSpecialProfile" & _
                " WHERE iSpecialProfileId = " & l_obj_rsSeisekiIchiran.Fields("iSpecialProfileId").Value
            l_obj_rsButtonName.Open l_str_sqlButtonName, g_obj_Conn
            If Not l_obj_rsButtonName.EOF Then
                .Col = 1
                .Text = l_obj_rsButtonName.Fields("vButtonName").Value
            End If
            l_obj_rsButtonName.Close
            Set l_obj_rsButtonName = Nothing

            .Col = 2
            If Len(l_str_SubjectList) > 0 Then
                .Text = l_str_SubjectList
            End If

            .Col = 3
            .Text = l_int_SeisekiIchiranId

            .Col = 4
            If Len(l_str_SpecialSubjectList) > 0 Then
                .Text = l_str_SpecialSubjectList
            End If
        Else
            .Col = 1
            If IsNull(l_obj_rsSeisekiIchiran.Fields("iSubjectQuestionId").Value) Then
                l_str_sqlSubjectName = "SELECT vSubjectName FROM tbSTESubjectProfile"
                l_str_sqlSubjectName = l_str_sqlSubjectName & " WHERE iSubjectProfileId = " & l_obj_rsSeisekiIchiran.Fields("iSubjectProfileId").Value
            Else
                l_str_sqlSubjectName = "SELECT vQuestionName as vSubjectName FROM tbSTESubjectQuestionProfile"
                l_str_sqlSubjectName = l_str_sqlSubjectName & " WHERE iSubjectQuestionId = " & l_obj_rsSeisekiIchiran.Fields("iSubjectQuestionId").Value
            End If
            l_obj_rsSubjectName.Open l_str_sqlSubjectName, g_obj_Conn
            
            If Not l_obj_rsSubjectName.EOF Then
                .Text = l_obj_rsSubjectName.Fields("vSubjectName").Value
            End If
            l_obj_rsSubjectName.Close
            Set l_obj_rsSubjectName = Nothing
            
            .Col = 2
            .Text = ""
            
            .Col = 3
            .Text = l_int_SeisekiIchiranId
            
            .Col = 4
            .Text = ""
        End If

        l_obj_rsSeisekiIchiran.MoveNext
    Loop

    End With

    l_obj_rsSeisekiIchiran.Close
    Set l_obj_rsSeisekiIchiran = Nothing
End Sub
'
'Public Sub f_void_Print()
'    ' procedure for printing the report
'    Dim l_int_Counter As Integer
'    Dim l_str_ReportParams As String
'    Dim l_lng_RptId As Long
'    Dim l_int_PrinterId As Long
'
'    On Error GoTo ErrorHandler
'
'    ' oreder by parameter 1
'    l_str_ReportParams = ";" & LoadResString(Frame1.Tag) & "-" & Frame1.Caption & "="
'
'    For l_int_Counter = 0 To optChouhyouShurui.Count - 1
'        If optChouhyouShurui(l_int_Counter).Value = True Then
'            l_str_ReportParams = l_str_ReportParams & optChouhyouShurui(l_int_Counter).Caption
'            Exit For
'        End If
'    Next l_int_Counter
'
'    ' oreder by parameter 2
'    l_str_ReportParams = l_str_ReportParams & ";" & LoadResString(Frame2.Tag) & "-" & Frame2.Caption & "="
'
'    For l_int_Counter = 0 To optShutsuryokuJun.Count - 1
'        If optShutsuryokuJun(l_int_Counter).Value = True Then
'            l_str_ReportParams = l_str_ReportParams & optShutsuryokuJun(l_int_Counter).Caption
'            Exit For
'        End If
'    Next l_int_Counter
'
'    ' report number
'    l_str_ReportParams = l_str_ReportParams & ";" & LoadResString(lblReportId.Tag) & "=" & _
'        txtReportId.Text
'
'    l_str_ReportParams = "'" & l_str_ReportParams & "'"
'
'    l_lng_RptId = g_int_SeisekiReportId & "105"
'    l_int_PrinterId = 1
'    Call gp_void_InsertReportData(l_lng_RptId, l_int_PrinterId, l_str_ReportParams)
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation
'End Sub

Private Sub f_void_InitGrid(ByRef l_grd_MyGrid As VSFlexGrid) 'Initialize the grid
    On Error GoTo ErrorHandler
    
    With l_grd_MyGrid
        .FixedRows = 1
        .FixedCols = 1
        .TextStyleFixed = flexTextFlat
        .CellTextStyle = "0"
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .BackColor = &HFFFFFF
        .BackColorBkg = &HFFFFFF
        .BackColorFixed = &H8000000F
'        .BackColorSel = g_str_GridSelBackColor
'        .ForeColorSel = g_str_GridSelForeColor
        .Font.Bold = False
        .ForeColorFixed = &H80000008
        .ForeColor = &H800000
        .GridColor = &H808080
        .Rows = 1
    End With
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(4312)
End Sub

Private Sub f_void_SetReportNo()

    Dim l_int_ReportNo As Long            ' to get the latest Seiseki Ichiran Id
    Dim l_str_sqlGetId As String                    ' sql string to get the latest Seiseki Ichiran Id
    Dim l_obj_rsGetId As New ADODB.Recordset        ' recordset object to get the latest Seiseki Ichiran Id

    On Error GoTo ErrorHandler

    ' get the latest SeisekiIchiranId
    l_str_sqlGetId = "SELECT ISNULL( Max( iReportNo ) + 1 , 1 ) as iReportNo FROM tbSTESeisekiIchiranProfile"
    l_obj_rsGetId.Open l_str_sqlGetId, g_obj_Conn

    l_int_ReportNo = l_obj_rsGetId.Fields("iReportNo").Value

    l_obj_rsGetId.Close
    Set l_obj_rsGetId = Nothing

    Me.txtReportId.Text = Trim(str(l_int_ReportNo))

Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(4312)
End Sub
