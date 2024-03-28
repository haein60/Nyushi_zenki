VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmChooseiScore1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   10110
   ClientLeft      =   1275
   ClientTop       =   900
   ClientWidth     =   13230
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmChooseiScore1.frx":0000
   ScaleHeight     =   10110
   ScaleWidth      =   13230
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "1071"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "1060"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   255
      TabIndex        =   0
      Top             =   4935
      Width           =   1695
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfselectRawScore 
      Height          =   3615
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
      _cx             =   7435
      _cy             =   6376
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   134217736
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   134217736
      BackColorSel    =   8388608
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmChooseiScore1.frx":3AD3
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
      Editable        =   2
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
   Begin VB.Label lblErrorDetails 
      BackStyle       =   0  '透明
      Caption         =   "Error Details"
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
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   9735
   End
End
Attribute VB_Name = "frmChooseiScore1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmChooseiScore
'Author         :   Dileep Cherian
'Created On     :   18/9/01
'Description    :   This form will be used as a mechanism to insert choosei score for first Examination.
'Reference      :   FunctionalSpecs OF CHOSEISCORE.doc(Ver 1.1)
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History   -   04/04/2002  -   Dileep Cherian
'User should be able to resize the coulmns, incase part of data is not visible in the normal display
'On pressing ente after editing a column value, the focus should move to the next row (same column)
'**************************************************************************************************
' Modification History
' Changes by : Mahesh (10-11-2001)
' Changed to allow selection of raw score in a grid fashion instead of combo selection
' Ammendments - NyushiChangesSummary.doc ver 1.0
' Changes by : Mahesh (17-5-2002)
' Chenged to add room checkbox, on clicking of which records will be added to grid for every room in rommMaster
' Changed by : Mahesh (25/5/2002)
' Changed to add list box for rooms. instead of all rooms from room master, only selected rooms data will be added to grod.

' 2004/04/21
' １次試験、科目別用の調整点入力画面
' 全受験生のScore（tbSTEScoreProfile）に対して入力された調整点で(iChooseiScoreを)更新する
' tbSTEScoreDetailはさわらない
' また、調整点入力の履歴としてtbSTEChoseiJokenを登録・更新する

Option Explicit

' database related variables
Dim m_obj_Rst As New ADODB.Recordset    ' recordset object
Dim m_str_SQL As String                 ' to store the SQL string
Dim m_int_SelectedSubject As Long    ' to store the selected subject from the subject combo
'Dim m_int_NoOfErr As Long            ' to keep track of no of errors
Dim m_int_NoOfConditions As Long     ' to track the no of conditions
Public m_int_ChoseiJoken As Long          ' to diff b/w Grace Score and Suisen Score
Dim m_bln_OnceEntered As Boolean        ' boolean stores whether the conditions have been entered once. if so,user hav to clear off first

Private Sub cmdOK_Click()

    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdSubmit_Click()

' 入力情報によってtbSTEScoreProfileを更新する
' また、tbSTEChoseiJokenに履歴として調整情報を登録する
' ただし、2004/04/21現在、登録した情報を表示する画面はなし
' 毎回入力されたデータにて調整点として更新する

Dim sSQL As String
Dim lRow As Long
Dim lRtn As Long
Dim dCScore As Double
Dim l_Bln_RecordsUpdated As Boolean

    l_Bln_RecordsUpdated = False

'トランザクション開始
    g_obj_Conn.BeginTrans

'入力情報をGridから取得するループStart
    For lRow = 1 To vsfselectRawScore.Rows - 1
        If Trim(vsfselectRawScore.TextMatrix(lRow, 2)) <> "" And IsNumeric(vsfselectRawScore.TextMatrix(lRow, 2)) Then
        '数値が入力されている
            dCScore = CDbl(vsfselectRawScore.TextMatrix(lRow, 2))
            lRtn = gflDelChoseiJoken(g_int_CurrentNendo, CInt(vsfselectRawScore.TextMatrix(lRow, 0)), 1)
            lRtn = gflInsChoseiJoken(g_int_CurrentNendo, CInt(vsfselectRawScore.TextMatrix(lRow, 0)), 1, -1, -1, "-1", -1, dCScore)
            If lRtn <> 0 Then GoTo ErrorHandler
            sSQL = "update tbSTEScoreProfile "
            sSQL = sSQL & " set fChoseiScore = " & vsfselectRawScore.TextMatrix(lRow, 2)
            sSQL = sSQL & ", dtUpdate='" & Format(Date, "HH:MM:SS MM/DD/YYYY") & "'"
            sSQL = sSQL & " where exists ( select 1 from tbSTEExamineeProfile as ep "
            sSQL = sSQL & " where ep.iExamineeProfileId = tbSTEScoreProfile.iExamineeProfileId "
            sSQL = sSQL & " and iNendo = " & str(g_int_CurrentNendo) & " ) "
            sSQL = sSQL & " and iSubjectProfileID = " & vsfselectRawScore.TextMatrix(lRow, 0)
            g_obj_Conn.Execute sSQL
            l_Bln_RecordsUpdated = True
        End If
    Next
'入力情報をGridから取得するループEnd

'トランザクション終了
    g_obj_Conn.CommitTrans

    If l_Bln_RecordsUpdated Then
        lblErrorDetails.Caption = LoadResString(2404)
    Else
        lblErrorDetails.Caption = LoadResString(2427)
    End If
    lblErrorDetails.Visible = True

    Exit Sub
ErrorHandler:
'エラーなのでトランザクションをロールバックする
    g_obj_Conn.RollbackTrans
    lblErrorDetails.Visible = True
    lblErrorDetails.Caption = LoadResString(1761) & _
        vbCrLf & LoadResString(1125)
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

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    LoadResStrings Me
    If m_int_ChoseiJoken = 1 Then
        Me.Caption = LoadResString(1012)
    Else
        Me.Caption = "科目別補正点入力"
    End If
    Call g_void_SetFontProperties(Me)     ' set the font properties
    m_int_NoOfConditions = 0    ' initialise the no of conditions
    ' select all subjects that come under the selected exam type
    m_str_SQL = "SELECT sp.iSubjectProfileId,sp.vSubjectName "
    m_str_SQL = m_str_SQL & " , isnull( STR( fChoseiScore , 5 , 1 ) , '' ) as fChoseiScore "
    m_str_SQL = m_str_SQL & " FROM tbSTESubjectProfile as sp "
    m_str_SQL = m_str_SQL & " LEFT OUTER JOIN tbSTEChoseiJoken as cj On cj.iSubjectProfileID = sp.iSubjectProfileID "
    m_str_SQL = m_str_SQL & "                                        AND cj.iNendo = " & g_int_CurrentNendo

    ' changed on 14/05/02 to incorporate choosei for Hyotei also
    m_str_SQL = m_str_SQL & " WHERE sp.iExamType = " & g_int_ExamType

    m_str_SQL = m_str_SQL & " ORDER BY sp.iDispOrder "
    Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQL)
'    cmdSubmit.Enabled = False
    cmdOk.Visible = False

    If Not m_obj_Rst.EOF Then
        m_int_SelectedSubject = m_obj_Rst("iSubjectProfileId")
        ' add the subjects to combo box
Dim lRow As Long
        lRow = 1
        Do While Not m_obj_Rst.EOF
            vsfselectRawScore.Rows = lRow + 1
            vsfselectRawScore.TextMatrix(lRow, 0) = m_obj_Rst("iSubjectProfileId")
            vsfselectRawScore.TextMatrix(lRow, 1) = m_obj_Rst("vSubjectName")
            vsfselectRawScore.TextMatrix(lRow, 2) = m_obj_Rst(2)
            m_obj_Rst.MoveNext
            lRow = lRow + 1
        Loop
        vsfselectRawScore.ColWidth(0) = 0
        vsfselectRawScore.ColWidth(1) = 2000
        vsfselectRawScore.ColWidth(2) = 800
        vsfselectRawScore.Row = 1
        vsfselectRawScore.Col = 2
    End If
    
    ' release the object variables
    Set m_obj_Rst = Nothing
    
'    Call f_void_InitGrid            ' reinitialize the grid
'    Call f_void_InitRawScoreGrid    ' reinitialize the rawscore grid
'    Call f_void_LoadRoom            ' Room is a checkbox now
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_InitRawScoreGrid()
    
    With vsfselectRawScore
        .Visible = False
        .BackColor = &HFFFFFF
        .BackColorBkg = &HFFFFFF
        .BackColorFixed = &H8000000F
        .BackColorSel = &H800000
        .FixedCols = 0
        .TextStyleFixed = flexTextFlat
        
        ' change made in com design, arka , 11 apr02
        '.Font.Name = "ＭＳ Ｐゴシック"
        '.Font.Name = "Verdana"
        
        .ForeColorFixed = &H80000008
        .ForeColor = &H800000
        '.CellTextStyle = "0"
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .GridColor = &H808080
        .Visible = True
       
        .Row = 0
        .Col = 0
        .ColWidth(0) = 1200
        .Text = LoadResString(1769)
        .CellAlignment = flexAlignRightBottom
        
        .Col = .Col + 1
        .ColWidth(1) = 1200
        .Text = LoadResString(1770)
        .Editable = flexEDNone
        .Row = .Row + 1
        .Col = 0
        .Text = 0
        .Col = .Col + 1
        .Text = 100
    End With
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_bln_OnceEntered = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call g_void_CloseChildForm
End Sub

Private Sub vsfselectRawScore_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrorHandler
    With vsfselectRawScore
            If Col < .Cols - 1 Then
                .Col = .Col + 1
            ElseIf Col = .Cols - 1 Then
                If Trim(.Text) <> "" Then
                    .Col = 0
                    If Trim(.Text) <> "" Then
                        .Row = .Rows - 1    'Go to last row and if its not blank, add a row
                        .Col = 0
                        If .Text <> "" Then
                            If .Rows < 11 Then
'                                .Rows = .Rows + 1
                                .Row = .Rows - 1
                                .Col = 0
                            End If
                        End If
                    End If
                End If
            End If
    End With
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub
'add,xzg,2009/12/17,S---------
Private Sub vsfselectRawScore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col < 2 Then Cancel = True
End Sub
'add,xzg,2009/12/17,E---------

Private Sub vsfselectRawScore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Clipboard.Clear
    End If
End Sub

Private Sub vsfselectRawScore_Click()
    lblErrorDetails.Caption = ""
    lblErrorDetails.Visible = False
'    If chkRawScore.Value = 1 Then
        vsfselectRawScore.EditCell
'    End If
End Sub

Private Sub vsfselectRawScore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    vsfselectRawScore.Redraw = flexRDDirect
'    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> vbKeyReturn And KeyAscii <> 45 And KeyAscii <> 46 Then '45:- 46:.
'       KeyAscii = 0
'    End If
    If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyEscape Then
        Call NumericPeriodMinusVsfGrd(vsfselectRawScore, KeyAscii)
    End If
End Sub

Private Function f_bln_ValidateRange() As Integer

    Dim l_int_Rows As Long 'total rows in grid
    Dim l_int_Counter As Long ' current row
    Dim l_bln_RetVal As Long  ' return value
    Dim l_int_PrevColVal As Long  'previous col value of same row
    'Dim l_int_PrevRowVal As Integer  ' previous col value of prev row
    ' 0 means all ok
    ' 1 means check box checked but no values entered
    ' 2 means Continuity is missing
    On Error GoTo ErrorHandler
    l_bln_RetVal = 0
    
    l_int_Rows = vsfselectRawScore.Rows
    vsfselectRawScore.Row = 1
    vsfselectRawScore.Col = 0
    
    With vsfselectRawScore
        If .Text = "" Then
            l_bln_RetVal = 1
            f_bln_ValidateRange = l_bln_RetVal
            Exit Function
        End If
        l_int_PrevColVal = vsfselectRawScore.Text
         For l_int_Counter = 1 To .Rows - 1
             .Row = l_int_Counter
             .Col = 0
             If .Text = "" Then Exit For
             If .Text <> l_int_PrevColVal + 1 And l_int_Counter > 1 Then l_bln_RetVal = 2
             If .Text <= l_int_PrevColVal And l_int_Counter > 1 Then l_bln_RetVal = 2
             l_int_PrevColVal = .Text
             .Col = 1

             l_int_PrevColVal = .Text
         Next
    End With
    f_bln_ValidateRange = l_bln_RetVal
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function
