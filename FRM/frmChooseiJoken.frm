VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmChooseiJoken 
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
   Picture         =   "frmChooseiJoken.frx":0000
   ScaleHeight     =   10110
   ScaleWidth      =   13230
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdDelRow 
      Caption         =   "行削除"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   23
      Top             =   9000
      Width           =   1575
   End
   Begin VB.TextBox txtSerial 
      BackColor       =   &H80000009&
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
      Left            =   6360
      TabIndex        =   21
      Top             =   9060
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddRow 
      Caption         =   "行追加"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   20
      Top             =   9000
      Width           =   1575
   End
   Begin VB.ListBox lstRooms 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1425
      Left            =   8280
      MultiSelect     =   2  '拡張
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   2655
   End
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
      Left            =   240
      TabIndex        =   10
      Top             =   9030
      Width           =   1695
   End
   Begin VB.ComboBox cboSubjectId 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   9720
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CheckBox chkRoom 
      Caption         =   "Check1"
      Height          =   200
      Left            =   10740
      TabIndex        =   5
      Top             =   1680
      Width           =   200
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "1066"
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
      Left            =   6135
      TabIndex        =   8
      Top             =   3975
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cboSubject 
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
      Left            =   2820
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1080
      Width           =   2100
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
      Left            =   4215
      TabIndex        =   7
      Top             =   3975
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox chkSex 
      Caption         =   "Check1"
      Height          =   200
      Left            =   2820
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   1770
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox chkRawScore 
      Caption         =   "Check1"
      Height          =   200
      Left            =   2820
      TabIndex        =   1
      Top             =   3615
      Value           =   1  'ﾁｪｯｸ
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox chkDay 
      Caption         =   "Check1"
      Height          =   200
      Left            =   2820
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox chkSuisen 
      Caption         =   "Check1"
      Height          =   200
      Left            =   2820
      TabIndex        =   3
      Top             =   2385
      Visible         =   0   'False
      Width           =   200
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfSearchGrid 
      Height          =   6735
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   11535
      _cx             =   20346
      _cy             =   11880
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
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
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmChooseiJoken.frx":3AD3
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
      TabBehavior     =   1
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
   Begin VSFlex7LCtl.VSFlexGrid vsfselectRawScore 
      Height          =   2775
      Left            =   5280
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
      _cx             =   4683
      _cy             =   4895
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
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmChooseiJoken.frx":3BA8
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
   Begin VB.Label lblSerial 
      BackStyle       =   0  '透明
      Caption         =   "行番号"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      Top             =   9120
      Width           =   1035
   End
   Begin VB.Label lblRoom 
      BackStyle       =   0  '透明
      Caption         =   "2002"
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
      Left            =   8280
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblRawScore 
      BackStyle       =   0  '透明
      Caption         =   "1752"
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
      Left            =   360
      TabIndex        =   16
      Top             =   3615
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  '透明
      Caption         =   "1753"
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
      Left            =   360
      TabIndex        =   15
      Top             =   1095
      Width           =   2175
   End
   Begin VB.Label lblDay 
      BackStyle       =   0  '透明
      Caption         =   "1755"
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
      Left            =   360
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblSex 
      BackStyle       =   0  '透明
      Caption         =   "1754"
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
      Left            =   360
      TabIndex        =   13
      Top             =   1650
      Visible         =   0   'False
      Width           =   2175
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
      TabIndex        =   12
      Top             =   8520
      Visible         =   0   'False
      Width           =   9735
   End
   Begin VB.Label lblSuisen 
      BackStyle       =   0  '透明
      Caption         =   "1768"
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
      Left            =   360
      TabIndex        =   11
      Top             =   2325
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmChooseiJoken"
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
Option Explicit

' database related variables
Dim m_obj_Rst As New ADODB.Recordset    ' recordset object
Dim m_str_SQL As String                 ' to store the SQL string
Dim m_int_SelectedSubject As Long    ' to store the selected subject from the subject combo
'Dim m_int_NoOfErr As Long            ' to keep track of no of errors
Dim m_int_NoOfConditions As Long     ' to track the no of conditions
Public m_int_ChoseiJoken As Long          ' to diff b/w Grace Score and Suisen Score
Dim m_bln_OnceEntered As Boolean        ' boolean stores whether the conditions have been entered once. if so,user hav to clear off first

Private prvbEditEnd As Boolean 'Edit後のセル移動のさい、EditModeのままとするためのフラグ

Private prviSvGridKeyDownEdit_KeyCode  As Integer

'表の行位置を設定
Private prvlSerialCol As Long
Private prvlSubjectNameCol As Long
Private prvlScoreFromCol As Long
Private prvlScoreToCol As Long
Private prvlSexCol As Long
Private prvlDayCol As Long
Private prvlRoomIdCol As Long
Private prvlRoomNameCol As Long
Private prvlAverageCol As Long
Private prvlChooSeiCol As Long

Private Sub cboSubject_Click()

Static stbCheck As Boolean
'add,xzg,2009/12/17,S--------
lblErrorDetails.Caption = ""
'add,xzg,2009/12/17,E--------
    ' ask user to clear off the grid, if some data is already displayed on the grid
    If m_bln_OnceEntered Then
'         lblErrorDetails.Caption = LoadResString(1772)
'         lblErrorDetails.Visible = True
        If cboSubject.ListIndex = cboSubjectId.ListIndex Then Exit Sub
        If vbCancel = MsgBox("データが変更されています。保存せずに別の科目を表示してもよろしいですか？", vbOKCancel, "変更確認") Then
            cboSubject.ListIndex = cboSubjectId.ListIndex
            Exit Sub
        End If
    End If

    cboSubjectId.ListIndex = cboSubject.ListIndex
'    Call f_void_LoadRoom
    Call f_void_InitGrid    ' reinitialize the rawscore grid
    Call f_void_ReadAlsoData
End Sub

Private Sub f_void_ReadAlsoData()

Dim sSQL As String
Dim oRs As ADODB.Recordset

On Error GoTo ErrProc

    m_bln_OnceEntered = False
'    cmdSubmit.Enabled = False

    sSQL = "SELECT "
    Select Case cboSubject.Text
    Case gcsHyotei
        sSQL = sSQL & "  CASE fChoseiStartScore WHEN -1 then 'a' WHEN -2 then 'b' ELSE STR( fChoseiStartScore , 5 , 1 ) END as fChoseiStartScore "
        sSQL = sSQL & ", CASE fChoseiEndScore WHEN -1 then 'a' WHEN -2 then 'b' ELSE STR( fChoseiEndScore , 5 , 1 ) END as fChoseiEndScore "
    Case gcsKessekiNissuu
        sSQL = sSQL & "  CASE fChoseiStartScore WHEN -1 then 'a' WHEN -2 then 'b' ELSE STR( fChoseiStartScore , 5 ) END as fChoseiStartScore "
        sSQL = sSQL & ", CASE fChoseiEndScore WHEN -1 then 'a' WHEN -2 then 'b' ELSE STR( fChoseiEndScore , 5 ) END as fChoseiEndScore "
    Case Else
        sSQL = sSQL & "  STR( fChoseiStartScore , 5 ) as fChoseiStartScore "
        sSQL = sSQL & ", STR( fChoseiEndScore , 5 ) as fChoseiEndScore "
    End Select
    sSQL = sSQL & ", STR( fChoseiScore , 5 , 1 ) as fChoseiScore "
    sSQL = sSQL & " FROM  tbSTEChoseiJoken "
    sSQL = sSQL & " WHERE iSubjectProfileID = " & cboSubjectId.Text
    sSQL = sSQL & " AND   iNendo = " & g_int_CurrentNendo
    sSQL = sSQL & " AND   iChoseiJokenType = 1 " 'iSubjectProfileIdにはそのものが入っている
    sSQL = sSQL & " ORDER BY fChoseiStartScore , fChoseiEndScore "
    'add,xzg,2009/12/02,S-----
    If cboSubject.Text = "ピンポイント" Then
'    sSQL = "SELECT "
'    sSQL = sSQL & " FROM tbSTEChoseiJoken"
'    sSQL = sSQL & " WHERE iSubjectProfileID = " & cboSubjectId.Text
'    sSQL = sSQL & " AND   iNendo = " & g_int_CurrentNendo
'    sSQL = sSQL & " AND   iChoseiJokenType = 1 " 'iSubjectProfileIdにはそのものが入っている
'    sSQL = sSQL & " ORDER BY fChoseiStartScore , fChoseiEndScore "
    End If
    'add,xzg,2009/12/02,E-----
    Set oRs = g_obj_Conn.Execute(sSQL)

    If oRs.EOF Then
        Call l_SearchGridAddRow(1)
        Set oRs = Nothing
        Exit Sub
    End If

    With vsfSearchGrid

        .Redraw = flexRDNone

        Do Until oRs.EOF
'Colsは7
'0:行番号
'1:科目
'2:RawScoreStart
'3:RawScoreEnd
'4:Sex(未使用)
'5:Avarage(未使用)
'6:ChooseiScore
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = prvlSerialCol
            .Text = Trim(str(.Row))
            .Col = prvlSubjectNameCol
            .Text = cboSubject.Text
            .Col = prvlScoreFromCol
            .Text = oRs.Fields(0)
            .Col = prvlScoreToCol
            .Text = oRs.Fields(1)
            .Col = prvlChooSeiCol
            .Text = oRs.Fields(2)

            oRs.MoveNext

        Loop

        oRs.Close
        Set oRs = Nothing

        .Redraw = flexRDDirect

    End With

Exit Sub
ErrProc:

End Sub

Private Sub chkRawScore_Click()
    ' if its already checked and some values are there in the rawscore grid, then clear it
        ' and then make id disabled
    ' if its not checked yet, check it and make the grid editable - default value beig 0-100
    Dim l_int_Counter As Long        ' counter variable
    On Error GoTo ErrorHandler
    
    If chkRawScore.Value = 1 Then
        ' not checked yet - enable the grid
        vsfselectRawScore.Editable = flexEDKbdMouse
    Else
        ' already checked - clear and make disabled
        vsfselectRawScore.Editable = flexEDNone
        With vsfselectRawScore
            If .Rows > 1 Then
               For l_int_Counter = .Rows - 1 To 2 Step -1  ' for all rows.. remove them
                   .RemoveItem l_int_Counter
               Next
               .Row = 1
               .Col = 0
               .Text = 0
               .Col = 1
               .Text = 100
            End If
        End With
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub chkRoom_Click()
    ' enable/disable the check box for room
    If chkRoom.Value = 1 Then
        lstRooms.Enabled = True
    Else
        lstRooms.Enabled = False
    End If
End Sub

Private Sub cmdAddRow_Click()

Dim sWk As String
Dim lRow As Long

    sWk = Trim(txtSerial.Text)
    If sWk = "" Then
        Call l_SearchGridAddRow(1)
        Exit Sub
    End If
    If Not gf_IntCheck(sWk) Then
        Call l_SearchGridAddRow(1)
        Exit Sub
    End If

    With vsfSearchGrid
        Call l_SearchGridAddRow(CInt(sWk))
        For lRow = CInt(sWk) To .Rows - 1
            .TextMatrix(lRow, prvlSerialCol) = Trim(str(lRow))
        Next
    End With

End Sub

Private Sub cmdClear_Click()
    ' clear the main grid as well as the raw score grid
    Dim l_int_Counter As Long                ' counter variable
    On Error GoTo ErrorHandler
    
    ' clear the main grid
    With vsfSearchGrid
         For l_int_Counter = .Rows - 1 To 1 Step -1    ' for all rows.. remove them
            .RemoveItem l_int_Counter
        Next
    End With
    
    ' clear the raw score grid
    With vsfselectRawScore
        If vsfselectRawScore.Rows > 1 Then
           For l_int_Counter = .Rows - 1 To 2 Step -1  ' for all rows.. remove them
               .RemoveItem l_int_Counter
           Next
        End If
        .Row = 1
        .Col = 0
        .Text = 0
        .Col = 1
        .Text = 100
    End With
    lblErrorDetails.Caption = ""
    m_bln_OnceEntered = False
'    cmdSubmit.Enabled = False
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdDelRow_Click()

Dim sWk As String
Dim lRow As Long

On Error GoTo ErrProc

    sWk = Trim(txtSerial.Text)
    If sWk = "" Then
        MsgBox "行が指定されていません", vbOKOnly, "確認"
        Exit Sub
    End If
    If Not gf_IntCheck(sWk) Then
        MsgBox "行が指定されていません", vbOKOnly, "確認"
        Exit Sub
    End If
    If CInt(sWk) < 1 Then
        MsgBox "行の指定が不正です", vbOKOnly, "確認"
        Exit Sub
    End If

    With vsfSearchGrid
        If vbOK = MsgBox("指定された行を削除します。よろしいですか？" & vbCrLf & vbCrLf & "行番号:   " & sWk & vbCrLf, vbOKCancel, "削除確認") Then
            .Row = CInt(sWk)
            m_bln_OnceEntered = True
            cmdSubmit.Enabled = True
            For lRow = .Row + 1 To .Rows - 1
                .TextMatrix(lRow, prvlSerialCol) = Trim(str(lRow - 1))
            Next
            .RemoveItem .Row
        End If
    End With

Exit Sub
ErrProc:
    MsgBox "Err:" & Err.Number & ":" & Err.Description, vbOKOnly, "エラー"

End Sub

Private Sub cmdOK_Click()
    ' add ros to the grid and populate it, based on the selected input criteria
    Dim l_int_Counter As Long        ' counter
    Dim l_dbl_RawScoreFrom As Double    ' to store lower limit of raw score
    Dim l_dbl_RawScoreTo As Double      ' to store upper limit of raw score
    Dim l_int_ChkDay As Long         ' day is checked or not
    Dim l_int_Count As Long          ' counter
    Dim l_int_room  As Long          ' Room is checked or not
    Dim l_int_RoomId As Long         'Room Id to be populated in Grid
    Dim l_Str_RoomDesc As String        'Room Desc to be populated in Grid
    Dim l_int_RoomCount As Long
    Dim l_bln_RoomSelected As Boolean       ' boolean stores whether a room is selected
    
    On Error GoTo ErrorHandler

    ' ask user to clear off the grid, if some data is already displayed on the grid
    If m_bln_OnceEntered Then
         lblErrorDetails.Caption = LoadResString(1772)
         lblErrorDetails.Visible = True
        Exit Sub
    End If
        
    vsfSearchGrid.Redraw = flexRDNone

    If chkRawScore.Value = 1 Then
        If f_bln_ValidateRange > 0 Then
           If f_bln_ValidateRange = 1 Then
              lblErrorDetails.Caption = LoadResString(1762)
              lblErrorDetails.Visible = True
           Else
              lblErrorDetails.Caption = LoadResString(1771)
              lblErrorDetails.Visible = True
           End If
           Exit Sub
        End If
    End If

    m_bln_OnceEntered = True
'    m_int_NoOfErr = 0
           
    'Instead of combo, loop through the vsfselectRawScore Grid
    For l_int_Counter = 1 To vsfselectRawScore.Rows - 1  ' for all rows
         If chkRawScore.Value = 1 Then
             vsfselectRawScore.Row = l_int_Counter  ' row counter
             vsfselectRawScore.Col = 0   '0th column
            
             If vsfselectRawScore.Text = "" Then Exit For
             
             If IsNull(vsfselectRawScore.Text) Then Exit For   'exit if no value in the row
             If vsfselectRawScore.Text = "a" Then
                l_dbl_RawScoreFrom = -1
             Else
                If vsfselectRawScore.Text = "b" Then
                    l_dbl_RawScoreFrom = -2
                Else
                    l_dbl_RawScoreFrom = vsfselectRawScore.Text
                End If
             End If

             vsfselectRawScore.Col = 1   'fist column
             If vsfselectRawScore.Text = "" Then
                 l_dbl_RawScoreTo = l_dbl_RawScoreFrom
             Else
                 l_dbl_RawScoreTo = vsfselectRawScore.Text
             End If
        Else
            l_dbl_RawScoreFrom = 0
            l_dbl_RawScoreTo = 100
            If l_int_Counter > 1 Then Exit For
        End If

        With vsfSearchGrid
        'del,xzg,2009/12/02,
'        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'            l_int_ChkDay = IIf(chkDay.Value = 1, 1, 0)   'Day is checked?
'            l_int_room = IIf(chkRoom.Value = 1, 1, 0)    'Room is Checked?
'        Else
            l_int_ChkDay = 0
            l_int_room = 0
'        End If
        'loop for all rows of romm master if room checkbox is checked
        If l_int_room = 1 Then
            'check whether any room is selected or not in the listbox
            For l_int_RoomCount = 0 To lstRooms.ListCount - 1
                If lstRooms.Selected(l_int_RoomCount) = True Then
                    l_bln_RoomSelected = True
                End If
            Next
            If Not l_bln_RoomSelected Then
                lblErrorDetails.Caption = LoadResString(2495)   '"Select a room"
                lblErrorDetails.Visible = True
                Exit Sub
            End If
            For l_int_RoomCount = 0 To lstRooms.ListCount - 1
                If lstRooms.Selected(l_int_RoomCount) = True Then 'if the current item is selected
                    l_int_RoomId = lstRooms.ItemData(l_int_RoomCount)
                    l_Str_RoomDesc = lstRooms.List(l_int_RoomCount)
                    If chkSex.Value = 1 Then
                        If l_int_ChkDay = 1 Then
                            For l_int_Count = 1 To 3  'adds 3 rows
                                ' sex is checked, so add 2 rows to the grid
                                .AddItem "", .Rows
                                .Row = .Rows - 1
                                Call f_void_PopulateGrid(1, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
                                 
                                .AddItem "", .Rows
'                                .Row=.Row + 1
                                .Row = .Rows - 1
                                Call f_void_PopulateGrid(2, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
                            Next
                        Else
                            ' sex is checked, so add 2 rows to the grid
                            .AddItem "", .Rows
                            .Row = .Rows - 1
                            Call f_void_PopulateGrid(1, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
                            
                            .AddItem "", .Rows
'                            .Row = .Row + 1
                            .Row = .Rows - 1
                            Call f_void_PopulateGrid(2, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
                        End If
                    Else
                        If l_int_ChkDay = 1 Then
                            For l_int_Count = 1 To 3
                                ' sex not checked, so add only 1 row to the grid
                                .AddItem "", .Rows
'                                .Row = .Row + 1
                                .Row = .Rows - 1
                                Call f_void_PopulateGrid(0, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
                            Next
                        Else
                            ' sex not checked, so add only 1 row to the grid
                                .AddItem "", .Rows
'                                .Row = .Row + 1
                                .Row = .Rows - 1
                                Call f_void_PopulateGrid(0, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo, l_int_RoomId, l_Str_RoomDesc)
                        End If
                    End If
                End If 'if the item in list is selected
    
            Next 'for all items in list box
        Else     'original case. No room checkbox checked
            If chkSex.Value = 1 Then
                If l_int_ChkDay = 1 Then
                    For l_int_Count = 1 To 3  'adds 3 rows
                        ' sex is checked, so add 2 rows to the grid
                        .AddItem "", .Rows
                        .Row = .Rows - 1
                        Call f_void_PopulateGrid(1, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
                         
                        .AddItem "", .Rows
'                        .Row = .Row + 1
                        .Row = .Rows - 1
                        Call f_void_PopulateGrid(2, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
                    Next
                Else
                    ' sex is checked, so add 2 rows to the grid
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    Call f_void_PopulateGrid(1, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
                    
                    .AddItem "", .Rows
'                    .Row = .Row + 1
                    .Row = .Rows - 1
                    Call f_void_PopulateGrid(2, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
                End If
            Else
                If l_int_ChkDay = 1 Then
                    For l_int_Count = 1 To 3
                        ' sex not checked, so add only 1 row to the grid
                        .AddItem "", .Rows
'                        .Row = .Row + 1
                        .Row = .Rows - 1
                        Call f_void_PopulateGrid(0, l_int_Count, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
                    Next
                Else
                    ' sex not checked, so add only 1 row to the grid
                        .AddItem "", .Rows
'                        .Row = .Row + 1
                        .Row = .Rows - 1
                        Call f_void_PopulateGrid(0, 0, l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
                End If
            End If
        End If   'chkRoom checked
        End With
    Next
    cmdSubmit.Enabled = True
    vsfSearchGrid.Redraw = flexRDBuffered
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdSubmit_Click()
    Dim l_int_Counter As Long            ' counter
    Dim l_str_Sql As String                 ' to store SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset variable
    Dim l_dbl_ChooseiScore As Double        ' to store the choosei score
    Dim l_Bln_RecordsUpdated  As Boolean    ' to check whether any variables are updated or not
    Dim l_int_rawScoreFrom As Double
    Dim l_int_rawScoreTo As Double

Dim lRtn As Long

Dim lAffectRecCount As Long

On Error GoTo ErrorHandler

    vsfSearchGrid.Redraw = flexRDNone

    If chkRawScore.Value = 1 Then
        lRtn = f_bln_ValidateRange
        If lRtn > 0 Then
           If lRtn = 1 Then
              lblErrorDetails.Caption = LoadResString(1762)
              lblErrorDetails.Visible = True
           Else
              lblErrorDetails.Caption = LoadResString(1771)
              lblErrorDetails.Visible = True
           End If
           Exit Sub
        End If
    End If

    g_obj_Conn.BeginTrans                   ' all the records in the grid has to be updated or else rollback

    lRtn = gflDelChoseiJoken(g_int_CurrentNendo, cboSubjectId.Text, 1)
    'add,xzg,2009/12/17,S----------------
    '調整点を更新する前、調整点を全部0で設定する。
    m_str_SQL = "UPDATE tbSTEScoreProfile"
    m_str_SQL = m_str_SQL & " SET fChoseiScore =0 "
    m_str_SQL = m_str_SQL & " ,dtUpdate=getdate()"
    m_str_SQL = m_str_SQL & " WHERE iSubjectProfileId = " & cboSubjectId.Text
    m_str_SQL = m_str_SQL & " AND iExamineeProfileId IN("
    m_str_SQL = m_str_SQL & " SELECT iExamineeProfileId FROM tbSTEExamineeProfile"
    m_str_SQL = m_str_SQL & " WHERE inendo=" & g_int_CurrentNendo
    m_str_SQL = m_str_SQL & " AND iAbsentFlag = 0"
    m_str_SQL = m_str_SQL & " )"
    g_obj_Conn.Execute m_str_SQL
    'add,xzg,2009/12/17,E----------------
    With vsfSearchGrid
        For l_int_Counter = 1 To .Rows - 1
            .Row = l_int_Counter
            
            'Changes start
            .Col = 2
            If .Text = "" Then Exit For
            If .Text = "a" Then
                l_int_rawScoreFrom = -1
            ElseIf .Text = "b" Then
                l_int_rawScoreFrom = -2
            Else
                l_int_rawScoreFrom = .Text
            End If
            .Col = 3
            If .Text = "a" Then
                l_int_rawScoreTo = -1
            ElseIf .Text = "b" Then
                l_int_rawScoreTo = -2
            Else
                l_int_rawScoreTo = .Text
            End If
            .Col = 6  'RommProfileId
            If cboSubject.Text = "現浪区分" Then
                l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile as ep "
                l_str_Sql = l_str_Sql & " WHERE ep.iAdmissionType1 between " & l_int_rawScoreFrom & " and " & l_int_rawScoreTo & " "
                l_str_Sql = l_str_Sql & " and ep.inendo=" & g_int_CurrentNendo
                l_str_Sql = l_str_Sql & " And ep.iAbsentFlag = 0 "
                'add,xzg,2009/12/02,S------------------
            ElseIf cboSubject.Text = "ピンポイント" Then
                l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile as ep "
                l_str_Sql = l_str_Sql & " WHERE ep.inendo=" & g_int_CurrentNendo
                l_str_Sql = l_str_Sql & " And ep.iAbsentFlag = 0 "
                l_str_Sql = l_str_Sql & " And ep.iJukenNumber  between " & l_int_rawScoreFrom & " and " & l_int_rawScoreTo & " "
                'add,xzg,2009/12/02,E------------------
            Else
                l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile as ep "
                l_str_Sql = l_str_Sql & " WHERE exists ( "
                l_str_Sql = l_str_Sql & " select 1 from tbSteScoreProfile as sp "
                l_str_Sql = l_str_Sql & " Where ep.iExamineeProfileId = sp.iExamineeProfileId "
                l_str_Sql = l_str_Sql & " and exists ( "
                l_str_Sql = l_str_Sql & " select 1 from tbSTESubjectQuestionProfile as sq "
                l_str_Sql = l_str_Sql & " inner join tbSTEScoreDetail as sd "
                l_str_Sql = l_str_Sql & " on sd.iScoreProfileId = sp.iScoreProfileId "
                l_str_Sql = l_str_Sql & " and sd.iSubjectQuestionId = sq.iSubjectQuestionId "
                l_str_Sql = l_str_Sql & " Where sq.iSubjectProfileId = sp.iSubjectProfileId "
                l_str_Sql = l_str_Sql & " and sq.vQuestionName = '" & cboSubject.Text & "' "
                l_str_Sql = l_str_Sql & " and sd.fDetailScore between " & l_int_rawScoreFrom & " and " & l_int_rawScoreTo & " and iAbsentFlag=0))"
                l_str_Sql = l_str_Sql & " and ep.inendo=" & g_int_CurrentNendo
                l_str_Sql = l_str_Sql & " And ep.iAbsentFlag = 0 "
            End If
            'changes end
            'update,xzg,2009/12/02,S----------
'            Select Case g_int_ExamType
'            Case 1
'                .Col = 6
'            Case 2, 3, 4, 5
'                If chkRoom.Value = 1 Then
'                    .Col = 9      '7
'                Else
'                    .Col = 7
'                End If
'
'            End Select
            .Col = 6
            'update,xzg,2009/12/02,E----------
            If Len(Trim(.Text)) = 0 Then
                l_dbl_ChooseiScore = 0
            Else
                l_dbl_ChooseiScore = .Text
            End If
            .Col = 4
'            If .Text = LoadResString(1837) Then
'                l_str_Sql = l_str_Sql & " AND iSex = 0"
'            ElseIf .Text = LoadResString(1838) Then
'                l_str_Sql = l_str_Sql & " AND iSex = 1"
'            End If
'            If chkSuisen.Value = 1 Then
'                l_str_Sql = l_str_Sql & " AND iSuisenFlagId = 1"
'            End If
            .Col = .Col + 1
            
            l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly

            If Not l_obj_Rst.EOF Then
                Do While Not l_obj_Rst.EOF
                    m_str_SQL = "UPDATE tbSTEScoreProfile"
                    m_str_SQL = m_str_SQL & " SET fChoseiScore = " & l_dbl_ChooseiScore
                    m_str_SQL = m_str_SQL & ", dtUpdate='" & Format(Date, "HH:MM:SS MM/DD/YYYY") & "'"
                    m_str_SQL = m_str_SQL & " WHERE iExamineeProfileId = " & l_obj_Rst("iExamineeProfileId")
'                    m_str_SQl = m_str_SQl & " AND exists ( select 1 from tbSTEsubjectProfile where iExamtype = 1 and tbSTEScoreProfile.iSubjectProfileId = tbSTESubjectProfile.iSubjectProfileId ) "
                    m_str_SQL = m_str_SQL & " AND iSubjectProfileId = " & cboSubjectId.Text
                    Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQL, lAffectRecCount)

                    If lAffectRecCount = 0 Then
                        Set m_obj_Rst = Nothing
                        m_str_SQL = "Insert Into tbSTEScoreProfile ( "
                        m_str_SQL = m_str_SQL & " iScoreProfileId "
                        m_str_SQL = m_str_SQL & " , iSubjectProfileId "
                        m_str_SQL = m_str_SQL & " , iExamineeProfileId "
                        m_str_SQL = m_str_SQL & " , fRawScore "
                        m_str_SQL = m_str_SQL & " , fChoseiScore "
                        m_str_SQL = m_str_SQL & " , iAbsentFlag "
                        m_str_SQL = m_str_SQL & " , dtCreate "
                        m_str_SQL = m_str_SQL & " , dtUpdate )"
                        m_str_SQL = m_str_SQL & " SELECT IsNull( Max( iScoreProfileId ) + 1 , 1 ) "
                        m_str_SQL = m_str_SQL & " , " & cboSubjectId.Text & " "
                        m_str_SQL = m_str_SQL & " , " & l_obj_Rst("iExamineeProfileId") & " "
                        m_str_SQL = m_str_SQL & " , 0 "
                        m_str_SQL = m_str_SQL & " , " & l_dbl_ChooseiScore & " "
                        m_str_SQL = m_str_SQL & " , 0 "
                        m_str_SQL = m_str_SQL & " , getdate() "
                        m_str_SQL = m_str_SQL & " , getdate() "
                        m_str_SQL = m_str_SQL & " From tbSTEScoreProfile "
                        Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQL, lAffectRecCount)
                    End If
                    
                    Set m_obj_Rst = Nothing

                    l_obj_Rst.MoveNext
                Loop
                l_Bln_RecordsUpdated = True
            End If
            l_obj_Rst.Close
            Set l_obj_Rst = Nothing
            lRtn = gflInsChoseiJoken(g_int_CurrentNendo, cboSubjectId.Text, 1, l_int_rawScoreFrom, l_int_rawScoreTo, "-1", -1, l_dbl_ChooseiScore)
        Next
    End With
    
    g_obj_Conn.CommitTrans
    If l_Bln_RecordsUpdated Then
        m_bln_OnceEntered = False
'        cmdSubmit.Enabled = False
        lblErrorDetails.Caption = LoadResString(2404)
    Else
        lblErrorDetails.Caption = LoadResString(2427)
    End If
    lblErrorDetails.Visible = True
    vsfSearchGrid.Redraw = flexRDBuffered
    Exit Sub
ErrorHandler:
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
        Me.Caption = "調整点条件入力"
    Else
        Me.Caption = LoadResString(1751)
    End If
    Call g_void_SetFontProperties(Me)     ' set the font properties
    m_int_NoOfConditions = 0    ' initialise the no of conditions
    ' select all subjects that come under the selected exam type
    m_str_SQL = "SELECT sq.iSubjectQuestionId, sq.vQuestionName FROM tbSTESubjectProfile as sp "
    m_str_SQL = m_str_SQL & " INNER JOIN tbSTESubjectQuestionProfile as sq On sq.iSubjectProfileId = sp.iSubjectProfileId "
'    m_str_SQl = m_str_SQl & " WHERE sp.iSubType = 0 "
'update,xzg,2009/12/02,S----
    'm_str_SQL = m_str_SQL & " WHERE sp.iExamType = 0 "
    If glUserLevel = 1 Then
        m_str_SQL = m_str_SQL & " WHERE sp.iExamType IN(0,-1)  "
    Else
        m_str_SQL = m_str_SQL & " WHERE sp.iExamType = 0 "
    End If
'update,xzg,2009/12/02,E----
    m_str_SQL = m_str_SQL & " ORDER BY iDispOrder"

    Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQL)
'    cmdSubmit.Enabled = False

    If Not m_obj_Rst.EOF Then
        m_int_SelectedSubject = m_obj_Rst("iSubjectQuestionId")
        ' add the subjects to combo box
        Do While Not m_obj_Rst.EOF
            cboSubject.AddItem m_obj_Rst("vQuestionName")
            cboSubjectId.AddItem m_obj_Rst("iSubjectQuestionId")
            m_obj_Rst.MoveNext
        Loop
        cboSubject.ListIndex = 0
        'update,xzg,2009/12/02,S----------
'        If g_int_ExamType = 1 Then
'            ' 1st Exam
'            lblDay.Visible = False
'            chkDay.Visible = False
'            lblRoom.Visible = False
'            chkRoom.Visible = False
'            lstRooms.Visible = False
'        ElseIf g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'            ' 2nd exam
'            lblDay.Visible = True
'            chkDay.Visible = True
'            lblRoom.Visible = True
'            chkRoom.Visible = True
'            lstRooms.Visible = True
'        End If
            ' 1st Exam
            lblDay.Visible = False
            chkDay.Visible = False
            lblRoom.Visible = False
            chkRoom.Visible = False
            lstRooms.Visible = False
        'update,xzg,2009/12/02,E----------
    End If
    
    ' release the object variables
    Set m_obj_Rst = Nothing
    
    Call f_void_InitGrid            ' reinitialize the grid
    Call f_void_InitRawScoreGrid    ' reinitialize the rawscore grid
    Call f_void_LoadRoom            ' Room is a checkbox now
    Call f_void_ReadAlsoData
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_InitGrid()

    prvlSerialCol = -1
    prvlSubjectNameCol = -1
    prvlScoreFromCol = -1
    prvlScoreToCol = -1
    prvlSexCol = -1
    prvlDayCol = -1
    prvlRoomIdCol = -1
    prvlRoomNameCol = -1
    prvlAverageCol = -1
    prvlChooSeiCol = -1

    With vsfSearchGrid
        .Redraw = flexRDNone
        .Clear
        .Visible = False
        .BackColor = &HFFFFFF
        .BackColorBkg = &HFFFFFF
        .BackColorFixed = &H8000000F
        .BackColorSel = &H800000
        .FixedCols = 0
        .TextStyleFixed = flexTextFlat
        .ForeColorFixed = &H80000008
        .ForeColor = &H800000
'        .CellTextStyle = "0"
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .GridColor = &H808080
        .AllowUserResizing = flexResizeColumns
        .Visible = True
        .Rows = 1
        'update,xzg,2009/12/02,Ifチェックを外す
'        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'            ' for second exam one additional column is required for the day combo
'            ' If room checkbox is checked, 2 columns for Room id and name
'            If chkRoom.Value = 1 Then
'                .Cols = 10
'            Else
'                .Cols = 8
'            End If
'        Else
            ' for ist exam, day column is not there, hence one column less
            .Cols = 7
'        End If
        
        .Row = 0
        .Col = 0
        prvlSerialCol = .Col
        .ColWidth(0) = 700
        .Text = LoadResString(1756)   'Sr no  0
        .CellAlignment = flexAlignRightBottom
        
        .Col = .Col + 1
        prvlSubjectNameCol = .Col
        .ColWidth(1) = 2200
        .Text = LoadResString(1757)    'subject  1
        
        .Col = .Col + 1
        prvlScoreFromCol = .Col
        .ColWidth(2) = 2000
        .Text = LoadResString(1758)  'Raw score from  2
        .CellAlignment = flexAlignRightBottom
        
        .Col = .Col + 1
        prvlScoreToCol = .Col
        .ColWidth(3) = 2000
        .Text = LoadResString(1759)   'raw score to  3
        .CellAlignment = flexAlignRightBottom
        
        .Col = .Col + 1
        prvlSexCol = .Col
'        .ColWidth(4) = 1200
        .ColWidth(4) = 0
        .Text = LoadResString(1754)   'Sex  4
'del,xzg,2009/12/02,Ifチェックを外す
'        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'            ' add the additional column for the day
'            .Col = .Col + 1
'            prvlDayCol = .Col
'            .ColWidth(7) = 1600
'            .Text = LoadResString(1755)  'Day   Col is 5
'            'new col for roomID
'            If chkRoom.Value = 1 Then
'                .Col = .Col + 1
'                prvlRoomIdCol = .Col
'                .ColWidth(6) = 0 'hidden column 6 for room Id
'                .Col = .Col + 1
'                prvlRoomNameCol = .Col
'                .ColWidth(7) = 2000   'Column 7 for Room Desc
'                .Text = LoadResString(2002)
'            End If
'        End If
        
        .Col = .Col + 1
        prvlAverageCol = .Col
'        .ColWidth(.Col) = 1000  '5
        .ColWidth(.Col) = 0  '5
        .Text = LoadResString(1760)  'Col 8 Average
        .CellAlignment = flexAlignRightBottom
        
        .Col = .Col + 1
        prvlChooSeiCol = .Col
        .ColWidth(.Col) = 1700   '6
        .Text = LoadResString(1751)  'col 9 choosei score (last column)
        .CellAlignment = flexAlignRightBottom
    End With
        vsfSearchGrid.Redraw = flexRDBuffered

    Exit Sub
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
        .Visible = False
       
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

Private Sub f_void_PopulateGrid(ByVal l_bln_SexFlag As Integer, ByVal l_bln_DayFlag As Integer, ByVal l_dbl_RawScoreFrom As Double, ByVal l_dbl_RawScoreTo As Double, Optional ByVal l_int_RoomNo As Integer, Optional ByVal l_str_RoomName As String)
    Dim l_dbl_Avg As Double         ' to store the average value calculated
    On Error GoTo ErrorHandler
    vsfSearchGrid.Redraw = flexRDNone
   
    With vsfSearchGrid
        .Col = 0
        .Text = .Rows - 1
        
        .Col = .Col + 1
        .Text = cboSubject.Text
        
        .Col = .Col + 1
        .Text = IIf(l_dbl_RawScoreFrom < 0, IIf(l_dbl_RawScoreFrom = -1, "a", "b"), l_dbl_RawScoreFrom)
        
        .Col = .Col + 1
        .Text = IIf(l_dbl_RawScoreTo < 0, IIf(l_dbl_RawScoreTo = -1, "a", "b"), l_dbl_RawScoreTo)
        
        .Col = .Col + 1
        If l_bln_SexFlag = 1 Then
            .Text = LoadResString(1837)
        ElseIf l_bln_SexFlag = 2 Then
            .Text = LoadResString(1838)
        Else
            .Text = LoadResString(1846)
        End If
        
        'del,xzg,2009/12/02
'        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'            .Col = .Col + 1
'            Select Case l_bln_DayFlag
'            Case 0
'                .Text = LoadResString(1764)
'            Case 1
'                .Text = LoadResString(1765)
'            Case 2
'                .Text = LoadResString(1766)
'            Case 3
'                .Text = LoadResString(1767)
'            End Select
'            If Not IsEmpty(l_int_RoomNo) Then
'                .Col = .Col + 1
'                .Text = l_int_RoomNo
'                .Col = .Col + 1
'                .Text = l_str_RoomName
'            End If
'        End If
        
        l_dbl_Avg = f_void_GetAverage(l_dbl_RawScoreFrom, l_dbl_RawScoreTo)
        .Col = .Cols - 2
        .Text = l_dbl_Avg
        
        .Col = .Cols - 1
        .CellBackColor = &HC0C0FF
        .Text = 0
    End With
    vsfSearchGrid.Redraw = flexRDBuffered
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_bln_OnceEntered = False
'    cmdSubmit.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call g_void_CloseChildForm
End Sub

Private Sub l_SearchGridAddRow(plRow As Long)

Dim sWk As String

    With vsfSearchGrid
        sWk = Trim(str(.Rows)) & vbTab & cboSubject.Text
        .AddItem sWk, plRow
    End With

End Sub

Private Sub txtSerial_KeyPress(KeyAscii As Integer)
    Call NumericOnly(Me, KeyAscii)
End Sub

Private Sub vsfSearchGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim lCol As Long
Dim lRow As Long

    ' this code is written to round off the decimal values to 2 digits precision
    With vsfSearchGrid
        Select Case .Col
        Case prvlChooSeiCol
            If Trim(.TextMatrix(Row, Col)) <> "" Then
    '            .TextMatrix(Row, Col) = Round(.TextMatrix(Row, Col), 2)
                If gf_DblCheck(.TextMatrix(Row, Col)) Then
                    .TextMatrix(Row, Col) = Format(CDbl(.TextMatrix(Row, Col)), "##0.0")
                    If .Rows = .Row + 1 Then
                        Call l_SearchGridAddRow(.Rows)
                        .Row = .Rows - 1
                        .Col = prvlScoreFromCol
                    End If
                Else
                    .TextMatrix(Row, Col) = ""
                End If
            End If
        Case prvlScoreFromCol
            If Trim(.TextMatrix(Row, Col)) <> "" Then
    '            .TextMatrix(Row, Col) = Round(.TextMatrix(Row, Col), 2)
                Select Case .TextMatrix(Row, prvlSubjectNameCol)
                Case gcsKessekiNissuu
                    If .TextMatrix(Row, Col) = "a" Or .TextMatrix(Row, Col) = "b" Or .TextMatrix(Row, Col) = "A" Or .TextMatrix(Row, Col) = "B" Then
                        .TextMatrix(Row, Col) = StrConv(.TextMatrix(Row, Col), vbLowerCase)
                        .TextMatrix(Row, prvlScoreToCol) = .TextMatrix(Row, Col)
                    ElseIf gf_IntCheck(.TextMatrix(Row, Col)) Then
                        .TextMatrix(Row, Col) = Format(CDbl(.TextMatrix(Row, Col)), "###0")
                    Else
                        .TextMatrix(Row, Col) = ""
                    End If
                    'add,xzg,2009/12/02,S---------
                Case "ピンポイント"
                    
                    'add,xzg,2009/12/02,E---------
                Case Else
                    If .TextMatrix(Row, Col) = "a" Or .TextMatrix(Row, Col) = "b" Or .TextMatrix(Row, Col) = "A" Or .TextMatrix(Row, Col) = "B" Then
                        .TextMatrix(Row, Col) = StrConv(.TextMatrix(Row, Col), vbLowerCase)
                        .TextMatrix(Row, prvlScoreToCol) = .TextMatrix(Row, Col)
                    ElseIf gf_DblCheck(.TextMatrix(Row, Col)) Then
                        .TextMatrix(Row, Col) = Format(CDbl(.TextMatrix(Row, Col)), "##0.0")
                    Else
                        .TextMatrix(Row, Col) = ""
                    End If
                End Select
                If .TextMatrix(Row, Col) = "a" Or .TextMatrix(Row, Col) = "b" Then
                    .Col = prvlChooSeiCol
                Else
                    .Col = prvlScoreToCol
                End If
            End If
        Case prvlScoreToCol
            If Trim(.TextMatrix(Row, Col)) <> "" Then
    '            .TextMatrix(Row, Col) = Round(.TextMatrix(Row, Col), 2)
                Select Case .TextMatrix(Row, prvlSubjectNameCol)
                Case gcsKessekiNissuu
                    If .TextMatrix(Row, Col) = "a" Or .TextMatrix(Row, Col) = "b" Or .TextMatrix(Row, Col) = "A" Or .TextMatrix(Row, Col) = "B" Then
                        .TextMatrix(Row, prvlScoreFromCol) = StrConv(.TextMatrix(Row, Col), vbLowerCase)
                        .TextMatrix(Row, Col) = .TextMatrix(Row, prvlScoreFromCol)
                    ElseIf gf_IntCheck(.TextMatrix(Row, Col)) Then
                        .TextMatrix(Row, Col) = Format(CDbl(.TextMatrix(Row, Col)), "###0")
                    Else
                        .TextMatrix(Row, Col) = ""
                    End If
                    'add,xzg,2009/12/02,S---------
                Case "ピンポイント"
                    
                    'add,xzg,2009/12/02,E---------
                Case Else
                    If .TextMatrix(Row, Col) = "a" Or .TextMatrix(Row, Col) = "b" Or .TextMatrix(Row, Col) = "A" Or .TextMatrix(Row, Col) = "B" Then
                        .TextMatrix(Row, prvlScoreFromCol) = StrConv(.TextMatrix(Row, Col), vbLowerCase)
                        .TextMatrix(Row, Col) = .TextMatrix(Row, prvlScoreFromCol)
                    ElseIf gf_DblCheck(.TextMatrix(Row, Col)) Then
                        .TextMatrix(Row, Col) = Format(CDbl(.TextMatrix(Row, Col)), "##0.0")
                    Else
                        .TextMatrix(Row, Col) = ""
                    End If
                End Select
                .Col = prvlChooSeiCol
            End If
        End Select
    End With
End Sub

Private Sub vsfSearchGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfSearchGrid
        If .Redraw <> flexRDNone Then
            Select Case .Col
            Case prvlChooSeiCol, prvlScoreFromCol, prvlScoreToCol
                .Editable = flexEDKbdMouse
            End Select
        End If
    End With
End Sub

Private Sub vsfSearchGrid_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfSearchGrid
        If .Redraw <> flexRDNone Then
            Select Case NewCol
            Case prvlChooSeiCol, prvlScoreFromCol, prvlScoreToCol
            Case Else
                Cancel = True
                .Select NewRow, .Cols - 1
            End Select
        End If
    End With
End Sub
'
'Private Sub vsfSearchGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'
'    If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
'        ' in second exam, only the 8th column is editable (choosei score)
'        If Col <> IIf(chkRoom.Value = 1, 9, 7) Then
'            KeyAscii = 0
'        ElseIf KeyAscii = 13 Then
'            If vsfSearchGrid.Row < vsfSearchGrid.Rows - 1 Then
'                vsfSearchGrid.Row = vsfSearchGrid.Row + 1
'                vsfSearchGrid.Col = Col
'            End If
'        'This is to restrict user from entering more than one "." in the value
''        ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
''            If Not ((KeyAscii = 46 And Not (InStr(vsfSearchGrid.EditText, ".") > 0)) Or (KeyAscii = 45 And vsfSearchGrid.EditSelStart = 0 And Not (InStr(vsfSearchGrid.EditText, "-") > 0))) Then
''                KeyAscii = 0
''            End If
'        Else
'            Call NumericPeriodMinusVsfGrd(vsfSearchGrid, KeyAscii)
'        End If
'    ElseIf g_int_ExamType = 1 Then
'        ' in first exam, only the 8th column is editable (choosei score)
'        If Col <> 6 Then
'            KeyAscii = 0
'        ElseIf KeyAscii = 13 Then
'            If vsfSearchGrid.Row < vsfSearchGrid.Rows - 1 Then
'                vsfSearchGrid.Row = vsfSearchGrid.Row + 1
'                vsfSearchGrid.Col = Col
'            End If
'        'This is to restrict user from entering more than one "." in the value
''        ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
''            If Not ((KeyAscii = 46 And Not (InStr(vsfSearchGrid.EditText, ".") > 0)) Or (KeyAscii = 45 And vsfSearchGrid.EditSelStart = 0 And Not (InStr(vsfSearchGrid.EditText, "-") > 0))) Then
''                KeyAscii = 0
''            End If
'        Else
'            Call NumericPeriodMinusVsfGrd(vsfSearchGrid, KeyAscii)
'        End If
'    End If
'End Sub
'
'Private Sub vsfSearchGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'
'Dim lRow As Long
'
'    If KeyCode = vbKeyDelete Then
'        With vsfSearchGrid
'            If .Row > 0 Then
'                If vbOK = MsgBox("指定された行を削除します。よろしいですか？" & vbCrLf & vbCrLf & "行番号:   " & Trim(str(.TextMatrix(.Row, prvlSerialCol))) & vbCrLf, vbOKCancel, "削除確認") Then
'                    m_bln_OnceEntered = True
'                    cmdSubmit.Enabled = True
'                    For lRow = .Row + 1 To .Rows - 1
'                        .TextMatrix(lRow, prvlSerialCol) = Trim(str(lRow - 1))
'                    Next
'                    .RemoveItem .Row
'                End If
'            End If
'        End With
'    End If
'
'End Sub

Private Sub vsfSearchGrid_Click()

    If vsfSearchGrid.Row > 0 Then
        txtSerial.Text = Trim(str(vsfSearchGrid.TextMatrix(vsfSearchGrid.Row, prvlSerialCol)))
    End If

End Sub

Private Sub vsfSearchGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

'del,xzg,2009/12/02,del If
'    If g_int_ExamType = 1 Then
        ' in first exam, only the 8th column is editable (choosei score)
        If KeyAscii = 13 Then Exit Sub
        m_bln_OnceEntered = True
        cmdSubmit.Enabled = True
        Select Case Col
        Case prvlChooSeiCol
            Call NumericPeriodMinusVsfGrd(vsfSearchGrid, KeyAscii)
        Case prvlScoreFromCol, prvlScoreToCol
            Select Case cboSubject.Text
            Case gcsHyotei
                Call NumericPeriodABVsfGrd(vsfSearchGrid, KeyAscii)
            Case gcsKessekiNissuu
                Call NumericABVsfGrd(vsfSearchGrid, KeyAscii)
            Case Else
                Call NumericVsfGrd(vsfSearchGrid, KeyAscii)
            End Select
        Case Else
            KeyAscii = 0
        End Select
'    End If
End Sub

 Private Function f_void_GetAverage(ByVal l_dbl_RawScoreFrom As Double, ByVal l_dbl_RawScoreTo As Double) As Double
    On Error GoTo ErrorHandler
    
    m_str_SQL = "SELECT Avg(fRawScore) from tbSTEScoreProfile where fRawScore BETWEEN "
    m_str_SQL = m_str_SQL & l_dbl_RawScoreFrom & " AND " & l_dbl_RawScoreTo
    m_str_SQL = m_str_SQL & " AND iSubjectProfileId=" & cboSubjectId.Text
    
    m_obj_Rst.Open m_str_SQL, g_obj_Conn, adOpenStatic, adLockReadOnly
    
    If Trim(m_obj_Rst(0)) <> "" Then
        f_void_GetAverage = Round(m_obj_Rst(0), 2)
    Else
        f_void_GetAverage = 0
    End If
    
    m_obj_Rst.Close
    Set m_obj_Rst = Nothing
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
    f_void_GetAverage = 0
End Function

Private Sub vsfSearchGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Clipboard.Clear
    End If
End Sub

Private Sub vsfselectRawScore_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrorHandler
Dim vWk As String
    vWk = Trim(vsfselectRawScore.TextMatrix(Row, Col))
    If vWk <> "" Then
        If Me.cboSubject.Text = gcsHyotei Then
            If gf_DblCheck(vWk) Then
                vsfselectRawScore.TextMatrix(Row, Col) = Format(CDbl(vWk), "0.00")
            End If
        End If
    End If
    If Me.ActiveControl.Name = Me.vsfselectRawScore.Name Then
        With vsfselectRawScore
                If Col < .Cols - 1 Then
                    If cboSubject.Text = gcsKessekiNissuu And (Trim(.Text) = "a" Or Trim(.Text) = "A" Or Trim(.Text) = "b" Or Trim(.Text) = "B") Then
                        .Text = StrConv(.Text, vbLowerCase)
                        .Col = .Col + 1
                        .Text = ""
                        GoTo NextRow
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf Col = .Cols - 1 Then
                    If Trim(.Text) <> "" Then
NextRow:
                        .Col = 0
                        If Trim(.Text) <> "" Then
                            If .Row < .Rows - 1 Then
                                If prviSvGridKeyDownEdit_KeyCode <> vbKeyDown Then
                                    .Row = .Row + 1    'Go to last row and if its not blank, add a row
                                End If
                            Else
                                .Row = .Rows - 1    'Go to last row and if its not blank, add a row
                                .Col = 0
                                If .Text <> "" Then
                                    If .Rows < 11 Then
                                        .Rows = .Rows + 1
                                        .Row = .Rows - 1
                                        .Col = 0
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
        End With
        prvbEditEnd = True
    End If
    prviSvGridKeyDownEdit_KeyCode = 0
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub vsfselectRawScore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    With vsfselectRawScore
'        If .Redraw <> flexRDNone And Col <> .Cols - 1 Then
'         Cancel = True
'            Exit Sub
'        Else
'            .Editable = flexEDKbdMouse
'        End If
'    End With
End Sub

Private Sub vsfselectRawScore_EnterCell()
On Error GoTo ErrProc
'FormLoad時にアクティブコントロールがなくてエラーになるけど、無視するためにエラーハンドルする
    If Me.ActiveControl.Name = Me.vsfselectRawScore.Name Then
        If prvbEditEnd Then
                prvbEditEnd = False
'                vsfselectRawScore.EditCell
                vsfselectRawScore.EditSelStart = 0
                vsfselectRawScore.EditSelLength = Len(vsfselectRawScore.Text)
        End If
    End If
ErrProc:
End Sub

Private Sub vsfselectRawScore_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    prviSvGridKeyDownEdit_KeyCode = KeyCode
End Sub

Private Sub vsfselectRawScore_KeyPress(KeyAscii As Integer)
Dim bCheck As Boolean
    If chkRawScore.Value = 1 Then
        prviSvGridKeyDownEdit_KeyCode = 0
        Call vsfselectRawScore_KeyPressEdit(vsfselectRawScore.Row, vsfselectRawScore.Col, KeyAscii)
        bCheck = False
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 65 Or KeyAscii = 97 Or KeyAscii = 66 Or KeyAscii = 98 Then
            vsfselectRawScore.Text = Chr(KeyAscii)
            bCheck = True
        End If
        vsfselectRawScore.EditCell
        If bCheck Then
            vsfselectRawScore.EditSelStart = Len(vsfselectRawScore.Text)
            vsfselectRawScore.EditSelLength = 0
        End If
    End If
End Sub

Private Sub vsfselectRawScore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Clipboard.Clear
    End If
End Sub

Private Sub vsfselectRawScore_Click()
    lblErrorDetails.Caption = ""
    lblErrorDetails.Visible = False
    If chkRawScore.Value = 1 Then
        vsfselectRawScore.EditCell
    End If
End Sub

Private Sub vsfselectRawScore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim l_int_PrevCol As Integer
    
    vsfselectRawScore.Redraw = flexRDDirect
'    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> vbKeyReturn And KeyAscii <> 46 And KeyAscii <> vbKeyEscape Then '46:.
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyEscape Then '46:.
'        If Not (cboSubject.Text = gcsHyotei And KeyAscii = 46 And Not (InStr(vsfselectRawScore.EditText, ".") > 0)) Then
'            If Not ((cboSubject.Text = gcsKessekiNissuu Or cboSubject.Text = gcsHyotei) And (KeyAscii = 65 Or KeyAscii = 97 Or KeyAscii = 66 Or KeyAscii = 98)) Then
'               KeyAscii = 0
'            End If
'        End If
        If cboSubject.Text = gcsHyotei Then
            Call NumericPeriodABVsfGrd(vsfselectRawScore, KeyAscii)
        ElseIf cboSubject.Text = gcsKessekiNissuu Then
            Call NumericABVsfGrd(vsfselectRawScore, KeyAscii)
        Else
            Call NumericPeriodMinusVsfGrd(vsfselectRawScore, KeyAscii)
        End If
    End If
    If KeyAscii = vbKeyReturn Then
        prvbEditEnd = True
    End If

End Sub

Private Function f_bln_ValidateRange() As Long

    Dim l_int_Rows As Long 'total rows in grid
    Dim l_int_Counter As Long ' current row
    Dim l_bln_RetVal As Double  ' return value
    Dim l_int_PrevColVal As Double  'previous col value of same row
    'Dim l_int_PrevRowVal As Long  ' previous col value of prev row
    ' 0 means all ok
    ' 1 means check box checked but no values entered
    ' 2 means Continuity is missing
    On Error GoTo ErrorHandler
    l_bln_RetVal = 0
    
    l_int_Rows = vsfselectRawScore.Rows
    vsfselectRawScore.Row = 1
    vsfselectRawScore.Col = 0
    
    With vsfselectRawScore
NextRow:
        If .Text = "" Then
            l_bln_RetVal = 1
            f_bln_ValidateRange = l_bln_RetVal
            Exit Function
        End If
        If cboSubject.Text = gcsKessekiNissuu And (.Text = "a" Or .Text = "b") Then
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                If .Text = "" Then
                    f_bln_ValidateRange = l_bln_RetVal
                    Exit Function
                End If
                GoTo NextRow
            End If
            f_bln_ValidateRange = l_bln_RetVal
            Exit Function
        Else
            l_int_PrevColVal = .Text
        End If
         For l_int_Counter = 1 To .Rows - 1
             .Row = l_int_Counter
             .Col = 0
             If .Text = "" Then Exit For
'増加が１や０．１とは限らないため、連続値であるチェックはしない。大小比較のみとする
'どうしても必要な場合、サブジェクトが評定値なら0.1単位、などとなるはず
'             If .Text <> l_int_PrevColVal + 1 And l_int_Counter > 1 Then l_bln_RetVal = 2
'欠席日数でのａ，ｂはノーチェックで次に
             If (cboSubject.Text = gcsKessekiNissuu Or cboSubject.Text = gcsHyotei) And (.Text = "a" Or .Text = "b") Then GoTo EndFor
             If .Text <= l_int_PrevColVal And l_int_Counter > 1 Then l_bln_RetVal = 2
             .Col = 1
             l_int_PrevColVal = .Text
EndFor:
         Next
    End With
    f_bln_ValidateRange = l_bln_RetVal
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Private Function f_bln_ValidateRangeSearchGrid() As Long

    Dim l_int_Rows As Long 'total rows in grid
    Dim l_int_Counter As Long ' current row
    Dim l_bln_RetVal As Double  ' return value
    Dim l_int_PrevColVal As Double  'previous col value of same row
    'Dim l_int_PrevRowVal As Long  ' previous col value of prev row
    ' 0 means all ok
    ' 1 means check box checked but no values entered
    ' 2 means Continuity is missing
    On Error GoTo ErrorHandler
    l_bln_RetVal = 0
    
    With Me.vsfSearchGrid
        l_int_Rows = .Rows
        .Row = 1
        .Col = prvlScoreFromCol

NextRow:
        If .Text = "" Then
            l_bln_RetVal = 1
            f_bln_ValidateRangeSearchGrid = l_bln_RetVal
            Exit Function
        End If
        If (cboSubject.Text = gcsKessekiNissuu Or cboSubject.Text = gcsHyotei) And (.Text = "a" Or .Text = "b") Then
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                If .Text = "" Then
                    f_bln_ValidateRangeSearchGrid = l_bln_RetVal
                    Exit Function
                End If
                GoTo NextRow
            End If
            f_bln_ValidateRangeSearchGrid = l_bln_RetVal
            Exit Function
        Else
            l_int_PrevColVal = .Text
        End If
         For l_int_Counter = 1 To .Rows - 1
             .Row = l_int_Counter
             .Col = prvlScoreFromCol
             If .Text = "" Then Exit For
'増加が１や０．１とは限らないため、連続値であるチェックはしない。大小比較のみとする
'どうしても必要な場合、サブジェクトが評定値なら0.1単位、などとなるはず
'             If .Text <> l_int_PrevColVal + 1 And l_int_Counter > 1 Then l_bln_RetVal = 2
'欠席日数でのａ，ｂはノーチェックで次に
             If (cboSubject.Text = gcsKessekiNissuu Or cboSubject.Text = gcsHyotei) And (.Text = "a" Or .Text = "b") Then GoTo EndFor
             If .Text <= l_int_PrevColVal And l_int_Counter > 1 Then l_bln_RetVal = 2
             .Col = prvlScoreToCol
             l_int_PrevColVal = .Text
EndFor:
         Next
    End With
    f_bln_ValidateRangeSearchGrid = l_bln_RetVal
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Public Sub f_void_LoadRoom()        'populate the room names
'stores roomprofileid in listbox itself using NweIndex method. This saves extra sql to b fired for ID
    Dim l_obj_RsRoom As New ADODB.Recordset
    Dim l_str_sqlRoom As String
    Dim l_str_ExamType As String
    Dim l_obj_rsExamType As New ADODB.Recordset
    Dim l_int_iExamType As Integer
    ' code changed on 31/07/02 to accomodate the change in iInterviewRoomFlag
    On Error GoTo ErrorHandler
    
    l_str_ExamType = "SELECT iExamType FROM tbSTESubjectProfile" & _
        " WHERE iSubjectProfileId = " & cboSubjectId.Text
    l_obj_rsExamType.Open l_str_ExamType, g_obj_Conn
    
    If Not l_obj_rsExamType.EOF Then
        l_int_iExamType = l_obj_rsExamType.Fields("iExamType").Value
        
        l_str_sqlRoom = "SELECT iRoomProfileid,vRoomName FROM tbSTERoomProfile" & _
            " WHERE iMaxCapacity >0"
        If l_int_iExamType = 2 Or l_int_iExamType = 4 Then
            l_str_sqlRoom = l_str_sqlRoom & " AND iInterviewRoomFlag = 0"
        Else
            l_str_sqlRoom = l_str_sqlRoom & " AND iInterviewRoomFlag = 1"
        End If
        lstRooms.Clear
        l_obj_RsRoom.Open l_str_sqlRoom, g_obj_Conn
        Do While Not l_obj_RsRoom.EOF
            lstRooms.AddItem l_obj_RsRoom.Fields("vRoomName").Value ', l_obj_RsRoom.Fields("iRoomProfileid").Value
            lstRooms.ItemData(lstRooms.NewIndex) = l_obj_RsRoom.Fields("iRoomProfileid").Value
            l_obj_RsRoom.MoveNext
        Loop
        l_obj_RsRoom.Close
        Set l_obj_RsRoom = Nothing
    
    End If
    l_obj_rsExamType.Close
    Set l_obj_rsExamType = Nothing
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

