VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmChooseiScore2 
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
   MDIChild        =   -1  'True
   Picture         =   "frmChooseiScore2.frx":0000
   ScaleHeight     =   10110
   ScaleWidth      =   13230
   WindowState     =   2  'Å‘å‰»
   Begin VB.ListBox lstRooms 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1425
      Left            =   8280
      MultiSelect     =   2  'Šg’£
      TabIndex        =   6
      Top             =   2040
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
      Left            =   5768
      TabIndex        =   10
      Top             =   9000
      Width           =   1695
   End
   Begin VB.ComboBox cboSubjectId 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   9720
      Style           =   2  'ÄÞÛ¯ÌßÀÞ³Ý Ø½Ä
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
      Style           =   2  'ÄÞÛ¯ÌßÀÞ³Ý Ø½Ä
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
      Height          =   3855
      Left            =   240
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   11535
      _cx             =   20346
      _cy             =   6800
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
      HighLight       =   1
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
      FormatString    =   $"frmChooseiScore2.frx":3AD3
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
      ShowComboButton =   1
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
      Width           =   2655
      _cx             =   4683
      _cy             =   4895
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmChooseiScore2.frx":3BA8
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
      ShowComboButton =   1
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
   Begin VB.Label lblRoom 
      BackStyle       =   0  '“§–¾
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
      Width           =   2175
   End
   Begin VB.Label lblRawScore 
      BackStyle       =   0  '“§–¾
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
      BackStyle       =   0  '“§–¾
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
      BackStyle       =   0  '“§–¾
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
      BackStyle       =   0  '“§–¾
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
      BackStyle       =   0  '“§–¾
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
      Top             =   4560
      Visible         =   0   'False
      Width           =   9735
   End
   Begin VB.Label lblSuisen 
      BackStyle       =   0  '“§–¾
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
Attribute VB_Name = "frmChooseiScore2"
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
Dim m_str_SQl As String                 ' to store the SQL string
Dim m_int_SelectedSubject As Integer    ' to store the selected subject from the subject combo
'Dim m_int_NoOfErr As Integer            ' to keep track of no of errors
Dim m_int_NoOfConditions As Integer     ' to track the no of conditions
Public m_int_ChoseiJoken As Integer          ' to diff b/w Grace Score and Suisen Score
Dim m_bln_OnceEntered As Boolean        ' boolean stores whether the conditions have been entered once. if so,user hav to clear off first

Private Sub cboSubject_Click()
    cboSubjectId.ListIndex = cboSubject.ListIndex
    Call f_void_LoadRoom
End Sub

Private Sub chkRawScore_Click()
    ' if its already checked and some values are there in the rawscore grid, then clear it
        ' and then make id disabled
    ' if its not checked yet, check it and make the grid editable - default value beig 0-100
    Dim l_int_Counter As Integer        ' counter variable
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

Private Sub cmdClear_Click()
    ' clear the main grid as well as the raw score grid
    Dim l_int_Counter As Integer                ' counter variable
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
    cmdSubmit.Enabled = False
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdOK_Click()
    ' add ros to the grid and populate it, based on the selected input criteria
    Dim l_int_Counter As Integer        ' counter
    Dim l_dbl_RawScoreFrom As Double    ' to store lower limit of raw score
    Dim l_dbl_RawScoreTo As Double      ' to store upper limit of raw score
    Dim l_int_ChkDay As Integer         ' day is checked or not
    Dim l_int_Count As Integer          ' counter
    Dim l_int_room  As Integer          ' Room is checked or not
    Dim l_int_RoomId As Integer         'Room Id to be populated in Grid
    Dim l_Str_RoomDesc As String        'Room Desc to be populated in Grid
    Dim l_int_RoomCount As Integer
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
                l_dbl_RawScoreFrom = vsfselectRawScore.Text
             
             vsfselectRawScore.Col = 1   'fist column
             l_dbl_RawScoreTo = vsfselectRawScore.Text
        Else
            l_dbl_RawScoreFrom = 0
            l_dbl_RawScoreTo = 100
            If l_int_Counter > 1 Then Exit For
        End If

        With vsfSearchGrid
        
        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
            l_int_ChkDay = IIf(chkDay.Value = 1, 1, 0)   'Day is checked?
            l_int_room = IIf(chkRoom.Value = 1, 1, 0)    'Room is Checked?
        Else
            l_int_ChkDay = 0
            l_int_room = 0
        End If
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
    Dim l_int_Counter As Integer            ' counter
    Dim l_str_Sql As String                 ' to store SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset variable
    Dim l_dbl_ChooseiScore As Double        ' to store the choosei score
    Dim l_Bln_RecordsUpdated  As Boolean    ' to check whether any variables are updated or not
    Dim l_int_rawScoreFrom As Integer
    Dim l_int_rawScoreTo As Integer
    On Error GoTo ErrorHandler
    vsfSearchGrid.Redraw = flexRDNone

    g_obj_Conn.BeginTrans                   ' all the records in the grid has to be updated or else rollback
    With vsfSearchGrid
        For l_int_Counter = 1 To .Rows - 1
            .Row = l_int_Counter
            
            'Changes start
            .Col = 2
            l_int_rawScoreFrom = .Text
            .Col = 3
            l_int_rawScoreTo = .Text
            .Col = 6  'RommProfileId
            If (g_int_ExamType = 2 Or g_int_ExamType = 3) And chkRoom.Value = 1 Then
                l_str_Sql = "select iExamineeProfileId from tbSTEExamineeProfile where iExamineeProfileId In" & _
                    " (select a.iExamineeProfileId from tbSTEScoreProfile a inner join tbSTEExamineeRoomProfile b" & _
                    " on a.iExamineeProfileId=b.iExamineeProfileId and b.iroomprofileid= " & .Text & " and b.isubjectprofileid=" & cboSubjectId.Text & _
                    " and a.frawscore between " & l_int_rawScoreFrom & " and " & l_int_rawScoreTo & " and a.iSubjectProfileId= " & cboSubjectId.Text & " and a.iAbsentFlag=0) and iNendo=" & g_int_CurrentNendo & _
                    "And iAbsentFlag = 0 "
            ElseIf g_int_ExamType = 1 Or m_int_ChoseiJoken = 1 Or ((g_int_ExamType = 2 Or g_int_ExamType = 3) And chkRoom.Value = 0) Then
                l_str_Sql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile " & _
                    " WHERE iExamineeProfileId in ( select iExamineeProfileId from tbSteScoreProfile where iSubjectProfileId = " & cboSubjectId.Text & _
                    " and frawscore between " & l_int_rawScoreFrom & " and " & l_int_rawScoreTo & " and iAbsentFlag=0) and inendo=" & g_int_CurrentNendo & _
                    " And iAbsentFlag = 0 "
            End If
            'changes end
            
            Select Case g_int_ExamType
            Case 1
                .Col = 6
            Case 2, 3, 4, 5
                If chkRoom.Value = 1 Then
                    .Col = 9      '7
                Else
                    .Col = 7
                End If
                
            End Select
            If Len(Trim(.Text)) = 0 Then
                l_dbl_ChooseiScore = 0
            Else
                l_dbl_ChooseiScore = .Text
            End If
            .Col = 4
            If .Text = LoadResString(1837) Then
                l_str_Sql = l_str_Sql & " AND iSex = 0"
            ElseIf .Text = LoadResString(1838) Then
                l_str_Sql = l_str_Sql & " AND iSex = 1"
            End If
            If chkSuisen.Value = 1 Then
                l_str_Sql = l_str_Sql & " AND iSuisenFlagId = 1"
            End If
            .Col = .Col + 1
            If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
                Select Case .Text
                Case LoadResString(1765)
                    l_str_Sql = l_str_Sql & " AND CONVERT(VARCHAR(10),dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay1,101) FROM tbSTESecondExamProfile"
                    l_str_Sql = l_str_Sql & " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile"
                    l_str_Sql = l_str_Sql & " WHERE iActiveFlag=1))"
                Case LoadResString(1766)
                    l_str_Sql = l_str_Sql & " AND CONVERT(VARCHAR(10),dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay2,101) FROM tbSTESecondExamProfile"
                    l_str_Sql = l_str_Sql & " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile"
                    l_str_Sql = l_str_Sql & " WHERE iActiveFlag=1))"
                Case LoadResString(1767)
                    l_str_Sql = l_str_Sql & " AND CONVERT(VARCHAR(10),dtSecondExamDay,101)=(SELECT CONVERT(VARCHAR(10),dtSecondExamDay3,101) FROM tbSTESecondExamProfile"
                    l_str_Sql = l_str_Sql & " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile"
                    l_str_Sql = l_str_Sql & " WHERE iActiveFlag=1))"
                End Select
                
                If chkRoom.Value = 1 Then
                    .Col = 6
                End If
            End If
            
            l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
                        
            If Not l_obj_Rst.EOF Then
                Do While Not l_obj_Rst.EOF

                    m_str_SQl = "UPDATE tbSTEScoreProfile"
                    m_str_SQl = m_str_SQl & " SET fChoseiScore=" & l_dbl_ChooseiScore
                    m_str_SQl = m_str_SQl & ", dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'"
                    m_str_SQl = m_str_SQl & " WHERE iSubjectProfileId= " & cboSubjectId.Text
                    m_str_SQl = m_str_SQl & " AND iExamineeProfileId = " & l_obj_Rst("iExamineeProfileId")

                    Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQl)
                    
                    Set m_obj_Rst = Nothing
                    
                    l_obj_Rst.MoveNext
                Loop
                l_Bln_RecordsUpdated = True
            End If
            l_obj_Rst.Close
            Set l_obj_Rst = Nothing
        Next
    End With
    
    g_obj_Conn.CommitTrans
    If l_Bln_RecordsUpdated Then
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
    fMainForm.mnuTools.Enabled = False
    Dim index As Integer
    For index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(index).Enabled = False
    Next
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    LoadResStrings Me
    If m_int_ChoseiJoken = 1 Then
        Me.Caption = LoadResString(1012)
    Else
        Me.Caption = LoadResString(1751)
    End If
    Call g_void_SetFontProperties(Me)     ' set the font properties
    m_int_NoOfConditions = 0    ' initialise the no of conditions
    ' select all subjects that come under the selected exam type
    m_str_SQl = "SELECT iSubjectProfileId,vSubjectName FROM tbSTESubjectProfile"
    
    ' changed on 14/05/02 to incorporate choosei for Hyotei also
    If m_int_ChoseiJoken = 1 Then
        m_str_SQl = m_str_SQl & " WHERE iExamType = 0"
    ElseIf g_int_ExamType = 1 Then
        m_str_SQl = m_str_SQl & " WHERE iExamType = " & g_int_ExamType
    ElseIf g_int_ExamType = 2 Or g_int_ExamType = 3 Or g_int_ExamType = 4 Or g_int_ExamType = 5 Then
        m_str_SQl = m_str_SQl & " WHERE iExamType = 2 or iExamType = 3 or iExamType = 4 or iExamType = 5"
    End If
    
    m_str_SQl = m_str_SQl & " ORDER BY vSubjectName"
    Set m_obj_Rst = g_obj_Conn.Execute(m_str_SQl)
    cmdSubmit.Enabled = False

    If Not m_obj_Rst.EOF Then
        m_int_SelectedSubject = m_obj_Rst("iSubjectProfileId")
        ' add the subjects to combo box
        Do While Not m_obj_Rst.EOF
            cboSubject.AddItem m_obj_Rst("vSubjectName")
            cboSubjectId.AddItem m_obj_Rst("iSubjectProfileId")
            m_obj_Rst.MoveNext
        Loop
        cboSubject.ListIndex = 0
        
        If g_int_ExamType = 1 Then
            ' 1st Exam
            lblDay.Visible = False
            chkDay.Visible = False
            lblRoom.Visible = False
            chkRoom.Visible = False
            lstRooms.Visible = False
        ElseIf g_int_ExamType = 2 Or g_int_ExamType = 3 Then
            ' 2nd exam
            lblDay.Visible = True
            chkDay.Visible = True
            lblRoom.Visible = True
            chkRoom.Visible = True
            lstRooms.Visible = True
        End If
    End If
    
    ' release the object variables
    Set m_obj_Rst = Nothing
    
    Call f_void_InitGrid            ' reinitialize the grid
    Call f_void_InitRawScoreGrid    ' reinitialize the rawscore grid
    Call f_void_LoadRoom            ' Room is a checkbox now
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_InitGrid()
     vsfSearchGrid.Redraw = flexRDNone
   
    With vsfSearchGrid
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
        
        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
            ' for second exam one additional column is required for the day combo
            ' If room checkbox is checked, 2 columns for Room id and name
            If chkRoom.Value = 1 Then
                .Cols = 10
            Else
                .Cols = 8
            End If
        Else
            ' for ist exam, day column is not there, hence one column less
            .Cols = 7
        End If
        
        .Row = 0
        .Col = 0
        .ColWidth(0) = 700
        .Text = LoadResString(1756)   'Sr no  0
        .CellAlignment = flexAlignRightBottom
        
        .Col = .Col + 1
        .ColWidth(1) = 2200
        .Text = LoadResString(1757)    'subject  1
        
        .Col = .Col + 1
        .ColWidth(2) = 2000
        .Text = LoadResString(1758)  'Raw score from  2
        .CellAlignment = flexAlignRightBottom
        
        .Col = .Col + 1
        .ColWidth(3) = 2000
        .Text = LoadResString(1759)   'raw score to  3
        .CellAlignment = flexAlignRightBottom
        
        .Col = .Col + 1
        .ColWidth(4) = 1200
        .Text = LoadResString(1754)   'Sex  4
        
        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
            ' add the additional column for the day
            .Col = .Col + 1
            .ColWidth(7) = 1600
            .Text = LoadResString(1755)  'Day   Col is 5
            'new col for roomID
            If chkRoom.Value = 1 Then
                .Col = .Col + 1
                .ColWidth(6) = 0 'hidden column 6 for room Id
                .Col = .Col + 1
                .ColWidth(7) = 2000   'Column 7 for Room Desc
                .Text = LoadResString(2002)
            End If
        End If
        
        .Col = .Col + 1
        .ColWidth(.Col) = 1000  '5
        .Text = LoadResString(1760)  'Col 8 Average
        .CellAlignment = flexAlignRightBottom
        
        .Col = .Col + 1
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
        '.Font.Name = "‚l‚r ‚oƒSƒVƒbƒN"
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
        .Text = l_dbl_RawScoreFrom
        
        .Col = .Col + 1
        .Text = l_dbl_RawScoreTo
        
        .Col = .Col + 1
        If l_bln_SexFlag = 1 Then
            .Text = LoadResString(1837)
        ElseIf l_bln_SexFlag = 2 Then
            .Text = LoadResString(1838)
        Else
            .Text = LoadResString(1846)
        End If
        
        If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
            .Col = .Col + 1
            Select Case l_bln_DayFlag
            Case 0
                .Text = LoadResString(1764)
            Case 1
                .Text = LoadResString(1765)
            Case 2
                .Text = LoadResString(1766)
            Case 3
                .Text = LoadResString(1767)
            End Select
            If Not IsEmpty(l_int_RoomNo) Then
                .Col = .Col + 1
                .Text = l_int_RoomNo
                .Col = .Col + 1
                .Text = l_str_RoomName
            End If
        End If
        
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call g_void_CloseChildForm
End Sub

Private Sub vsfSearchGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    ' this code is written to round off the decimal values to 2 digits precision
    Dim l_int_ChooseiCol As Integer
    If chkRoom.Value = 1 Then
        l_int_ChooseiCol = 9
    Else
        l_int_ChooseiCol = 8
    End If
    With vsfSearchGrid
        If Trim(.TextMatrix(Row, Col)) <> "" And .Col = l_int_ChooseiCol Then
            .TextMatrix(Row, Col) = Round(.TextMatrix(Row, Col), 2)
        End If
    End With
End Sub

Private Sub vsfSearchGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfSearchGrid
        If .Redraw <> flexRDNone And Col <> vsfSearchGrid.Cols - 1 Then
            Cancel = True
            Exit Sub
        Else
            vsfSearchGrid.Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vsfSearchGrid_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfSearchGrid
        If .Redraw <> flexRDNone And NewCol <> .Cols - 1 Then
            Cancel = True
            .Select NewRow, .Cols - 1
        End If
    End With
End Sub

Private Sub vsfSearchGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If g_int_ExamType = 2 Or g_int_ExamType = 3 Then
        ' in second exam, only the 8th column is editable (choosei score)
        If Col <> IIf(chkRoom.Value = 1, 9, 7) Then
            KeyAscii = 0
        ElseIf KeyAscii = 13 Then
            If vsfSearchGrid.Row < vsfSearchGrid.Rows - 1 Then
                vsfSearchGrid.Row = vsfSearchGrid.Row + 1
                vsfSearchGrid.Col = Col
            End If
        ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
            KeyAscii = 0
        'This is to restrict user from entering more than one "." in the value
        ElseIf KeyAscii = 46 And InStr(1, vsfSearchGrid.EditText, ".") > 0 Then
            KeyAscii = 0
        End If
    ElseIf g_int_ExamType = 1 Then
        ' in first exam, only the 8th column is editable (choosei score)
        If Col <> 6 Then
            KeyAscii = 0
        ElseIf KeyAscii = 13 Then
            If vsfSearchGrid.Row < vsfSearchGrid.Rows - 1 Then
                vsfSearchGrid.Row = vsfSearchGrid.Row + 1
                vsfSearchGrid.Col = Col
            End If
        ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
            KeyAscii = 0
        'This is to restrict user from entering more than one "." in the value
        ElseIf KeyAscii = 46 And InStr(1, vsfSearchGrid.EditText, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

 Private Function f_void_GetAverage(ByVal l_dbl_RawScoreFrom As Double, ByVal l_dbl_RawScoreTo As Double) As Double
    On Error GoTo ErrorHandler
    
    m_str_SQl = "SELECT Avg(fRawScore) from tbSTEScoreProfile where fRawScore BETWEEN "
    m_str_SQl = m_str_SQl & l_dbl_RawScoreFrom & " AND " & l_dbl_RawScoreTo
    m_str_SQl = m_str_SQl & " AND iSubjectProfileId=" & cboSubjectId.Text
    
    m_obj_Rst.Open m_str_SQl, g_obj_Conn, adOpenStatic, adLockReadOnly
    
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
                                .Rows = .Rows + 1
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
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> vbKeyReturn Then
       KeyAscii = 0
    End If
End Sub

Private Function f_bln_ValidateRange() As Integer

    Dim l_int_Rows As Integer 'total rows in grid
    Dim l_int_Counter As Integer ' current row
    Dim l_bln_RetVal As Integer  ' return value
    Dim l_int_PrevColVal As Integer  'previous col value of same row
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

