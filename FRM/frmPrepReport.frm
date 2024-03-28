VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmPrepReport 
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmPrepReport.frx":0000
   ScaleHeight     =   9105
   ScaleWidth      =   12660
   WindowState     =   2  'ç≈ëÂâª
   Begin VSFlex7LCtl.VSFlexGrid msfRoomAlloc 
      Height          =   3855
      Left            =   240
      TabIndex        =   21
      Top             =   3840
      Width           =   11535
      _cx             =   20346
      _cy             =   6800
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
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
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   2
      FixedRows       =   0
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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "2494"
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
      Left            =   4560
      TabIndex        =   7
      Top             =   2985
      Width           =   3255
   End
   Begin VB.ComboBox cboRoomId 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6150
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   2565
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ComboBox cboSubjectId 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6330
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtRandomNo 
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   10125
      MaxLength       =   5
      TabIndex        =   17
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtCapacity 
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   6150
      MaxLength       =   5
      TabIndex        =   16
      Top             =   1080
      Width           =   1770
   End
   Begin VB.TextBox txtUnallocatedExaminees 
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
      Height          =   375
      Left            =   10125
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2175
      Width           =   1650
   End
   Begin VB.ComboBox cboDay 
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
      Left            =   2040
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   2
      Top             =   1080
      Width           =   1665
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
      Left            =   6150
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "1070"
      Enabled         =   0   'False
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
      Left            =   240
      TabIndex        =   6
      Top             =   7920
      Width           =   3255
   End
   Begin VB.CommandButton cmdAddRoom 
      Caption         =   "1069"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2985
      Width           =   3255
   End
   Begin VB.ComboBox cboRooms 
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
      Left            =   6120
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   0
      Top             =   2175
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtTotalExaminees 
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
      Left            =   1995
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid msfRoomAllocold 
      Height          =   3855
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      BackColor       =   16641260
      ForeColor       =   4194304
      BackColorFixed  =   16047044
      ForeColorFixed  =   8388608
      BackColorSel    =   8388608
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblRandomNo 
      Alignment       =   1  'âEëµÇ¶
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
      Height          =   360
      Left            =   8700
      TabIndex        =   18
      Top             =   1140
      Width           =   1245
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'ìßñæ
      Caption         =   "ñ ê⁄ì˙"
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
      Height          =   360
      Left            =   240
      TabIndex        =   14
      Top             =   1140
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "â»ñ⁄"
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
      Height          =   345
      Left            =   4395
      TabIndex        =   13
      Top             =   1620
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblCapacity 
      Alignment       =   1  'âEëµÇ¶
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
      Height          =   375
      Left            =   4410
      TabIndex        =   11
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Label lblRoomNo 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "âÔèÍî‘çÜ"
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
      Height          =   300
      Left            =   3870
      TabIndex        =   10
      Top             =   2220
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Label lblTotalNoOfExaminees 
      BackStyle       =   0  'ìßñæ
      Caption         =   "çáåvéÛå±é“êî"
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
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1620
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "ñ¢êUï™ÇØéÛå±ê∂"
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
      Height          =   300
      Left            =   7995
      TabIndex        =   8
      Top             =   2235
      Width           =   1950
   End
   Begin VB.Label lblErrorDetails 
      BackStyle       =   0  'ìßñæ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   240
      TabIndex        =   12
      Top             =   3420
      Width           =   11520
   End
End
Attribute VB_Name = "frmPrepReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmPrepReport
'Author         :   Dileep Cherian
'Created On     :   17/05/2002
'Description    :   This form is used to allocate the examinees who are eligible for
'                   for the second phase report exams
'Reference      :   FunctionalSpecs OF Preparation of Report.doc(ver 1.0)
'***************************************************************************************************

Dim m_int_ToBeAllotted As Long       ' total number of examinees remaining to be allotted
Dim m_int_TotalExaminees  As Long    ' total number of eligible examinees for this phase of the examinees
Dim m_int_SrNo As Long               ' assign serial number to the grid
Dim f_dt_dtDay As Date                  ' the selected day of report
Dim f_int_NoOfRooms As Long          ' total number of room available for the selected day
Dim f_int_NoOfExaminee As Long       ' total number of examinees alowed for the selected day
Dim f_str_AddedJuken As String          ' list of already added juken
Dim f_int_Day1Allocated As Long      ' number of examinees allocated for day1
Dim f_int_Day2Allocated As Long      ' number of examinees allocated for day2
Dim f_int_Day3Allocated As Long      ' number of examinees allocated for day3
Dim f_bln_DataChanged As Boolean        ' to see whether data changed for the current subject
Private prvsCurSerial As String 'ëIëÅAì¸óÕíÜÇÃí˘ê≥ëOÉVÉäÉAÉãî‘çÜ

Private prvlSerialCol As Long
Private prvlRoomNameCol As Long
Private prvlRandomNoCol As Long
Private prvlCapacityCol As Long
Private prvlSubjectCol As Long
Private prvlDayCol As Long
Private prvlStartNoCol As Long
Private prvlEndNoCol As Long

Private Sub f_void_InitGrid()
    ' initializes the grid with it sheaders, column width etc
    With msfRoomAlloc
        .Visible = False
        .BackColor = &HFFFFFF
        .BackColorBkg = &HFFFFFF
        .BackColorFixed = &H8000000F
        .BackColorSel = &H800000
        .FixedCols = 0
        .TextStyleFixed = flexTextFlat
        .Font.Bold = False
        .ForeColorFixed = &H80000008
        .ForeColor = &H800000
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .GridColor = &H808080
        .AllowUserResizing = flexResizeColumns
        .Visible = True
    
        .Rows = 1
        .cols = 8                                   ' fix the number of columns to be displayed
        
        .FixedRows = 1
        .FixedCols = 0
        
        .Row = 0
        .Col = 0
        prvlSerialCol = .Col
        .ColWidth(0) = 972
        .CellAlignment = flexAlignRightBottom
        .Text = LoadResString(1756)                 ' serial number

        .Col = 1
        prvlRoomNameCol = .Col
        .ColWidth(1) = 0
        .Text = LoadResString(1503)                 ' room name
        
        .Col = 2
        prvlRandomNoCol = .Col
        .ColWidth(2) = 1400
        .CellAlignment = flexAlignRightBottom
        .Text = LoadResString(1504)                 ' random number of room
        
        .Col = 3
        prvlCapacityCol = .Col
        .ColWidth(3) = 1600
        .CellAlignment = flexAlignRightBottom
        .Text = LoadResString(1505)                 ' max capacity of room
        
        .Col = 4
        prvlSubjectCol = .Col
        .ColWidth(4) = 0
        .Text = LoadResString(1753)                 ' selected subject
        
        .Col = 5
        prvlDayCol = .Col
        .ColWidth(5) = 1600
        .CellAlignment = flexAlignRightBottom
        .Text = LoadResString(1755)                 ' selected day
        
        .Col = 6
        prvlStartNoCol = .Col
        .ColWidth(6) = 2000
        .Text = LoadResString(1803) & "é©"                 ' allotted examinee's start juken numbers
        
        .Col = 7
        prvlEndNoCol = .Col
        .ColWidth(7) = 2000
        .Text = LoadResString(1803) & "éä"                 ' allotted examinee's end juken number
    End With
End Sub

Private Sub cboDay_Click()
    On Error GoTo ErrorHandler

'ÉOÉäÉbÉhï\é¶ÇÇ∑ÇÈ
    Call f_void_PopulateGrid

    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cboRooms_Click()
    ' get the details of the room like max capacity and random number and & _
        display in the corresponding text boxes
    ' also identify the rooms already allocated for interview and for those rooms & _
        the max capacity and random number should not be allowed to change
    Dim l_str_sqlRoom As String                 ' SQL string to pick up room details
    Dim l_obj_RsRoom As New ADODB.Recordset     ' recordset variable to pick up room details
    Dim l_str_RoomIds As String                 ' store the list of rooms allocated for interview
               
    On Error GoTo ErrorHandler
    
    cboRoomId.ListIndex = cboRooms.ListIndex
     
    l_str_sqlRoom = "SELECT iRoomProfileId, iRandom, iMaxCapacity FROM tbSTERoomProfile WHERE vRoomName='" & cboRooms.Text & "'"
    Set l_obj_RsRoom = g_obj_Conn.Execute(l_str_sqlRoom)

    If IsNull(l_obj_RsRoom("iMaxCapacity")) And m_int_ToBeAllotted = 0 Then
        lblErrorDetails.Caption = LoadResString(2011)   ' all examinees are allotted
        Exit Sub
    End If
    
    If Not l_obj_RsRoom.EOF Then
        
        ' display the random number
        If Trim(l_obj_RsRoom.Fields("iRandom").Value) <> "" Then
            txtRandomNo.Text = l_obj_RsRoom.Fields("iRandom").Value
        Else
            txtRandomNo.Text = 0
        End If
           
        ' display the max capacity
        If Trim(l_obj_RsRoom.Fields("iMaxCapacity").Value) <> "" Then
            txtCapacity.Text = l_obj_RsRoom.Fields("iMaxCapacity").Value
        Else
            txtCapacity.Text = 0
        End If
    End If
    
    l_obj_RsRoom.Close
    Set l_obj_RsRoom = Nothing
    
    lblErrorDetails.Caption = ""    ' clear the error label once the room is changed
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cboSubject_Click()
    ' display the new new grid for the selected subject
    ' before that, user should be prompted for any changes for the current grid
    Dim l_int_Ans As Long        ' to get the user response
    On Error GoTo ErrorHandler
    
'    lblErrorDetails.Caption = ""
'    If f_bln_DataChanged Then   ' if any changes are made to the grid, the ask whether to save or not
'        l_int_Ans = MsgBox(LoadResString(1118), vbQuestion + vbYesNo, LoadResString(1729))
'        If l_int_Ans = vbYes Then
'             ' save the data
'            Call cmdFinish_Click
'        End If
'        f_bln_DataChanged = False
'    End If
    cboSubjectId.ListIndex = cboSubject.ListIndex
'    Call f_void_GetAllocation   ' get allocation for the new subject
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdAddRoom_Click()
    ' check all the validations
    ' update any changes in max capacity or random number of the selected room
    ' allocate examinees to the selected room for the slected subject & _
        on the selected day
        
Dim sSQL As String                 ' SQL string variable
Dim oRs As New ADODB.Recordset    ' recordset variable

Dim lRow As Long
Dim sRandomNo As String

    On Error GoTo ErrorHandler

    lblErrorDetails.Caption = ""
    ' validate random number
    If Trim(txtRandomNo.Text) <> "" And IsNumeric(txtRandomNo.Text) Then
        'ëºÇÃì˙Ç…ìØÇ∂óêêîÇÕê›íËÇ≈Ç´Ç»Ç¢
        sSQL = "SELECT top 1 iJukenNumber FROM tbSTEExamineeProfile" & _
            " WHERE iShoronbunRandomNo=" & CInt(Trim(txtRandomNo.Text)) & _
            " AND iNendo = " & g_int_CurrentNendo

        Set oRs = g_obj_Conn.Execute(sSQL)
        If Not oRs.EOF Then
            ' this random number already exists for another room
            lblErrorDetails.Caption = LoadResString(2006)
            txtRandomNo.SetFocus
            txtRandomNo.SelStart = 0
            txtRandomNo.SelLength = Len(txtRandomNo.Text)
            Exit Sub
        End If
        oRs.Close
        Set oRs = Nothing
        sRandomNo = Trim(txtRandomNo.Text)
        For lRow = 1 To msfRoomAlloc.Rows - 1
            If sRandomNo = msfRoomAlloc.TextMatrix(lRow, 2) Then
                ' this random number already exists for another room
                lblErrorDetails.Caption = LoadResString(2006)
                txtRandomNo.SetFocus
                txtRandomNo.SelStart = 0
                txtRandomNo.SelLength = Len(txtRandomNo.Text)
                Exit Sub
            End If
        Next
    Else
        lblErrorDetails.Caption = LoadResString(2007)
        txtRandomNo.SetFocus
        Exit Sub
    End If
    
    If Trim(txtCapacity.Text) = "" Then
        ' capacity cannot be null
        lblErrorDetails.Caption = LoadResString(2008)
        txtCapacity.SetFocus
        Exit Sub
    End If

Dim sWk As String

    sWk = Trim(str(msfRoomAlloc.Rows))
    sWk = sWk & vbTab & "" 'ïîâÆñºèÃÇÕÇ»Ç¢
    sWk = sWk & vbTab & Trim(txtRandomNo.Text)
    sWk = sWk & vbTab & Trim(txtCapacity.Text)
    sWk = sWk & vbTab & "" 'â»ñ⁄ñºèÃÇÕÇ»Ç¢
    sWk = sWk & vbTab & Trim(cboDay.Text)

    sSQL = "SELECT isnull(min( iJukenNumber ),-1) as smJukenNumber ," & _
        " max( iJukenNumber ) as mxJukenNumber " & _
        " FROM ( SELECT top " & Trim(txtCapacity.Text) & _
        " iJukenNumber FROM tbSTEExamineeProfile" & _
        " WHERE iNendo=" & g_int_CurrentNendo & _
        " AND iExamineeStatus = " & gclExamineeStatus_1stPass & " AND iAbsentFlag = 0" & _
        " AND dtSecondExamDay = ( select top 1 dtSecondExamDay" & Trim(str(cboDay.ListIndex + 1)) & _
        "                         from tbSTESecondExamProfile as se " & _
        "                         where exists ( select 1 from tbSTESystemProfile as sp where sp.iSystemProfileId = se.iSystemProfileId ) ) " & _
        " AND iJukenNumber > " & Trim(str(IIf(msfRoomAlloc.Rows = 1, 0, msfRoomAlloc.TextMatrix(msfRoomAlloc.Rows - 1, 7)))) & " ORDER BY iJukenNumber ) as t1 "

    Set oRs = g_obj_Conn.Execute(sSQL)

    If oRs.EOF Then
    'äÑÇËìñÇƒÇÈêlÇ™Ç¢Ç»Ç¢
        lblErrorDetails.Caption = LoadResString(2008)
        txtCapacity.SetFocus
        Exit Sub
    End If
    If oRs.Fields(0) = -1 Then
    'äÑÇËìñÇƒÇÈêlÇ™Ç¢Ç»Ç¢
        lblErrorDetails.Caption = LoadResString(2008)
        txtCapacity.SetFocus
        Exit Sub
    End If

    sWk = sWk & vbTab & Trim(str(oRs.Fields(0)))
    sWk = sWk & vbTab & Trim(str(oRs.Fields(1)))

    oRs.Close
    Set oRs = Nothing

    msfRoomAlloc.Rows = msfRoomAlloc.Rows + 1
    msfRoomAlloc.Row = msfRoomAlloc.Rows - 1
    msfRoomAlloc.RowSel = msfRoomAlloc.Rows - 1
    msfRoomAlloc.Col = 0
    msfRoomAlloc.ColSel = msfRoomAlloc.cols - 1
    msfRoomAlloc.Clip = sWk

    Call lsGetUnallocatedExamineeCnt

    cmdFinish.Enabled = True

    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdDelete_Click()

Dim lDelRow As Long
Dim sSQL As String
Dim oRs As New ADODB.Recordset    ' recordset variable

Dim lRow As Long
Dim sRandomNo As String

    On Error GoTo ErrorHandler

    lblErrorDetails.Caption = ""
    If msfRoomAlloc.Rows <= 1 Then Exit Sub         ' exit if there are no rows in the grid
    If msfRoomAlloc.Row < 1 Then Exit Sub         ' exit if there are no rows in the grid

    ' confirm the deletion
    If MsgBox(LoadResString(1122), vbQuestion + vbYesNo) = vbNo Then Exit Sub  ' don't delete - exit

    lDelRow = msfRoomAlloc.Row
    msfRoomAlloc.RemoveItem lDelRow

    If lDelRow < msfRoomAlloc.Rows Then

        For lRow = lDelRow To msfRoomAlloc.Rows - 1

            sSQL = "SELECT isnull(min( iJukenNumber ),-1) as smJukenNumber ," & _
                " max( iJukenNumber ) as mxJukenNumber " & _
                " FROM ( SELECT top " & msfRoomAlloc.TextMatrix(lRow, 3) & _
                " iJukenNumber FROM tbSTEExamineeProfile" & _
                " WHERE iNendo=" & g_int_CurrentNendo & _
                " AND iExamineeStatus = " & gclExamineeStatus_1stPass & " AND iAbsentFlag = 0" & _
                " AND dtSecondExamDay = ( select top 1 dtSecondExamDay" & Trim(str(cboDay.ListIndex + 1)) & _
                "                         from tbSTESecondExamProfile as se " & _
                "                         where exists ( select 1 from tbSTESystemProfile as sp where sp.iSystemProfileId = se.iSystemProfileId ) ) " & _
                " AND iJukenNumber > " & Trim(str(IIf(lRow = 1, 0, msfRoomAlloc.TextMatrix(lRow - 1, 7)))) & " ORDER BY iJukenNumber ) as t1 "
    
            Set oRs = g_obj_Conn.Execute(sSQL)

            msfRoomAlloc.TextMatrix(lRow, 0) = Trim(str(lRow))
            msfRoomAlloc.TextMatrix(lRow, 6) = Trim(str(oRs.Fields(0)))
            msfRoomAlloc.TextMatrix(lRow, 7) = Trim(str(oRs.Fields(1)))

            oRs.Close
            Set oRs = Nothing
        Next
    Else
        cmdFinish.Enabled = True
    End If

    Call lsGetUnallocatedExamineeCnt

    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdFinish_Click()

'DBÇ÷è¨ò_ï∂êUÇËï™ÇØÇìoò^Ç∑ÇÈ
'tbSTEInterviewRoomProfileÇ∆tbSTEExamineeÇçXêV

Dim sSQL As String
Dim lRow As Long

On Error GoTo ErrorHandler

    With msfRoomAlloc

        g_obj_Conn.BeginTrans

'äYìñì˙ÇÃè¨ò_ï∂ÇÃÉOÉãÅ[ÉvèÓïÒÇÉNÉäÉA
        sSQL = "Delete From tbSTEInterviewRoomProfile "
        sSQL = sSQL & " WHERE iNendo=" & g_int_CurrentNendo
        sSQL = sSQL & " AND iDayFlag = " & Trim(str(cboDay.ListIndex))
        sSQL = sSQL & " AND iRandomNo is not null "
'        sSQL = sSQL & " AND dtSecondExamDay = ( SELECT top 1 dtSecondExamDay" & Trim(str(cboDay.ListIndex + 1))
'        sSQL = sSQL & "                           FROM tbSTESecondExamProfile as se "
'        sSQL = sSQL & "                          WHERE exists ( select 1 from tbSTESystemProfile as sp where sp.iSystemProfileId = se.iSystemProfileId ) )"

        g_obj_Conn.Execute sSQL

'äYìñì˙ÇÃéÛå±ê∂ÇÃè¨ò_ï∂ÇÃÉOÉãÅ[ÉvèÓïÒÇÉNÉäÉA
        sSQL = "Update tbSTEExamineeProfile "
        sSQL = sSQL & " SET iShoronbunRandomNo = null "
        sSQL = sSQL & " WHERE iNendo=" & g_int_CurrentNendo
        sSQL = sSQL & " AND dtSecondExamDay = ( SELECT top 1 dtSecondExamDay" & Trim(str(cboDay.ListIndex + 1))
        sSQL = sSQL & "                           FROM tbSTESecondExamProfile as se "
        sSQL = sSQL & "                          WHERE exists ( select 1 from tbSTESystemProfile as sp where sp.iSystemProfileId = se.iSystemProfileId ) )"

        g_obj_Conn.Execute sSQL

        For lRow = 1 To .Rows - 1

            If .TextMatrix(lRow, prvlStartNoCol) = "-" Then Exit For

'äYìñì˙ÇÃè¨ò_ï∂ÇÃóêêîÇìoò^
            sSQL = "Insert Into tbSTEInterviewRoomProfile ( "
            sSQL = sSQL & "  iInterviewRoomProfileId "
            sSQL = sSQL & " ,iNendo "
            sSQL = sSQL & " ,iRandomNo "
            sSQL = sSQL & " ,iSubjectProfileId "
            sSQL = sSQL & " ,iDayFlag "
            sSQL = sSQL & " ) "
            sSQL = sSQL & " SELECT isnull( max( iInterviewRoomProfileId ) + 1 , 1 ) "
            sSQL = sSQL & " , " & g_int_CurrentNendo
            sSQL = sSQL & " , " & Trim(txtRandomNo.Text)
            sSQL = sSQL & " , " & cboSubjectId.Text
            sSQL = sSQL & " , " & Trim(str(cboDay.ListIndex))
            sSQL = sSQL & " FROM tbSTEInterviewRoomProfile "

            g_obj_Conn.Execute sSQL

            sSQL = "Update tbSTEExamineeProfile "
            sSQL = sSQL & " Set iShoronbunRandomNo = " & .TextMatrix(lRow, 2)
            sSQL = sSQL & " WHERE iNendo=" & g_int_CurrentNendo
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass & " AND iAbsentFlag = 0"
            sSQL = sSQL & " AND dtSecondExamDay = ( SELECT top 1 dtSecondExamDay" & Trim(str(cboDay.ListIndex + 1))
            sSQL = sSQL & "                           FROM tbSTESecondExamProfile as se "
            sSQL = sSQL & "                          WHERE exists ( select 1 from tbSTESystemProfile as sp where sp.iSystemProfileId = se.iSystemProfileId ) )"
            sSQL = sSQL & " AND iJukenNumber >= " & Trim(str(.TextMatrix(lRow, prvlStartNoCol)))
            sSQL = sSQL & " AND iJukenNumber <= " & Trim(str(.TextMatrix(lRow, prvlEndNoCol)))

            g_obj_Conn.Execute sSQL

        Next

        g_obj_Conn.CommitTrans

    End With

    lblErrorDetails.Caption = "çXêVäÆóπÇµÇ‹ÇµÇΩ"
    cmdFinish.Enabled = False

Exit Sub

ErrorHandler:
    g_obj_Conn.RollbackTrans
    MsgBox Err.Description, vbInformation, LoadResString(1729)

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
    Dim l_str_sqlExaminee As String                 ' get the examinee details
    Dim l_obj_rsExaminee As New ADODB.Recordset     ' recordset to hold the examinee details
    
    On Error GoTo ErrorHandler
    
    ' assumption made that interview 1 has to happen for the report to happen
    l_str_sqlExaminee = "SELECT a.iExamineeRoomProfileId FROM tbSTEExamineeRoomProfile a, tbSTEExamineeProfile b" & _
        " WHERE a.iSubjectProfileId in (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE iSubType=3)" & _
        " AND b.iAbsentFlag = 0" & _
        " AND a.iExamineeProfileId = b.iExamineeProfileId"
    
    l_obj_rsExaminee.Open l_str_sqlExaminee, g_obj_Conn, adOpenStatic, adLockReadOnly
    m_int_TotalExaminees = l_obj_rsExaminee.RecordCount
    
    l_obj_rsExaminee.Close
    Set l_obj_rsExaminee = Nothing
    
    If m_int_TotalExaminees <= 0 Then
        ' interview 1 has not occures - exit
        g_bln_InterviewHappened = False
        MsgBox LoadResString(2497), vbInformation, LoadResString(1729)
        Exit Sub
    End If
    
    ' interview 1 has already occured - continue loading the form
    LoadResStrings Me
    Me.Caption = LoadResString(2433)
    g_void_SetFontProperties Me
    
    cboRoomId.Visible = False
    cboSubjectId.Visible = False
    cmdDelete.Enabled = False
                
    g_bln_InterviewHappened = True
'    Call f_void_AddRooms                                ' populate the room combo
    Call f_void_InitGrid                       ' call global function that initializes the grid
    
    If m_int_TotalExaminees > 0 Then
        txtTotalExaminees.Text = m_int_TotalExaminees
    Else
        ' no examinees left to be all0cated
        txtTotalExaminees.Text = "0"
        lblErrorDetails.Caption = LoadResString(2009)
        cmdAddRoom.Enabled = False
        cmdFinish.Enabled = False
    End If
    Call l_void_PopulateSubject         ' populate the subject combo
        
'    With cboDay                         ' populate the day combo
'        .AddItem LoadResString(2424)
'        .AddItem LoadResString(2425)
'        .AddItem LoadResString(2426)
'        .ListIndex = 0
'    End With
    Call l_void_PopulateDayCombo

    Exit Sub
ErrorHandler:
   MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub l_void_PopulateDayCombo()

Dim sSQL As String
Dim oRs As ADODB.Recordset
Dim bThirdDay As Boolean

    bThirdDay = False

    sSQL = "SELECT dtSecondExamDay3 FROM tbSTESecondExamProfile "
    sSQL = sSQL & " WHERE iSystemProfileId = "
    sSQL = sSQL & " (SELECT iSystemProfileId FROM tbSTESystemProfile WHERE iActiveFlag=1) "

    Set oRs = g_obj_Conn.Execute(sSQL)

    If Not oRs.EOF Then
        If Not IsNull(oRs.Fields(0)) Then
            bThirdDay = True
        End If
    End If

    oRs.Close
    Set oRs = Nothing

    With cboDay
        .Clear
        .AddItem LoadResString(2424)
        .AddItem LoadResString(2425)
        If bThirdDay Then .AddItem LoadResString(2426)
        .ListIndex = 0
    End With

End Sub

Private Sub f_void_AddRooms()
    ' populate the rooms combo
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
        
    ' change made on 31/07/02
    l_str_Sql = "SELECT iRoomProfileId, vRoomName FROM tbSTERoomProfile" & _
        " WHERE iMaxCapacity > 0" & _
        " AND iInterviewRoomFlag = 1" & _
        " ORDER BY iRoomProfileId"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    Do While Not l_obj_Rst.EOF
        cboRoomId.AddItem l_obj_Rst.Fields("iRoomProfileId").Value
        cboRooms.AddItem l_obj_Rst.Fields("vRoomName").Value
        l_obj_Rst.MoveNext
    Loop
    If cboRooms.ListCount > 0 Then
        cboRooms.ListIndex = 0
    Else
        lblErrorDetails.Caption = LoadResString(2010)
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
End Sub

Private Sub l_void_PopulateSubject()
    ' populate the subjects combo
    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
        
    l_str_Sql = "SELECT iSubjectProfileId, vSubjectName FROM tbSTESubjectProfile" & _
        " WHERE iSubType = 4 "

    With l_obj_Rst
        .Open l_str_Sql, g_obj_Conn
        cboSubject.Clear
        
        Do While Not .EOF
            cboSubjectId.AddItem .Fields("iSubjectProfileId").Value
            cboSubject.AddItem .Fields("vSubjectName").Value
            .MoveNext
        Loop
        If cboSubject.ListCount > 0 Then cboSubject.ListIndex = 0
    End With
    
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
End Sub

Private Sub f_void_PopulateGrid()

Dim oRs As ADODB.Recordset
Dim sSQL As String
Dim lRow As Long

On Error GoTo ErrorHandler

    msfRoomAlloc.Rows = 1

    sSQL = "SELECT isnull(iShoronbunRandomNo,-1) " & _
        " , count(*) as iRoomCnt " & _
        " , dtSecondExamDay " & _
        " , min(iJukenNumber) as imiJNo " & _
        " , max(iJukenNumber) as imxJNo " & _
        " FROM tbSTEExamineeProfile" & _
        " WHERE iNendo=" & g_int_CurrentNendo & _
        " AND iExamineeStatus = " & gclExamineeStatus_1stPass & " AND iAbsentFlag = 0" & _
        " AND dtSecondExamDay = ( select top 1 dtSecondExamDay" & Trim(str(cboDay.ListIndex + 1)) & _
        "                         from tbSTESecondExamProfile as se " & _
        "                         where exists ( select 1 from tbSTESystemProfile as sp where sp.iSystemProfileId = se.iSystemProfileId ) ) " & _
        " Group by iShoronbunRandomNo" & _
        "        , dtSecondExamDay" & _
        " Order by iShoronbunRandomNo"

    Set oRs = g_obj_Conn.Execute(sSQL)

    lRow = 1

Dim sWk As String

    Do Until oRs.EOF

        If oRs.Fields(0) = -1 Then Exit Do

        sWk = Trim(str(lRow))
        sWk = sWk & vbTab & "" 'ïîâÆñºèÃÇÕÇ»Ç¢
        sWk = sWk & vbTab & Trim(str(oRs.Fields(0))) 'óêêî
        sWk = sWk & vbTab & Trim(str(oRs.Fields(1))) 'íËàıÇ…ÇÕìoò^êlêîÇï\é¶
        sWk = sWk & vbTab & "" 'â»ñ⁄ñºèÃÇÕÇ»Ç¢
        sWk = sWk & vbTab & Trim(str(oRs.Fields(2))) 'ééå±ì˙

        sWk = sWk & vbTab & Trim(str(oRs.Fields(3)))
        sWk = sWk & vbTab & Trim(str(oRs.Fields(4)))

        msfRoomAlloc.Rows = lRow + 1
        msfRoomAlloc.Row = lRow
        msfRoomAlloc.RowSel = lRow
        msfRoomAlloc.Col = 0
        msfRoomAlloc.ColSel = msfRoomAlloc.cols - 1
        msfRoomAlloc.Clip = sWk

        lRow = lRow + 1

        oRs.MoveNext

    Loop

    Call lsGetTotalExamineeCnt
    Call lsGetUnallocatedExamineeCnt

Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub f_void_GetAllocation()
    Dim l_int_SrNo As Long                       ' to set the serial number of grid
    Dim l_str_sqlExaminee As String                 ' for the examinee sql string
    Dim l_obj_rsExaminee As New ADODB.Recordset     ' recordset for examinee details
    Dim l_str_sqlRooms As String                    ' for the room sql string
    Dim l_obj_rsRooms As New ADODB.Recordset        ' recordset for room details
    Dim l_str_sqlDays As String                     ' sql string to get the day details
    Dim l_obj_rsDays As New ADODB.Recordset         ' recordset to get the day details
    Dim l_str_sqlExamDays As String                 ' form sqlstring for the exam days
    Dim l_obj_rsExamDays As New ADODB.Recordset     ' recordset for exam days
    Dim l_dt_day1 As Date                           ' store the day1 date
    Dim l_dt_day2 As Date                           ' store the day2 date
    Dim l_dt_day3 As Date                           ' store the day3 date
    Dim l_str_Examinees As String                   ' to store the list of examinees
    Dim l_str_sqlSubject As String                  ' to store the list of subjects
    Dim l_obj_rsSubject As New ADODB.Recordset      ' recordset fore the subject details
    
    ' re-initialize the formlevel variables
    f_int_Day1Allocated = 0
    f_int_Day2Allocated = 0
    f_int_Day3Allocated = 0
    f_str_AddedJuken = ""
    
    ' pick up the exam days
    l_str_sqlExamDays = "SELECT dtSecondExamDay1,dtSecondExamDay2,dtSecondExamDay3" & _
        " FROM tbSTESecondExamProfile" & _
        " WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag=1)"
    l_obj_rsExamDays.Open l_str_sqlExamDays, g_obj_Conn
     
    If Not l_obj_rsExamDays.EOF Then
        l_dt_day1 = Format(l_obj_rsExamDays.Fields("dtSecondExamDay1").Value, "MM/DD/YYYY")
        l_dt_day2 = Format(l_obj_rsExamDays.Fields("dtSecondExamDay2").Value, "MM/DD/YYYY")
        l_dt_day3 = Format(l_obj_rsExamDays.Fields("dtSecondExamDay3").Value, "MM/DD/YYYY")
    Else
        Exit Sub
    End If
    l_obj_rsExamDays.Close
    Set l_obj_rsExamDays = Nothing
    
    msfRoomAlloc.Rows = 2
    
    ' pick the room details
    ' change made on 31/07/02
    l_str_sqlRooms = "SELECT iRoomProfileId, vRoomName, iMaxCapacity, iRandom FROM tbSTERoomProfile" & _
        " WHERE iRoomProfileId IN(SELECT DISTINCT iRoomProfileId FROM tbSTEExamineeRoomProfile" & _
        " WHERE iRoomProfileId IS NOT NULL)" & _
        " AND iInterviewRoomFlag = 1" & _
        " ORDER BY iRoomProfileId"
    l_obj_rsRooms.Open l_str_sqlRooms, g_obj_Conn
    
    Do While Not l_obj_rsRooms.EOF
        ' loop through all exam days
        l_str_sqlDays = "SELECT DISTINCT dtSecondExamDay FROM tbSTEExamineeProfile" & _
            " WHERE dtSecondExamDay IS NOT NULL" & _
            " AND iNendo=" & g_int_CurrentNendo
        l_obj_rsDays.Open l_str_sqlDays, g_obj_Conn
        
        Do While Not l_obj_rsDays.EOF
            ' pick all the eligible examinees
            l_str_sqlExaminee = "SELECT a.iExamineeProfileId, b.iSubjectProfileId FROM tbSTEExamineeProfile a, tbSTEExamineeRoomProfile b" & _
                " WHERE a.iExamineeProfileId = b.iExamineeProfileId" & _
                " AND a.dtSecondExamDay='" & Format(l_obj_rsDays.Fields("dtSecondExamDay").Value, "MM/DD/YYYY") & "'" & _
                " AND b.iRoomProfileId=" & l_obj_rsRooms.Fields("iRoomProfileId").Value & _
                " AND a.iNendo=" & g_int_CurrentNendo & _
                " AND a.iAbsentFlag = 0" & _
                " AND a.iExamineeStatus = " & gclExamineeStatus_1stPass & _
                " AND b.iSubjectProfileId=" & cboSubjectId.Text
                
            l_obj_rsExaminee.Open l_str_sqlExaminee, g_obj_Conn, adOpenStatic, adLockReadOnly
            
            If Not l_obj_rsExaminee.EOF Then
                
                With msfRoomAlloc
                
                l_int_SrNo = l_int_SrNo + 1
                .Row = l_int_SrNo
    
                .Col = 0
                .Text = l_int_SrNo
    
                .Col = 1
                .Text = l_obj_rsRooms.Fields("vRoomName").Value
    
                .Col = 2
                .Text = l_obj_rsRooms.Fields("iRandom").Value
    
                .Col = 3
                .Text = l_obj_rsRooms.Fields("iMaxCapacity").Value
    
                .Col = 4    ' get subject name
                l_str_sqlSubject = "SELECT vSubjectName FROM tbSTESubjectProfile" & _
                    " WHERE iSubjectProfileId=" & l_obj_rsExaminee.Fields("iSubjectProfileId").Value
                l_obj_rsSubject.Open l_str_sqlSubject, g_obj_Conn
                
                If Not l_obj_rsSubject.EOF Then
                    .Text = l_obj_rsSubject.Fields("vSubjectName").Value
                Else
                    .Text = ""
                End If
                l_obj_rsSubject.Close
                Set l_obj_rsSubject = Nothing
                
                .Col = 5    ' get a count of examinees getting allocated for each day
                Select Case Format(l_obj_rsDays.Fields("dtSecondExamDay").Value, "MM/DD/YYYY")
                Case l_dt_day1          ' day1
                    .Text = LoadResString(2424)
                    f_int_Day1Allocated = f_int_Day1Allocated + l_obj_rsExaminee.RecordCount
                Case l_dt_day2          ' day2
                    .Text = LoadResString(2425)
                    f_int_Day2Allocated = f_int_Day2Allocated + l_obj_rsExaminee.RecordCount
                Case l_dt_day3          ' day3
                    .Text = LoadResString(2426)
                    f_int_Day3Allocated = f_int_Day3Allocated + l_obj_rsExaminee.RecordCount
                End Select
                
                l_str_Examinees = ""
                Do While Not l_obj_rsExaminee.EOF
                    l_str_Examinees = l_str_Examinees & l_obj_rsExaminee.Fields("iExamineeProfileId").Value & ","
                    f_str_AddedJuken = f_str_AddedJuken & l_obj_rsExaminee.Fields("iExamineeProfileId").Value & ","
                    l_obj_rsExaminee.MoveNext
                Loop
                
                If Trim(l_str_Examinees) <> "" Then
                    l_str_Examinees = Left(l_str_Examinees, Len(Trim(l_str_Examinees)) - 1)
                    .Col = 6
                    .CellAlignment = flexAlignLeftCenter
                    .Text = l_str_Examinees
                Else
                    .Col = 6
                    .CellAlignment = flexAlignLeftCenter
                    .Text = ""
                End If
                
                .Rows = .Rows + 1
                
                 End With
            End If
            
            l_obj_rsExaminee.Close
            Set l_obj_rsExaminee = Nothing
           
            l_obj_rsDays.MoveNext
        Loop
        
        l_obj_rsDays.Close
        Set l_obj_rsDays = Nothing
        
        l_obj_rsRooms.MoveNext
    Loop
    
    If Trim(f_str_AddedJuken) <> "" Then
        f_str_AddedJuken = Left(f_str_AddedJuken, Len(f_str_AddedJuken) - 1)
    End If
    
    msfRoomAlloc.Rows = msfRoomAlloc.Rows - 1
    l_obj_rsRooms.Close
    Set l_obj_rsRooms = Nothing
    Call cboDay_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    f_bln_DataChanged = False
    f_str_AddedJuken = ""
    f_int_Day1Allocated = 0
    f_int_Day2Allocated = 0
    f_int_Day3Allocated = 0
    
    Set frmPrepReport = Nothing
    Unload Me
End Sub


Private Sub msfRoomAlloc_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim sSerial As String
Dim lRow As Long
Dim lCurRow As Long
Dim lChangeRow As Long
Dim sCurrent As String
Dim sChange As String

    With msfRoomAlloc
        Select Case Col
        Case prvlSerialCol
            sSerial = .TextMatrix(Row, Col)
            If Not IsNumeric(sSerial) Then
                MsgBox "êîílÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", vbOKOnly Or vbInformation, "ì¸óÕÉGÉâÅ["
                .TextMatrix(Row, Col) = prvsCurSerial
                Exit Sub
            End If
    'ì¸óÕÇ≥ÇÍÇΩÉVÉäÉAÉãî‘çÜÇÉOÉäÉbÉhÇ©ÇÁíTÇ∑
            lChangeRow = -1
            For lRow = 1 To .Rows - 1
                If lRow <> Row And sSerial = .TextMatrix(lRow, 0) Then
                    lChangeRow = lRow
                    Exit For
                End If
            Next
            If lChangeRow <> -1 Then
                lCurRow = Row
                .Row = lCurRow
                .Col = 1
                .ColSel = .cols - 1
                sCurrent = sSerial & vbTab & .Clip
                .Row = lChangeRow
                .Col = 1
                .ColSel = .cols - 1
                sChange = prvsCurSerial & vbTab & .Clip
                .Row = lCurRow
                .Col = 0
                .ColSel = .cols - 1
                .Clip = sChange
                .Row = lChangeRow
                .Col = 0
                .ColSel = .cols - 1
                .Clip = sCurrent
        'éÛå±î‘çÜîÕàÕÇçƒåvéZÇ∑ÇÈ
                Call reCountRoomExaminee(IIf(lCurRow < lChangeRow, lCurRow, lChangeRow))
    '        Call cmdAddRoom_Click
                cmdFinish.Enabled = True
            End If
        Case prvlCapacityCol
            sSerial = .TextMatrix(Row, Col)
            If Not IsNumeric(sSerial) Then
                MsgBox "êîílÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", vbOKOnly Or vbInformation, "ì¸óÕÉGÉâÅ["
                .TextMatrix(Row, Col) = prvsCurSerial
                Exit Sub
            End If
            Call reCountRoomExaminee(.Row)
            Call lsGetUnallocatedExamineeCnt
        End Select
    End With

End Sub

Private Sub reCountRoomExaminee(plCurRow As Long)
'éÛå±î‘çÜîÕàÕÇçƒåvéZÇ∑ÇÈ
Dim sSQL As String
Dim oRs As ADODB.Recordset
Dim lStartNo As Long
Dim lRow As Long
Dim lRow2 As Long
    With msfRoomAlloc
        For lRow = plCurRow To .Rows - 1
            lStartNo = CLng(IIf(IsNumeric(.TextMatrix(lRow - 1, 7)), .TextMatrix(lRow - 1, 7), 0))
            sSQL = "SELECT min( iJukenNumber ) as smJukenNumber ,"
            sSQL = sSQL & " max( iJukenNumber ) as mxJukenNumber "
            sSQL = sSQL & " FROM ( SELECT top " & Trim(str(.TextMatrix(lRow, prvlCapacityCol)))
            sSQL = sSQL & " iJukenNumber FROM tbSTEExamineeProfile"
            sSQL = sSQL & " WHERE iNendo=" & g_int_CurrentNendo
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass & " AND iAbsentFlag = 0"
            sSQL = sSQL & " AND dtSecondExamDay = ( select top 1 dtSecondExamDay" & Trim(str(cboDay.ListIndex + 1))
            sSQL = sSQL & "                          from tbSTESecondExamProfile as se "
            sSQL = sSQL & "                          where exists ( select 1 from tbSTESystemProfile as sp where sp.iSystemProfileId = se.iSystemProfileId ) ) "
            sSQL = sSQL & " AND iJukenNumber > " & Trim(str(lStartNo)) & " ORDER BY iJukenNumber ) as t1 "
            Set oRs = g_obj_Conn.Execute(sSQL)
            If Not oRs.EOF Then
                If IsNull(oRs.Fields(0)) Then
                    For lRow2 = lRow To .Rows - 1
                        msfRoomAlloc.TextMatrix(lRow, prvlStartNoCol) = "-"
                        msfRoomAlloc.TextMatrix(lRow, prvlEndNoCol) = "-"
                    Next
                    Exit For
                End If
                lStartNo = oRs.Fields(0)
                msfRoomAlloc.TextMatrix(lRow, prvlStartNoCol) = Trim(str(lStartNo))
                lStartNo = oRs.Fields(1)
                msfRoomAlloc.TextMatrix(lRow, prvlEndNoCol) = Trim(str(lStartNo))
                oRs.Close
            End If
            Set oRs = Nothing
        Next
    End With

End Sub

Private Sub msfRoomAlloc_Click()
    ' on clicking any row, that form fields will be populated with data & _
        of the currently slected row
    
    Dim l_int_RowCounter As Long     ' row counter
    Dim l_int_ColCounter As Long     ' column counter
    Dim l_int_CurRow As Long         ' current row
    Dim l_int_CurCol As Long         ' current row
    On Error GoTo ErrorHandler
                
    With msfRoomAlloc
        If .Rows <= 1 Then Exit Sub     ' exit if there are no rows in the grid
        .FocusRect = flexFocusLight
        l_int_CurCol = .Col
        cmdDelete.Enabled = True
        l_int_CurRow = .Row
        
'        .Col = 5                        ' set the day
'        cboDay.Text = .Text

'        .Col = 1                        ' set the room
'        cboRooms.Text = .Text
        
        For l_int_RowCounter = 1 To .Rows - 1
            .Row = l_int_RowCounter
            For l_int_ColCounter = 0 To .cols - 1
                .Col = l_int_ColCounter
                If .Row <> l_int_CurRow Then
                    .CellBackColor = &HFFFFFF
                Else
                    .CellBackColor = &HC0C0FF   ' color the selected row
                End If
            Next
        Next
        .Row = l_int_CurRow
        .Col = l_int_CurCol
        .FocusRect = flexFocusNone
    End With
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Function f_void_GetRoomId(l_str_RoomName As String) As Long
    ' retreive the room Id, with room name as parameter
    Dim l_str_sqlRoom As String
    Dim l_obj_RsRoom As New ADODB.Recordset
    
    l_str_sqlRoom = "SELECT iRoomProfileId FROM tbSTERoomProfile" & _
        " WHERE vRoomName='" & l_str_RoomName & "'"
    l_obj_RsRoom.Open l_str_sqlRoom, g_obj_Conn
    
    If Not l_obj_RsRoom.EOF Then
        f_void_GetRoomId = l_obj_RsRoom.Fields("iRoomProfileId").Value
    Else
        f_void_GetRoomId = -1
    End If
    l_obj_RsRoom.Close
    Set l_obj_RsRoom = Nothing
End Function

Private Function f_void_GetSubjectId(l_str_SubjectName As String) As Long
    ' retreive the subject Id, with subject name as parameter
    Dim l_str_sqlSubject As String
    Dim l_obj_rsSubject As New ADODB.Recordset
    
    l_str_sqlSubject = "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & l_str_SubjectName & "'"
    l_obj_rsSubject.Open l_str_sqlSubject, g_obj_Conn
    
    If Not l_obj_rsSubject.EOF Then
        f_void_GetSubjectId = l_obj_rsSubject.Fields("iSubjectProfileId").Value
    Else
        f_void_GetSubjectId = -1
    End If
    l_obj_rsSubject.Close
    Set l_obj_rsSubject = Nothing
End Function

Private Sub msfRoomAlloc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then msfRoomAlloc_Click
End Sub

Private Sub msfRoomAlloc_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> prvlSerialCol And Col <> prvlCapacityCol Then
        Cancel = True
        Exit Sub
    End If
    prvsCurSerial = Trim(msfRoomAlloc.TextMatrix(Row, Col))
    If prvsCurSerial = "" Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub lsGetTotalExamineeCnt()

Dim oRs As ADODB.Recordset
Dim sSQL As String

On Error GoTo ErrProc

    sSQL = "SELECT count(*) from tbSTEExamineeProfile "
    sSQL = sSQL & " WHERE iNendo=" & g_int_CurrentNendo
    sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass & " AND iAbsentFlag = 0"
    sSQL = sSQL & " AND dtSecondExamDay = ( SELECT top 1 dtSecondExamDay" & Trim(str(cboDay.ListIndex + 1))
    sSQL = sSQL & "                           FROM tbSTESecondExamProfile as se "
    sSQL = sSQL & "                          WHERE exists ( select 1 from tbSTESystemProfile as sp where sp.iSystemProfileId = se.iSystemProfileId ) )"

    Set oRs = g_obj_Conn.Execute(sSQL)

    txtTotalExaminees.Text = Trim(str(oRs.Fields(0)))

    oRs.Close
    Set oRs = Nothing

Exit Sub

ErrProc:

End Sub

Private Sub lsGetUnallocatedExamineeCnt()

Dim oRs As ADODB.Recordset
Dim sSQL As String

On Error GoTo ErrProc

    sSQL = "SELECT count(*) from tbSTEExamineeProfile "
    sSQL = sSQL & " WHERE iNendo=" & g_int_CurrentNendo
    sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass & "  AND iAbsentFlag = 0"
    sSQL = sSQL & " AND dtSecondExamDay = ( SELECT top 1 dtSecondExamDay" & Trim(str(cboDay.ListIndex + 1))
    sSQL = sSQL & "                           FROM tbSTESecondExamProfile as se "
    sSQL = sSQL & "                          WHERE exists ( select 1 from tbSTESystemProfile as sp where sp.iSystemProfileId = se.iSystemProfileId ) )"
    sSQL = sSQL & " AND iJukenNumber > " & Trim(str(msfRoomAlloc.TextMatrix(msfRoomAlloc.Rows - 1, 7)))

    Set oRs = g_obj_Conn.Execute(sSQL)

    txtUnallocatedExaminees.Text = Trim(str(oRs.Fields(0)))

    oRs.Close
    Set oRs = Nothing

Exit Sub

ErrProc:

End Sub

Private Sub txtCapacity_KeyPress(KeyAscii As Integer)
    Call NumericOnly(Me, KeyAscii)
End Sub

Private Sub txtRandomNo_KeyPress(KeyAscii As Integer)
    Call NumericOnly(Me, KeyAscii)
End Sub
