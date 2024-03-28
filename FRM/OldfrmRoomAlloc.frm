VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmRoomAlloc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F4DBC4&
   Caption         =   "Room Allocation"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   6960
      Width           =   1455
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfRoomAlloc 
      Height          =   2175
      Left            =   2520
      TabIndex        =   0
      Top             =   4080
      Width           =   10695
      _cx             =   18865
      _cy             =   3836
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      AllowUserResizing=   0
      SelectionMode   =   0
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
      Begin VB.ComboBox cboRooms 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Label lblRoomAlloc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Room Allocation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   6135
   End
End
Attribute VB_Name = "frmRoomAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_str_SQl As String
Dim m_obj_Rst As New ADODB.Recordset

Dim m_int_TotalExaminees As Integer
Dim m_int_CurrentRow As Integer
Dim m_bln_Edit As Boolean
Dim m_int_OldCapacity As Integer
Dim m_int_OldRandom As Integer
Dim m_int_NewCapacity As Integer
Dim m_int_NewRandom As Integer
Dim m_int_TotalAllotted As Integer
Dim m_int_ToBeAllotted As Integer
Dim m_cbo_NewCombo As Object
Dim m_int_SrNo As Integer
Dim m_int_StartExamineeNo As Integer

Private Sub cboRooms_Click()
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    
    l_str_Sql = "Select iRandom, iMaxCapacity from tbSTERoomProfile where vRoomName='" & cboRooms.Text & "'"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    ' check for error
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
        Exit Sub
    End If
    
    If Not l_obj_Rst.EOF Then
        With vsfRoomAlloc
            .Col = 2
            If Trim(l_obj_Rst("iRandom")) <> "" Then
                .Text = l_obj_Rst("iRandom")
                m_int_OldRandom = l_obj_Rst("iRandom")
            Else
                .Text = 0
                m_int_OldRandom = 0
            End If
            
            .Col = 3
            If Trim(l_obj_Rst("iMaxCapacity")) <> "" Then
                .Text = l_obj_Rst("iMaxCapacity")
                m_int_OldCapacity = l_obj_Rst("iMaxCapacity")
            Else
                .Text = 0
                m_int_OldCapacity = 0
            End If
        End With
    End If
        
End Sub

Private Sub Form_Load()
         
    m_str_SQl = "Select iExamineeProfileId from tbSTEExamineeProfile"
    m_str_SQl = m_str_SQl & " where iNendo=(Select iNendo from tbSTESystemProfile where iActiveFlag=1)"
    
    m_obj_Rst.Open m_str_SQl, g_obj_Conn, adOpenStatic, adLockReadOnly
    
    ' check for error
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
        Exit Sub
    End If
    
    m_int_TotalExaminees = m_obj_Rst.RecordCount
        
    Call f_void_AddRooms
    
    Call f_void_InitGrid
    
    m_int_StartExamineeNo = 1

    m_obj_Rst.Close
    Set m_obj_Rst = Nothing
End Sub

Private Function f_void_PopulateGrid()
    Dim l_int_SrNo As Integer
       
    m_obj_Rst.MoveFirst
    
    With vsfRoomAlloc
        
        
'        .Row = 1
'        .Col = 0
'        .Text = l_int_SrNo
        
'        .Col = .Col + 1
        
    End With

End Function


Private Sub f_void_AddRooms()
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    
    l_str_Sql = "Select vRoomName from tbSTERoomProfile"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
        Exit Sub
    End If
    
    If Not l_obj_Rst.EOF Then
        Do While Not l_obj_Rst.EOF
            cboRooms.AddItem l_obj_Rst("vRoomName")
            l_obj_Rst.MoveNext
        Loop
        
        cboRooms.ListIndex = 0
    End If
    
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
End Sub

Private Sub vsfRoomAlloc_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim l_int_Capacity As Integer
    If OldRow <> NewRow And NewRow <> 0 Then
'        MsgBox "rowchanged" & OldRow & "--" & NewRow
        Call f_void_FindTotalAllocated
        With vsfRoomAlloc
            .Col = 3
            l_int_Capacity = CInt(.Text)
            m_int_ToBeAllotted = m_int_TotalExaminees - m_int_TotalAllotted
            
            If l_int_Capacity < m_int_ToBeAllotted And l_int_Capacity <> 0 Then
                .Col = .Col + 1
                .Text = m_int_StartExamineeNo
                .Col = .Col + 1
                .Text = m_int_StartExamineeNo + l_int_Capacity
            Else
                .Col = .Col + 1
                .Text = m_int_StartExamineeNo
                .Col = .Col + 1
                .Text = m_int_StartExamineeNo + m_int_TotalAllotted
            End If
        End With
    End If
End Sub

Private Sub vsfRoomAlloc_KeyPress(KeyAscii As Integer)
    If vsfRoomAlloc.Col <= 1 Or vsfRoomAlloc.Col >= 4 Then
        KeyAscii = 0
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

'Private Sub vsfRoomAlloc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    Dim l_int_counter As Integer
'    Dim l_bln_Update As Boolean
'
'    With vsfRoomAlloc
'        If OldRow = 1 And NewRow = 0 Then
'
'        ElseIf .Row <> m_int_CurrentRow And OldRow <> 0 Then
'            MsgBox "row changed"
'            .Rows = .Rows + 1
'        End If
''        If OldRow <> NewRow Then
''
''            .Row = NewRow
''            .Col = 2
''            If .Text <> "" Then
''                m_int_NewRandom = .Text
''            Else
''                m_int_NewRandom = 0
''            End If
''
''            .Col = 3
''            If .Text <> "" Then
''                m_int_NewCapacity = .Text
''            Else
''                m_int_NewCapacity = 0
''            End If
''
''            l_bln_Update = f_bln_UpdateData()
''        End If
''
''
'        If .Row <> 0 And .Row <> .Rows - 1 And .Col = 1 Then
'            cboRooms.Move .CellLeft, .CellTop, .CellWidth
'            cboRooms.Visible = True
'            .Text = cboRooms.Text
'        End If
'    End With
'
'End Sub
Private Sub vsfRoomAlloc_LostFocus()
    If vsfRoomAlloc.Rows = 2 And vsfRoomAlloc.Col <> 1 Then
        m_bln_Edit = True
        m_int_CurrentRow = 0
        Call vsfRoomAlloc_AfterRowColChange(vsfRoomAlloc.Row, vsfRoomAlloc.Col, 0, 0)
    End If
End Sub

Private Sub vsfRoomAlloc_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        MsgBox "edit true"
        m_bln_Edit = True
        m_int_CurrentRow = vsfRoomAlloc.Row
    End If
End Sub

Private Function f_bln_UpdateData() As Boolean
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    
    If (m_int_OldCapacity <> m_int_NewCapacity And m_int_OldCapacity <> 0) Or (m_int_OldRandom <> m_int_NewRandom And m_int_OldRandom <> 0) Then
        l_str_Sql = "Update tbSTERoomProfile set iRandom=" & m_int_NewRandom & ","
        l_str_Sql = l_str_Sql & " iMaxCapacity=" & m_int_NewCapacity
        l_str_Sql = l_str_Sql & " where vRoomName='" & cboRooms.Text & "'"
    
        Set m_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
        
        If Err.Number <> 0 Then
            MsgBox Err.Number & vbCrLf & Err.Description
            Exit Function
        End If
        
        m_obj_Rst.Close
        Set m_obj_Rst = Nothing
        
        Call f_void_CheckForNewRow
    End If
End Function

Private Sub f_void_CheckForNewRow()
    
    
    
        
        If m_int_TotalAllotted < m_int_TotalExaminees Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 1
        End If
    End With
        
    End If
End Sub

Private Sub f_void_FindTotalAllocated()
    Dim l_int_counter As Integer
    
    With vsfRoomAlloc
        For l_int_counter = 1 To .Rows - 1
            .Row = l_int_counter
            .Col = 3
            m_int_TotalAllotted = m_int_TotalAllotted + CInt(.Text)
            l_int_counter = l_int_counter + 1
        Next
    End With
End Sub


Private Sub f_void_InitGrid()
    
    With vsfRoomAlloc
        .Row = 0
        .Rows = 3
        .Cols = 6
        .RowHeight(1) = 320
        
        .FixedRows = 1
        .FixedCols = 0
        
        .Row = 0
        .Col = 0
        .ColWidth(0) = 972
        .Text = "Sr No"
        
        .Col = .Col + 1
        .ColWidth(1) = 2500
        .Text = "Room Name"
        
        .Col = .Col + 1
        .ColWidth(2) = 1800
        .Text = "Random No"
        
        .Col = .Col + 1
        .ColWidth(3) = 1800
        .Text = "Capacity"
        
        .Col = .Col + 1
        .ColWidth(4) = 1800
        .Text = "Start Juken No"
        
        .Col = .Col + 1
        .ColWidth(5) = 1800
        .Text = "End Juken No"
        
        .Row = 1
        .Col = 0
        m_int_SrNo = 1
        .Text = m_int_SrNo
        
        .Col = .Col + 1
    End With
End Sub
