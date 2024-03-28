VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmPassPersonData1 
   BackColor       =   &H00F4DBC4&
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   11175
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfAbsentee 
      Height          =   4215
      Left            =   5760
      TabIndex        =   1
      Top             =   1800
      Width           =   4335
      _cx             =   7646
      _cy             =   7435
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
      BackColor       =   16047044
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16047044
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
      Rows            =   10
      Cols            =   3
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
   Begin VB.Label lblError 
      BackStyle       =   0  'Transparent
      Caption         =   "Error"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblErrorDetails 
      BackStyle       =   0  'Transparent
      Caption         =   "details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   6015
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblPastPersonData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pass Person Data (status 1-->2)"
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
      Height          =   1335
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmPassPersonData1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The passed person data for the Second phase

Dim m_int_CurYear As Integer
Dim m_str_arrExistingJukenNo() As String    'Existing JukenNo
Dim m_str_arrOriginalJukenNo() As String    'Original JukenNo
Dim m_int_arrErr() As Integer

Private Sub cmdOK_Click()
    Dim l_str_ExID As String                'The examinee profile ID
    Dim l_str_Sql As String  '
    Dim l_int_Rowcount As Integer
    Dim l_int_OriginalRows As Integer       'To Check for
    
    Call f_void_Nendo                       'To fetch the current year
    
    l_str_ExID = ""
    With vsfAbsentee
        l_int_OriginalRows = 0
        For l_int_Rowcount = 1 To .Rows - 1
            .Row = l_int_Rowcount
            .Col = 0
            If Trim(.Text) <> "" Then
                l_int_OriginalRows = l_int_OriginalRows + 1
                ReDim Preserve m_str_arrOriginalJukenNo(l_int_OriginalRows)
                m_str_arrOriginalJukenNo(l_int_OriginalRows) = Trim(.Text)
            End If
        Next
    End With
    
    If l_int_OriginalRows > 0 Then
        l_str_ExID = Join(m_str_arrOriginalJukenNo, ",")
        l_str_ExID = Right(l_str_ExID, Len(l_str_ExID) - 1)
        lblError.Visible = False
        lblErrorDetails.Visible = False
    Else
        lblError.Visible = True
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = "You haven't entered any Juken Number"
        Exit Sub
    End If
            
    l_str_Sql = "SELECT iJukenNumber FROM tbSTEExamineeProfile  WHERE"
    l_str_Sql = l_str_Sql & " iNendo = " & m_int_CurYear
    l_str_Sql = l_str_Sql & " AND iJukenNumber in (" & l_str_ExID & ")"
    l_str_Sql = l_str_Sql & " AND iAbsentFlag =0"
    l_str_Sql = l_str_Sql & " AND iExamineeStatus =1"
    
    
    'To fetch the existing JukenNumbers
    Call f_void_Select(l_str_Sql)
    
End Sub

Public Function g_str_LPad(ByVal str As String, ByVal iLen As Integer) As String

'-------------------------------------------------------------
'Left pads a string with 0 up to iLen.
'-------------------------------------------------------------
Select Case iLen
Case 1
    g_str_LPad = "000" & str
Case 2
    g_str_LPad = "00" & str
Case 3
    g_str_LPad = "0" & str
End Select

End Function




Private Sub Form_Load()
    
    With vsfAbsentee
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 2
        .Cols = 1
        .Row = 0
        .Col = 0
        .Text = "Juken No"
        .Row = 1
    End With
End Sub


Private Sub f_void_Select(ByVal l_str_Sql As String)
Dim l_int_CurYear As Integer
Dim l_obj_Rst As ADODB.Recordset
Dim l_int_Row As Integer        'Counter to hold the
Dim l_int_counter As Integer    'Counter to hold the original rows
Dim l_str_Join  As String
Dim l_int_ExistingExaminees As Integer
Dim l_int_ErrCounter As Integer

'On Error Resume Next

Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)

'
If Err.Number <> 0 Then
    MsgBox Err.Number & vbCrLf & Err.Description
    Exit Sub
End If
 
 
 
If Not l_obj_Rst.EOF Then
    l_int_ExistingExaminees = 0
    While Not l_obj_Rst.EOF
        l_int_ExistingExaminees = l_int_ExistingExaminees + 1
        ReDim Preserve m_str_arrExistingJukenNo(l_int_ExistingExaminees)
        m_str_arrExistingJukenNo(l_int_ExistingExaminees) = l_obj_Rst("iJukenNumber").Value
        l_obj_Rst.MoveNext
    Wend
End If
 
' release the object variable
l_obj_Rst.Close
Set l_obj_Rst = Nothing
 
 'Form a string with comma as a delimiter and
 'remove the first comma.
 If l_int_ExistingExaminees > 0 Then
     l_str_Join = Join(m_str_arrExistingJukenNo, ",")
     l_str_Join = Right(l_str_Join, Len(l_str_Join) - 1)
     
     ' update the data for existing employees
     If Not f_bln_Update(l_str_Join) Then
         lblErrorDetails.Visible = True
         lblErrorDetails.Caption = "There was an Error while trying to do the Update operation"
         Exit Sub
     End If
     
     ' find the erreneous juken numbers
     l_int_ErrCounter = 0
 
     For l_int_counter = 1 To UBound(m_str_arrOriginalJukenNo)
         For l_int_Row = 1 To UBound(m_str_arrExistingJukenNo)
             If m_str_arrOriginalJukenNo(l_int_counter) = m_str_arrExistingJukenNo(l_int_Row) Then
                 Exit For
             Else
                 If l_int_Row = UBound(m_str_arrExistingJukenNo) Then
                     l_int_ErrCounter = l_int_ErrCounter + 1
                     ReDim Preserve m_int_arrErr(l_int_ErrCounter)
                     m_int_arrErr(l_int_ErrCounter) = m_str_arrOriginalJukenNo(l_int_counter)
                 End If
             End If
         Next
     Next
     
     ' display error
     If l_int_ErrCounter > 1 Then
         lblErrorDetails.Caption = "The following Juken No's were not present for the current year " & vbCrLf
         
         For l_int_ErrCounter = 1 To UBound(m_int_arrErr)
             lblErrorDetails.Caption = lblErrorDetails.Caption & m_int_arrErr(l_int_ErrCounter) & vbCrLf
         Next
         
         lblError.Visible = True
         lblErrorDetails.Visible = True
     End If
 
 Else
    lblErrorDetails.Caption = "The following Juken No's were not present for the current year" & vbCrLf
    ' all the entered values are wrong juken no's
    For l_int_ErrCounter = 1 To UBound(m_str_arrOriginalJukenNo)
         lblErrorDetails.Caption = lblErrorDetails.Caption & m_str_arrOriginalJukenNo(l_int_ErrCounter) & vbCrLf
    Next
    lblError.Visible = True
    lblErrorDetails.Visible = True
    
    
    
    
 End If
    
    
    
End Sub
'To fetch the Current Year
Private Sub f_void_Nendo()
Dim l_int_CurYear As Integer       'The current year
Dim l_obj_Rst As ADODB.Recordset
Dim l_str_Sql As String            'The SQL string
    
l_str_Sql = "Select iNendo from tbSTESystemProfile where iActiveFlag=1"
Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)

' check for error
If Err.Number <> 0 Then
    MsgBox Err.Number & vbCrLf & Err.Description
    Exit Sub
End If

If Not l_obj_Rst.EOF Then
    m_int_CurYear = l_obj_Rst("iNendo")
End If
' release the object variable
l_obj_Rst.Close
Set l_obj_Rst = Nothing
End Sub


Private Sub vsfAbsentee_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With vsfAbsentee
    If Len(Trim(.Text)) < 4 Then
        .Text = g_str_LPad(Trim(.Text), Len(Trim(.Text)))
    Else
        .Text = Trim(.Text)
    End If
End With
End Sub

Private Sub vsfAbsentee_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If Not m_bln_NewRow Then
            .AddItem "", .Rows
            .Row = .Rows - 1
            .Col = 0
            .Refresh
            
            If .Row > 10 Then
                m_bln_NewRow = True
                .Cols = .Cols + 1
                .Row = 1
                .Col = .Cols - 1
                .Refresh
            End If
        Else
            .Row = .Row + 1
            If .Row > 10 Then
                .Cols = .Cols + 1
                .Row = 1
                .Col = .Cols - 1
            End If
        End If
End If
End Sub

'Returns a False value if Error in Updating the tables tbSTEExamineeprofile
'AND table tbSTEScoreProfile
Private Function f_bln_Update(ByVal str As String) As Boolean
Dim l_str_Sql As String
If Trim(str) <> "" Then
    ''Update the examinee status to pass 1st Exam for all students having who are present.
    l_str_Sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = 2 WHERE"
    l_str_Sql = l_str_Sql & " iNendo = " & m_int_CurYear
    l_str_Sql = l_str_Sql & " AND iJukenNumber IN (" & str & ")"
    l_str_Sql = l_str_Sql & " AND iAbsentFlag =0"
    l_str_Sql = l_str_Sql & " AND iExamineeStatus =1"
    
    g_obj_Conn.Execute l_str_Sql
    If Err.Number <> 0 Then
        f_bln_Update = False
        Exit Function
    End If
    f_bln_Update = True
End If
End Function



