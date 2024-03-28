VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmPassPersonData 
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   11955
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6135
      Begin VB.CommandButton cmdOK 
         Caption         =   "1060"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   6360
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "1066"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   6360
         Width           =   1455
      End
      Begin VSFlex7LCtl.VSFlexGrid vsfPassData 
         Height          =   4335
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   4215
         _cx             =   7435
         _cy             =   7646
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
      ForeColor       =   &H000000FF&
      Height          =   5295
      Left            =   6960
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   4215
   End
End
Attribute VB_Name = "frmPassPersonData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmPassPersonData
'Author         :   Vishal Kamath
'Created On     :
'Description    :   This form makes a provision for entering the Juken numbers of students who have
'                   passed the First Examination System.
'Reference      :   Functional Specs Of Pass Person Data Ver 1.0
'**************************************************************************************************

Dim m_int_CurYear As Integer
Dim m_str_arrExistingJukenNo() As String    'Existing JukenNo
Dim m_str_arrOriginalJukenNo() As String    'Original JukenNo
Dim m_int_arrErr() As Integer
Dim m_bln_NewRow As Boolean                 'To check for 10 rows


Private Sub cmdClear_Click()
    Dim l_int_Rowcount As Integer
    Dim l_int_ColCount As Integer
    Dim l_int_OriginalRows As Integer       'To Check for
    
    With vsfPassData
        .Cols = 1
        .Rows = 2
        .Col = 0
        .Row = 1
        .Text = ""
        .Refresh
    End With
    lblErrorDetails.Caption = ""
    Frame1.Move 3480, 960
End Sub



Private Sub cmdOK_Click()
    Dim l_str_ExID As String                'The examinee profile ID
    Dim l_str_sql As String  '
    Dim l_int_Rowcount As Integer
    Dim l_int_OriginalRows As Integer       'To Check for
    Dim l_int_ColCount As Integer
    Dim l_int_ErrCounter As Integer
    On Error GoTo ErrorHandler
    Call f_void_Nendo                       'To fetch the current year
    
    l_str_ExID = ""
    With vsfPassData
        l_int_OriginalRows = 0
        For l_int_ColCount = 0 To .Cols - 1
            For l_int_Rowcount = 1 To .Rows - 1
                .Row = l_int_Rowcount
                .Col = l_int_ColCount
                If Trim(.Text) <> "" Then
                    l_int_OriginalRows = l_int_OriginalRows + 1
                    ReDim Preserve m_str_arrOriginalJukenNo(l_int_OriginalRows)
                    m_str_arrOriginalJukenNo(l_int_OriginalRows) = Trim(.Text)
                End If
            Next
        Next
    End With
    
    If l_int_OriginalRows > 0 Then
        l_str_ExID = Join(m_str_arrOriginalJukenNo, ",")
        l_str_ExID = Right(l_str_ExID, Len(l_str_ExID) - 1)
        lblErrorDetails.Visible = False
    Else
        lblErrorDetails.Visible = True
        lblErrorDetails.Caption = LoadResString(1928)
        With Frame1
            .Left = 720
            .Top = 960
         End With
         With lblErrorDetails
            .Top = 960
            .Left = 7200
            .Height = l_int_ErrCounter * 250 + 200
         End With
        Exit Sub
    End If
            
            
            
    l_str_sql = "SELECT iJukenNumber FROM tbSTEExamineeProfile  WHERE"
    l_str_sql = l_str_sql & " iNendo = " & m_int_CurYear
    l_str_sql = l_str_sql & " AND iJukenNumber in (" & l_str_ExID & ")"
    l_str_sql = l_str_sql & " AND iAbsentFlag =0"
    
    If fMainForm.m_int_Identifier = 2 Then
        l_str_sql = l_str_sql & " AND iExamineeStatus =1"
    End If
        
    If fMainForm.m_int_Identifier = 3 Then
        l_str_sql = l_str_sql & " AND iExamineeStatus =1"
    End If
            
    
    'To fetch the existing JukenNumbers
    Call f_void_Select(l_str_sql)
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
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




Private Sub Form_Activate()
    fMainForm.mnuTools.Enabled = False
    Dim index
    For index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(index).Enabled = False
    Next
End Sub

Private Sub Form_Load()
LoadResStrings Me
Call f_void_Setcaptions
    With vsfPassData
    
        .Visible = False
        .BackColor = &HFFFFFF    '"16641260"
        .BackColorBkg = &HFFFFFF       '&H80000018
        .BackColorFixed = &H8000000F         '&HF4DBC4    '"13500156"
        .BackColorSel = &H800000
        .FixedCols = 0
        .TextStyleFixed = flexTextFlat
        '.FontFixed.Bold = False
        .Font.Bold = False
        .Font.Name = "Verdana"
        .Font.Size = 8
        '.FontFixed.Underline = True
        .ForeColorFixed = &H80000008                '"8388608"
        .ForeColor = &H800000       '"4194304"
        .CellTextStyle = "0"
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .GridColor = &H808080
        .Visible = True
    
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 2
        .Cols = 1
        .Row = 0
        .Col = 0
        .Text = LoadResString(1703)
        .Row = 1
    End With
End Sub


Private Sub f_void_Select(ByVal l_str_sql As String)
Dim l_int_CurYear As Integer
Dim l_obj_Rst As ADODB.Recordset
Dim l_int_Row As Integer        'Counter to hold the
Dim l_int_counter As Integer    'Counter to hold the original rows
Dim l_str_Join  As String
Dim l_int_ExistingExaminees As Integer
Dim l_int_ErrCounter As Integer

On Error GoTo ErrorHandler
Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)

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
         lblErrorDetails.Caption = LoadResString(1761)
         With Frame1
            .Left = 720
            .Top = 960
         End With
         With lblErrorDetails
            .Top = 960
            .Left = 7200
            .Height = l_int_ErrCounter * 250 + 200
         End With
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
         lblErrorDetails.Visible = True
         With Frame1
            .Left = 720
            .Top = 960
        End With
        With lblErrorDetails
            .Top = 960
            .Left = 7200
            .Height = l_int_ErrCounter * 250 + 200
        End With
     End If
 
 Else
    lblErrorDetails.Caption = LoadResString(1929) & vbCrLf & vbCrLf
    ' all the entered values are wrong juken no's
    For l_int_ErrCounter = 1 To UBound(m_str_arrOriginalJukenNo)
         lblErrorDetails.Caption = lblErrorDetails.Caption & m_str_arrOriginalJukenNo(l_int_ErrCounter) & vbCrLf
    Next
    lblErrorDetails.Visible = True
    With Frame1
        .Left = 720
        .Top = 960
    End With
    With lblErrorDetails
        .Top = 960
        .Left = 7200
        .Height = l_int_ErrCounter * 250 + 200
    End With
    
 End If

Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
'To fetch the Current Year
Private Sub f_void_Nendo()
Dim l_int_CurYear As Integer       'The current year
Dim l_obj_Rst As ADODB.Recordset
Dim l_str_sql As String            'The SQL string
    
    
On Error GoTo ErrorHandler
l_str_sql = "SELECT iNendo FROM tbSTESystemProfile WHERE iActiveFlag=1"
Set l_obj_Rst = g_obj_Conn.Execute(l_str_sql)

If Not l_obj_Rst.EOF Then
    m_int_CurYear = l_obj_Rst("iNendo")
End If
' release the object variable
l_obj_Rst.Close
Set l_obj_Rst = Nothing
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
Private Sub vsfPassData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With vsfPassData
    If Len(Trim(.Text)) < 4 Then
        .Text = g_str_LPad(Trim(.Text), Len(Trim(.Text)))
    Else
        .Text = Trim(.Text)
    End If
End With
End Sub


Private Sub vsfPassData_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsfPassData
        If KeyAscii = vbKeyReturn Then
            If Trim(.Text) = "" Then
                Exit Sub
            End If
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
        ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
        
    End With
End Sub

'Returns a False value if Error in Updating the tables tbSTEExamineeprofile
'AND table tbSTEScoreProfile
Private Function f_bln_Update(ByVal str As String) As Boolean
Dim l_str_sql As String
On Error GoTo ErrorHandler

If Trim(str) <> "" Then
    ''Update the examinee status to pass 1st Exam for all students having who are present.
    
    
    If g_int_ExamType = 1 Then
        l_str_sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = 1 WHERE"
        l_str_sql = l_str_sql & " iNendo = " & m_int_CurYear
        l_str_sql = l_str_sql & " AND iJukenNumber IN (" & str & ")"
        l_str_sql = l_str_sql & " AND iAbsentFlag =0"
        l_str_sql = l_str_sql & " AND iExamineeStatus =0"
    End If
    
    
    Select Case fMainForm.m_int_Identifier
    Case 2 ''2nd Exam
        l_str_sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = 2 WHERE"
        l_str_sql = l_str_sql & " iNendo = " & m_int_CurYear
        l_str_sql = l_str_sql & " AND iJukenNumber IN (" & str & ")"
        l_str_sql = l_str_sql & " AND iAbsentFlag =0"
        l_str_sql = l_str_sql & " AND iExamineeStatus =1"
    Case 3 ''Waiting
        l_str_sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = 3 WHERE"
        l_str_sql = l_str_sql & " iNendo = " & m_int_CurYear
        l_str_sql = l_str_sql & " AND iJukenNumber IN (" & str & ")"
        l_str_sql = l_str_sql & " AND iAbsentFlag =0"
        l_str_sql = l_str_sql & " AND iExamineeStatus =1"
    End Select
    
    g_obj_Conn.Execute l_str_sql
    f_bln_Update = True
End If
Exit Function
ErrorHandler:
    MsgBox Err.Description
    f_bln_Update = False
End Function


Private Sub f_void_Setcaptions()
If g_int_ExamType = 1 Then
    frmPassPersonData.Caption = LoadResString(1926)
    lblPastPersonData.Caption = LoadResString(1926)
End If

Select Case fMainForm.m_int_Identifier
Case 2
    frmPassPersonData.Caption = LoadResString(1927)
    lblPastPersonData.Caption = LoadResString(1927)
End Select
End Sub

