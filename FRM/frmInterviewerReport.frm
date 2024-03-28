VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmInterviewerReport 
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmInterviewerReport.frx":0000
   ScaleHeight     =   10335
   ScaleWidth      =   13080
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdTeacher 
      Caption         =   "表示"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9990
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtTeacher 
      Height          =   390
      Left            =   8400
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
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
      Left            =   4650
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ListBox lstRoom 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      Left            =   9000
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   8760
      Width           =   1095
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfInterviewerRoom 
      Height          =   6330
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   8655
      _cx             =   15266
      _cy             =   11165
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
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
      Left            =   1080
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "教員"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7260
      TabIndex        =   10
      Top             =   1155
      Width           =   1035
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "科目名"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3570
      TabIndex        =   7
      Top             =   1150
      Width           =   975
   End
   Begin VB.Label lblErrorDetails 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   10455
   End
   Begin VB.Label lblDay 
      BackStyle       =   0  '透明
      Caption         =   "面接日"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      TabIndex        =   1
      Top             =   1150
      Width           =   855
   End
End
Attribute VB_Name = "frmInterviewerReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmInterviewerRoom
'Author         :   Vishal Kamath
'Created On     :   30/10/01
'Description    :   This form makes a provision for mapping interviewers to rooms.
'Reference      :   FunctionalSpecs OFInterviewerRoommapping Ver1.1.doc
'**************************************************************************************************
'Modification History:
'12/5/02 Included Day ComboBox
'14/5/02 Insertion in TbsteSubjectQuestionProfile
'**************************************************************************************************
' Ammendments - NyushiChangesSummary.doc ver 1.0
'20/5/02 - Changed to add day combo and subsequent changes

Private f_obj_RsInterviewerID As New ADODB.Recordset            'Recordset object variable to open Interviewer table
Private f_str_InterviewerName As String                         'String variable to hold the SQL string
Private f_int_InterviewRoomProfileID As Long                 'to store the interviewer profile id
Private f_bln_SelectAll As Boolean                              'Shows the status of the Select All button
Private f_bln_Select As Boolean                                 'Shows the status of the Select  button
Private f_bln_DeSelect As Boolean                               'Shows the status of the DeSelectAll button
Private f_bln_DeSelectAll As Boolean                            'Shows the status of the DeSelect  button
Private f_int_RoomProfileID As Long                          'Stores the RoomProfileID
Public m_int_IntRpt As Long                                  'form variable variable which indicated whether the form has to be instantiated for the "interview" or "report"

Private prvlRoomStartCol As Long

Private Sub cboDay_Click()
    Call f_void_LoadRoom
    Call f_void_PopulateGrid
End Sub

'refresh the all interviewers and selected interviewers list on change of the subject
Private Sub cboSubject_Click()
End Sub

Private Sub cmdTeacher_Click()
f_void_PopulateGrid
End Sub

Private Sub cmdUpdate_Click()

Dim lRow As Long
Dim lCol As Long
Dim sSQL As String
Dim oRs As ADODB.Recordset
Dim sInterviewerID As String
Dim sRoomID As String

On Error GoTo ErrorHandler

'トランザクション開始
    g_obj_Conn.BeginTrans
'del,xzg,2009/12/22,S--------------
''指定日の登録済みデータを削除する
'    sSQL = "Delete from tbSTEInterviewRoomProfile "
'    sSQL = sSQL & " WHERE iNendo = " & g_int_CurrentNendo
'    sSQL = sSQL & " AND iSubjectProfileId = " & cboSubject.ItemData(cboSubject.ListIndex)
'    If m_int_IntRpt = 0 Then
''面接の場合RoomProfile（面接グループを定義している）を使用
'        sSQL = sSQL & " and iRoomProfileId is not null "
'    Else
''小論文の場合、乱数のみを使用
'        sSQL = sSQL & " and iRandomNo is not null "
'    End If
'    sSQL = sSQL & " and iDayFlag = " & Trim(str(Me.cboDay.ListIndex))
'
'    g_obj_Conn.Execute sSQL
'del,xzg,2009/12/22,E--------------
    '試験日を保持
    With vsfInterviewerRoom
'行ループ：１行は担当者
        For lRow = 1 To .Rows - 1
            '担当者IDを保持
            sInterviewerID = .TextMatrix(lRow, 0)
 'add,xzg,2009/12/22,S--------------
            '指定日の登録済みデータを削除する
            sSQL = "Delete from tbSTEInterviewRoomProfile "
            sSQL = sSQL & " WHERE iNendo = " & g_int_CurrentNendo
            sSQL = sSQL & " AND iSubjectProfileId = " & cboSubject.ItemData(cboSubject.ListIndex)
            If m_int_IntRpt = 0 Then
        '面接の場合RoomProfile（面接グループを定義している）を使用
                sSQL = sSQL & " and iRoomProfileId is not null "
            Else
        '小論文の場合、乱数のみを使用
                sSQL = sSQL & " and iRandomNo is not null "
            End If
            sSQL = sSQL & " and iDayFlag = " & Trim(str(Me.cboDay.ListIndex))
            sSQL = sSQL & " and iInterviewerProfileId = " & sInterviewerID
            g_obj_Conn.Execute sSQL
  'add,xzg,2009/12/22,E--------------
'列ループ：複数の部屋および乱数
            For lCol = prvlRoomStartCol To .cols - 2 Step 2
                sRoomID = Trim(.TextMatrix(lRow, lCol))
                If sRoomID = "" Then GoTo EndFor
                'tbSTEInterviewerRoomProfileに登録
                sSQL = "Insert into tbSTEInterviewRoomProfile ( "
                sSQL = sSQL & "  iInterviewRoomProfileId "
                sSQL = sSQL & ", iNendo "
                sSQL = sSQL & ", iInterviewerProfileId "
                If m_int_IntRpt = 0 Then
'面接の場合RoomProfile（面接グループを定義している）を使用
                    sSQL = sSQL & ", iRoomProfileId "
                Else
'小論文の場合、乱数のみを使用
                    sSQL = sSQL & ", iRandomNo "
                End If
                sSQL = sSQL & ", iSubjectProfileID "
                sSQL = sSQL & ", iDayFlag "
                sSQL = sSQL & ") "
                sSQL = sSQL & " select "
                sSQL = sSQL & "  isnull( max( iInterviewRoomProfileId + 1 ) , 1 ) "
                sSQL = sSQL & ", " & g_int_CurrentNendo
                sSQL = sSQL & ", " & sInterviewerID
                sSQL = sSQL & ", " & sRoomID
                sSQL = sSQL & ", " & Trim(str(cboSubject.ItemData(cboSubject.ListIndex)))
                sSQL = sSQL & ", " & Trim(str(Me.cboDay.ListIndex))
                sSQL = sSQL & " from tbSTEInterviewRoomProfile "

                g_obj_Conn.Execute sSQL

EndFor:

            Next
'列ループ：複数の部屋および乱数End
        Next
'行ループ：１行は担当者End

    End With

'トランザクション終了
    g_obj_Conn.CommitTrans

    lblErrorDetails.Caption = LoadResString(2404)

Exit Sub

ErrorHandler:
'トランザクションロールバック
    g_obj_Conn.RollbackTrans
    MsgBox Err.Description, vbInformation, LoadResString(1729)


End Sub

Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    fMainForm.mnuTools.Enabled = False          'disable the tools menu as this is not a master maintenance screen
    Dim Index As Long
    For Index = 1 To fMainForm.Toolbar1.Buttons.Count   'disable all the toolbar button also
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next
    If m_int_IntRpt = 0 Then
        'its for the interview
        Me.Caption = LoadResString(1051)
    Else
        ' its for the report
        Me.Caption = LoadResString(1053)
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    LoadResStrings Me               'load captions from resource file
    Call g_void_SetFontProperties(Me)     ' set the font properties
    Call f_void_InitGrid
    Call l_void_PopulateDayCombo
    Call f_void_PopulateSubject
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Public Sub f_void_LoadRoom()        'populate the room names
    Dim l_obj_RsRoom As New ADODB.Recordset
    Dim l_str_sqlRoom As String
    
    On Error GoTo ErrorHandler

    lstRoom.Clear

    If m_int_IntRpt = 0 Then
'面接の場合RoomProfile（面接グループを定義している）を使用
        l_str_sqlRoom = "SELECT iRoomProfileid,vRoomName FROM tbSTERoomProfile" & _
            " WHERE iInterviewRoomFlag = 0 " & _
            "   AND iMaxCapacity > 0 " & _
            " ORDER BY iRoomProfileid "
    Else
'小論文の場合、乱数のみを使用
        l_str_sqlRoom = "SELECT distinct iShoronbunRandomNo , iShoronbunRandomNo FROM tbSTEExamineeProfile " & _
            " WHERE iNendo = " & g_int_CurrentNendo & _
            " AND dtSecondExamDay = ( SELECT top 1 dtSecondExamDay" & Trim(str(cboDay.ListIndex + 1)) & _
            "                           FROM tbSTESecondExamProfile as se " & _
            "                          WHERE exists ( select 1 from tbSTESystemProfile as sp where sp.iSystemProfileId = se.iSystemProfileId ) )" & _
            "   AND iShoronbunRandomNo is not null " & _
            " ORDER BY iShoronbunRandomNo "
    End If
    
    l_obj_RsRoom.Open l_str_sqlRoom, g_obj_Conn
    Do While Not l_obj_RsRoom.EOF
        lstRoom.AddItem l_obj_RsRoom.Fields(1).Value       'hidden combo to keep the id's of rooms
        lstRoom.ItemData(lstRoom.NewIndex) = l_obj_RsRoom.Fields(0).Value             'combo which displays the rooms names
        l_obj_RsRoom.MoveNext
    Loop

    l_obj_RsRoom.Close
    Set l_obj_RsRoom = Nothing
    Exit Sub
ErrorHandler:
        MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub
'
'Public Sub f_void_AllInterviewers()
'    'loads all the Interviewers
'    Dim l_str_Sql As String
'    Dim l_obj_RsInterviewers  As ADODB.Recordset
'    Dim l_int_Count As Integer
'    Dim l_str_SelectedIntwr() As String
'
'    On Error GoTo ErrorHandler
'    lstAllInterviewers.Clear
'    For l_int_Count = 0 To lstselectedInterviewers.ListCount - 1
'        ReDim Preserve l_str_SelectedIntwr(l_int_Count)
'        l_str_SelectedIntwr(l_int_Count) = Trim(lstselectedInterviewers.List(l_int_Count))
'    Next
'
'    Set l_obj_RsInterviewers = New ADODB.Recordset
'    l_str_Sql = "SELECT iInterviewerProfileId,vInterviewerName FROM tbSTEInterviewerProfile"
'    If l_int_Count > 0 Then ' select only those interviewers who are not assigned to any rooms for the selected subject
'        l_str_Sql = l_str_Sql & " WHERE vInterviewerName NOT IN('" & Join(l_str_SelectedIntwr, "','")
'        l_str_Sql = l_str_Sql & "')"
'    End If
'    l_obj_RsInterviewers.Open l_str_Sql, g_obj_Conn
'    Do While Not l_obj_RsInterviewers.EOF
'        lstAllInterviewers.AddItem l_obj_RsInterviewers.Fields("vInterviewerName").Value
'        l_obj_RsInterviewers.MoveNext
'    Loop
'    l_obj_RsInterviewers.Close
'    Set l_obj_RsInterviewers = Nothing
'Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Sub
'
'Private Sub cmdSelectAll_Click()
'    'On the click of this button all the Interviewers from the lstAllInterviewers will be transfered to lstSelectedInterviewers
'    Dim l_bln_existing As Boolean           ' variable to check whether an interviewer is already existing the in the list box or not
'    Dim l_int_Counter As Integer            ' counter variable
'    Dim l_int_AllInterviewers As Integer    ' count of interviewers in the list box
'    Dim l_str_Sql As String                 ' to store the SQL string
'    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
'    Dim l_int_IntwrId As Integer            ' to store the interviewer ID
'
'    On Error GoTo ErrorHandler
'
'    f_bln_SelectAll = True
'    If lstAllInterviewers.ListCount >= 1 Then
'        For l_int_AllInterviewers = 0 To lstAllInterviewers.ListCount - 1
'            l_bln_existing = False
'            For l_int_Counter = 0 To lstselectedInterviewers.ListCount - 1
'             If Trim(lstselectedInterviewers.List(l_int_Counter)) = Trim(lstAllInterviewers.List(l_int_AllInterviewers)) Then
'                l_bln_existing = True
'                Exit For
'             End If
'             Next
'            If Not l_bln_existing Then
'                lstselectedInterviewers.AddItem lstAllInterviewers.List(l_int_AllInterviewers)
'                lstAllInterviewers.ListIndex = l_int_AllInterviewers
'                f_str_InterviewerName = lstAllInterviewers.Text
'                l_str_Sql = "SELECT iInterviewerProfileId FROM tbSTEInterviewerProfile"
'                l_str_Sql = l_str_Sql & " WHERE vInterviewerName='" & f_str_InterviewerName & "'"
'                l_obj_Rst.Open l_str_Sql, g_obj_Conn
'                If Not l_obj_Rst.EOF Then
'                    l_int_IntwrId = l_obj_Rst("iInterviewerProfileId")
'                End If
'                l_obj_Rst.Close
'                Set l_obj_Rst = Nothing
'                f_void_UpdateDatabase (l_int_IntwrId)   ' update the database with the latest changes
'            End If
'        Next
'    End If
'
'    Call f_void_AllInterviewers     ' refresh the interviewers list
'    f_void_CheckButtonStatus        ' enable/disable the direction buttons after the latest change
'    f_bln_SelectAll = False
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Sub
'Private Sub cmdSelect_Click()
'    'on the click of this button only the Interviewer selected from the lstAllInterviewers will be transfered to
'    'lstSelectedInterviewers
'    Dim l_bln_existing As Boolean           ' variable to check whether an interviewer is already existing the in the list box or not
'    Dim l_int_Counter As Integer            ' counter variable
'    Dim l_int_Count As Integer              ' counter variable
'    Dim l_int_IntwrId As Integer            ' to store interviewer ID
'    Dim l_str_Sql As String                 ' to store the SQL string
'    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
'
'    On Error GoTo ErrorHandler
'
'    f_bln_Select = True
'    If lstAllInterviewers.SelCount > 0 Then
'        For l_int_Count = lstAllInterviewers.ListCount - 1 To 0 Step -1
'            If lstAllInterviewers.Selected(l_int_Count) Then
'                For l_int_Counter = 0 To lstselectedInterviewers.ListCount - 1
'                    If lstselectedInterviewers.List(l_int_Counter) = lstAllInterviewers.List(l_int_Count) Then
'                        l_bln_existing = True
'                        Exit For
'                    End If
'                Next
'
'                If Not l_bln_existing Then
'                    lstselectedInterviewers.AddItem lstAllInterviewers.List(l_int_Count)
'                    lstAllInterviewers.ListIndex = l_int_Count
'                    f_str_InterviewerName = lstAllInterviewers.Text
'                    l_str_Sql = "SELECT iInterviewerProfileId FROM tbSTEInterviewerProfile" & _
'                        " WHERE vInterviewerName='" & f_str_InterviewerName & "'"
'                    l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
'                    If Not l_obj_Rst.EOF Then
'                        l_int_IntwrId = l_obj_Rst("iInterviewerProfileId")
'                    End If
'                    l_obj_Rst.Close
'                    Set l_obj_Rst = Nothing
'                    f_void_UpdateDatabase (l_int_IntwrId)
'                End If
'                lstAllInterviewers.ListIndex = l_int_Count
'                f_str_InterviewerName = lstAllInterviewers.Text
'            End If
'        Next
'    End If
'
'    Call f_void_AllInterviewers
'    f_bln_Select = False
'    f_void_CheckButtonStatus
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Sub
'
'Private Sub cmdDeselect_Click()
'    'on the click of this button only the interviewer selected from the lstSelectedInterviewers will be
'    'transfered to lstAllInterviewers
'    Dim l_int_Count As Integer              ' counter variable
'    Dim l_int_IntwrId As Integer            ' to store the interviewer id
'    Dim l_str_Sql As String                 ' to store the SQL string
'    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
'
'    On Error GoTo ErrorHandler
'
'    f_bln_DeSelect = True
'        If lstselectedInterviewers.SelCount > 0 Then
'            For l_int_Count = lstselectedInterviewers.ListCount - 1 To 0 Step -1
'                If lstselectedInterviewers.Selected(l_int_Count) Then
'                    lstselectedInterviewers.ListIndex = l_int_Count
'                    f_str_InterviewerName = lstselectedInterviewers.Text
'                    l_str_Sql = "SELECT iInterviewerProfileId FROM tbSTEInterviewerProfile" & _
'                        " WHERE vInterviewerName='" & f_str_InterviewerName & "'"
'                    l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
'                    If Not l_obj_Rst.EOF Then
'                        l_int_IntwrId = l_obj_Rst("iInterviewerProfileId")
'                    End If
'                    l_obj_Rst.Close
'                    Set l_obj_Rst = Nothing
'                    f_void_UpdateDatabase (l_int_IntwrId)
'                    lstselectedInterviewers.RemoveItem l_int_Count
'                End If
'            Next
'        End If
'
'    Call f_void_AllInterviewers
'    f_bln_DeSelect = True
'    f_void_CheckButtonStatus
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Sub
'
'Private Sub cmdDeselectAll_Click()
'    'on the click of this button all the teachers from the lstSelectedTeachers will be removed from
'    'the particular department
'    Dim l_int_InterviewerCount As Integer   ' count of the interviewers
'    Dim l_int_IntwrId As Integer            ' to store the interviewer id
'    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
'    Dim l_str_Sql As String                 ' to store the SQL string
'
'    On Error GoTo ErrorHandler
'
'    f_bln_DeSelectAll = True
'    If lstselectedInterviewers.ListCount >= 1 Then
'       For l_int_InterviewerCount = lstselectedInterviewers.ListCount - 1 To 0 Step -1
'            lstselectedInterviewers.ListIndex = l_int_InterviewerCount
'            f_str_InterviewerName = lstselectedInterviewers.Text
'            l_str_Sql = "SELECT iInterviewerProfileId FROM tbSTEInterviewerProfile" & _
'                " WHERE vInterviewerName='" & f_str_InterviewerName & "'"
'            l_obj_Rst.Open l_str_Sql, g_obj_Conn
'            If Not l_obj_Rst.EOF Then
'                l_int_IntwrId = l_obj_Rst("iInterviewerProfileId")
'            End If
'            l_obj_Rst.Close
'            Set l_obj_Rst = Nothing
'            f_void_UpdateDatabase (l_int_IntwrId)
'            lstselectedInterviewers.RemoveItem l_int_InterviewerCount
'        Next
'    End If
'
'    Call f_void_AllInterviewers
'    f_void_CheckButtonStatus
'    f_bln_DeSelectAll = False
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbInformation, LoadResString(1729)
'End Sub
'
'Public Sub f_void_CheckButtonStatus()
'    'Procedure to check the status of the buttons
'    'i.e enabling and disabling the buttons based on the presense
'    'and selection of data in the list boxes
'
'    If lstAllInterviewers.ListCount = 0 Then
'        ' left side list box is empty
'        cmdSelectall.Enabled = False
'        cmdSelect.Enabled = False
'    Else
'        ' left side list box is not empty
'        cmdSelectall.Enabled = True
'        If lstAllInterviewers.SelCount > 0 Then
'            ' something is selected in the left list box
'            cmdSelect.Enabled = True
'        Else
'            ' no item is selected in left list box
'            cmdSelect.Enabled = False
'        End If
'    End If
'
'    If lstselectedInterviewers.ListCount = 0 Then
'        ' right side list box is empty
'        cmdDeselectall.Enabled = False
'        cmdDeselect.Enabled = False
'    Else
'        ' right side list box is not empty
'        cmdDeselectall.Enabled = True
'        If lstselectedInterviewers.SelCount > 0 Then
'            ' something is selected in the right list box
'            cmdDeselect.Enabled = True
'        Else
'            ' no item is selected in left right box
'            cmdDeselect.Enabled = False
'        End If
'    End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call g_void_CloseChildForm
End Sub
'
'Private Sub lstAllInterviewers_Click()
'    'Enables the cmdselect button when any element in the list box is selected else
'    'button remains disabled
'    f_void_CheckButtonStatus
'End Sub
'
'Private Sub lstAllInterviewers_DblClick()
'    ' double clcicking on an item in the list box should have the same effect as selecting an item and clicking on the ">" or "<" button
'    cmdSelect_Click
'    f_void_CheckButtonStatus
'End Sub
'
'Private Sub lstselectedInterviewers_Click()
'    'Enables the cmddeselect button when any element in the list box is selected else
'    'button remains disabled
'    f_void_CheckButtonStatus
'End Sub

Public Sub f_void_UpdateDatabase(ByVal l_int_IntwrId As Long)
    'Updating the database based on the status of the flags
    Dim l_str_Insert As String                          ' to hold the insert SQL string
    Dim l_str_Sql1 As String                            ' to hold the SQL string
    Dim l_obj_Rst As New ADODB.Recordset                ' recordset object
    Dim l_str_Sql As String
    Dim l_int_SubjectId As Long                      ' to store the subject id for later use
    Dim l_str_Delete As String                          ' to store the delete SQL String
    Dim l_obj_RsdelInterviewer As New ADODB.Recordset   ' to store the id's of those interviewers to be deleted
    'New set of variables
    Dim l_str_subjectSelect As String  'Select
    Dim l_str_subjectInsert As String
    Dim l_obj_rstSubjectInsert As New ADODB.Recordset
    Dim l_obj_rstSubject As New ADODB.Recordset
    Dim l_int_SubjectQuestionProfileID As Long
    Dim l_obj_subjectQuest As New ADODB.Recordset

    On Error GoTo ErrorHandler
    l_str_Sql = "SELECT iSubjectProfileId FROM tbSTESubjectProfile"
    l_str_Sql = l_str_Sql & " WHERE vSubjectName='" & cboSubject.Text & "'"
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    If Not l_obj_Rst.EOF Then
        l_int_SubjectId = l_obj_Rst("iSubjectProfileId")
    End If
    g_obj_Conn.BeginTrans

    If f_bln_SelectAll = True Or f_bln_Select = True Then
        
'        Set f_obj_RsInterviewerID = g_obj_Conn.Execute("SELECT iInterviewRoomProfileId FROM tbSTEInterviewRoomProfile")
'
'        If Not f_obj_RsInterviewerID.EOF Then
'            f_obj_RsInterviewerID.MoveLast
'            f_int_InterviewRoomProfileID = f_obj_RsInterviewerID.Fields(0).Value + 1
'        Else
'            l_str_Sql1 = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='tbSTEInterviewRoomProfile'"
'            Set f_obj_RsInterviewerID = g_obj_Conn.Execute(l_str_Sql1)
'            If Not f_obj_RsInterviewerID.EOF Then
'                f_int_InterviewRoomProfileID = f_obj_RsInterviewerID("iTableCounterIdMapping")
'            Else
'                f_int_InterviewRoomProfileID = 1
'            End If
'            Set f_obj_RsInterviewerID = Nothing
'        End If
Dim bRtn As Boolean
        bRtn = getNewId("tbSTEInterviewRoomProfile", "iInterviewRoomProfileId", f_int_InterviewRoomProfileID)

        Set f_obj_RsInterviewerID = New ADODB.Recordset

        l_str_Insert = "Insert into tbSTEInterViewRoomProfile values(" & _
            f_int_InterviewRoomProfileID & "," & _
            l_int_IntwrId & "," & _
            f_int_RoomProfileID & "," & _
            l_int_SubjectId & "," & _
            cboDay.ListIndex & ",'" & _
            Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
            
        Set f_obj_RsInterviewerID = g_obj_Conn.Execute(l_str_Insert)
        Set f_obj_RsInterviewerID = Nothing
        
        '****************************** changes Mahesh start
'
'        l_str_subjectSelect = "Select * from tbSteSubjectQuestionProfile"
'        l_obj_rstSubject.Open l_str_subjectSelect, g_obj_Conn, adOpenStatic, adLockReadOnly
'
'        If l_obj_rstSubject.RecordCount > 0 Then
'
'            If Not l_obj_rstSubject.EOF Then
'                l_obj_rstSubject.MoveLast
'                l_int_SubjectQuestionProfileID = l_obj_rstSubject.Fields(0).Value + 1
'            Else
'                l_str_Sql1 = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='tbSteSubjectQuestionProfile'"
'                Set l_obj_subjectQuest = g_obj_Conn.Execute(l_str_Sql1)
'                If Not l_obj_subjectQuest.EOF Then
'                    l_int_SubjectQuestionProfileID = l_obj_subjectQuest("iTableCounterIdMapping")
'                Else
'                    l_int_SubjectQuestionProfileID = 1
'                End If
'                Set l_obj_subjectQuest = Nothing
'            End If
'            Set l_obj_rstSubject = Nothing
'
'            l_str_subjectSelect = "Select iInterviewerProfileId from tbSteSubjectQuestionProfile where iSubjectProfileId=" & l_int_SubjectId & " and iInterviewerProfileId =" & l_int_IntwrId
'            l_obj_rstSubject.Open l_str_subjectSelect, g_obj_Conn, adOpenStatic, adLockReadOnly
'            If l_obj_rstSubject.RecordCount = 0 Then
'                l_str_subjectInsert = "Insert into tbSteSubjectQuestionProfile (iSubjectQuestionID,iSubjectProfileId,iInterviewerProfileId,dtCreate,dtUpdate)" _
'                & " values(" & l_int_SubjectQuestionProfileID & "," & l_int_SubjectId & "," & l_int_IntwrId & "," & Format(Date, "MM/DD/YYYY") & "," & Format(Date, "MM/DD/YYYY") & ")"
'                Set l_obj_rstSubjectInsert = g_obj_Conn.Execute(l_str_subjectInsert)
'            End If
'        End If

        '****************************** changes Mahesh End
        
    ElseIf f_bln_DeSelectAll = True Or f_bln_DeSelect = True Then
        'Changes Mahesh S
        Dim l_str_selSubject As String
        Dim l_obj_rstInterviewerRoom As New ADODB.Recordset
        
        'Changes Mahesh E
        l_str_Delete = "SELECT iInterviewRoomProfileId FROM tbSTEInterviewRoomProfile" & _
            " WHERE iRoomProfileId = " & f_int_RoomProfileID & _
            " AND iInterviewerProfileId = " & l_int_IntwrId & _
            " AND iSubjectProfileId = " & l_int_SubjectId & _
            " AND iDayFlag = " & cboDay.ListIndex

        Set l_obj_RsdelInterviewer = g_obj_Conn.Execute(l_str_Delete)
        f_int_InterviewRoomProfileID = l_obj_RsdelInterviewer.Fields("iInterviewRoomProfileId").Value
        Set l_obj_RsdelInterviewer = Nothing
        l_str_Delete = "DELETE from tbSTEInterviewRoomProfile where iInterviewRoomprofileId= " & f_int_InterviewRoomProfileID
        g_obj_Conn.Execute l_str_Delete
'
'        'Changes Mahesh S
'        l_str_selSubject = "Select * from tbSteInterviewRoomProfile where iSubjectProfileId=" & l_int_SubjectId & " and iInterviewerProfileId =" & l_int_IntwrId
'        l_obj_rstInterviewerRoom.Open l_str_selSubject, g_obj_Conn, adOpenStatic, adLockReadOnly
'        If l_obj_rstInterviewerRoom.RecordCount = 0 Then
'            l_str_Delete = "Delete from tbSTESubjectQuestionProfile where iSubjectProfileId=" & l_int_SubjectId & " and iInterviewerProfileId =" & l_int_IntwrId
'            g_obj_Conn.Execute l_str_Delete
'        End If
        'Changes Mahesh S
    End If
    g_obj_Conn.CommitTrans
    Exit Sub
ErrorHandler:
    g_obj_Conn.RollbackTrans
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub
'
'Private Sub lstselectedInterviewers_DblClick()
'    cmdDeselect_Click
'    f_void_CheckButtonStatus
'End Sub

Private Sub f_void_PopulateSubject()
    ' populate the subject combo box
    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    
    On Error GoTo ErrorHandler
    l_str_Sql = "SELECT isubjectprofileid , vSubjectName FROM tbSTESubjectProfile"
    l_str_Sql = l_str_Sql & " WHERE iExamType=2 "
    If m_int_IntRpt = 0 Then
    '面接である科目の抽出
        l_str_Sql = l_str_Sql & " and iSubType in ( 3 , 5 ) "
    Else
    '小論文である科目の抽出
        l_str_Sql = l_str_Sql & " and iSubType = 4 "
    End If
    l_obj_Rst.Open l_str_Sql, g_obj_Conn
    If Not l_obj_Rst.EOF Then
        Do While Not l_obj_Rst.EOF
            cboSubject.AddItem l_obj_Rst("vSubjectName")
            cboSubject.ItemData(cboSubject.NewIndex) = l_obj_Rst("isubjectprofileid")
            l_obj_Rst.MoveNext
        Loop
        cboSubject.ListIndex = 0
    End If
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
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

Private Sub f_void_InitGrid()

Dim ii As Long

    With vsfInterviewerRoom
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
        '.CellTextStyle = "0"
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .GridColor = &H808080
        .Visible = True

        If m_int_IntRpt = 0 Then
        '面接は部屋名称（面接グループ名）なので表示を大きめにする。よって列数少
            .cols = 2 + 8
        Else
        '小論文は乱数（採点グループ）なので表示を小さめにする。よって列数多
            .cols = 2 + 20
        End If
        .Rows = 1

        .Row = 0
        .Col = 0
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .FixedCols = 0
        .FixedRows = 1
        .ColWidth(.Col) = 0
        .Text = IIf(Me.m_int_IntRpt = 0, "InterviewID", "CheckerID")
        .Col = .Col + 1
        .ColWidth(.Col) = 2200
        .Text = IIf(Me.m_int_IntRpt = 0, "面接者名", "採点者名")
        '部屋・乱数設定のカラムの開始場所をprvlRoomStartColに保持
        prvlRoomStartCol = .Col + 1
        For ii = 0 To (((.cols - prvlRoomStartCol) / 2) - 1)
            .Col = .Col + 1
            .ColWidth(.Col) = 0
            .Text = IIf(Me.m_int_IntRpt = 0, "グループ番号", "乱数")
            .Col = .Col + 1
            If m_int_IntRpt = 0 Then
                .ColWidth(.Col) = 1400
            Else
                .ColWidth(.Col) = 560
            End If
            .Text = IIf(Me.m_int_IntRpt = 0, "グループ番号", "乱数")
        Next

    End With

End Sub

Private Sub f_void_PopulateGrid()

Dim sSQL As String
Dim oRs As ADODB.Recordset
Dim sSQL2 As String
Dim oRs2 As ADODB.Recordset
Dim lRow As Long
Dim sWk As String
'add,xzg,2009/12/22,S-----------
Dim strTeacher As String
'add,xzg,2009/12/22,E-----------
    On Error GoTo ErrorHandler
    'add,xzg,2009/12/22,S-----------
    strTeacher = Trim(Me.txtTeacher.Text)
    'add,xzg,2009/12/22,E-----------
    sSQL = "SELECT "
    sSQL = sSQL & "  iInterviewerProfileId "
    sSQL = sSQL & ", vInterviewerName "
    sSQL = sSQL & "  FROM tbSTEInterviewerProfile "
    'add,xzg,2009/12/22,S-----------
    If Len(strTeacher) > 0 Then
    sSQL = sSQL & " WHERE vInterviewerName like '%" & strTeacher & "%'"
    End If
    'add,xzg,2009/12/22,E-----------
    
'    If Me.m_int_IntRpt = 0 Then
''面接官として登録されている
'        sSQL = sSQL & " WHERE siInterviewFlag = 1"
'    Else
''小論文採点者として登録されている
'        sSQL = sSQL & " WHERE siReportFlag = 1"
'    End If

    Set oRs = g_obj_Conn.Execute(sSQL)

    With vsfInterviewerRoom

        .Rows = 1
        lRow = 1
        Do Until oRs.EOF

            sWk = Trim(str(oRs.Fields(0)))
            sWk = sWk & vbTab & oRs.Fields(1)

            If Me.m_int_IntRpt = 0 Then
'面接の場合
                sSQL2 = "SELECT "
                sSQL2 = sSQL2 & "  ir.iRoomProfileID "
                sSQL2 = sSQL2 & ", rp.vRoomName "
                sSQL2 = sSQL2 & "  FROM tbSTEInterviewRoomProfile as ir "
                sSQL2 = sSQL2 & " INNER JOIN tbSTERoomProfile as rp "
                sSQL2 = sSQL2 & "    ON rp.iRoomProfileID = ir.iRoomProfileID "
                sSQL2 = sSQL2 & " WHERE ir.iInterviewerProfileId = " & Trim(str(oRs.Fields(0)))
                sSQL2 = sSQL2 & "   AND iNendo = " & g_int_CurrentNendo
                sSQL2 = sSQL2 & "   AND iDayFlag = " & Trim(str(Me.cboDay.ListIndex))
                sSQL2 = sSQL2 & "   AND iRandomNo is null "
            Else
'小論文の場合
                sSQL2 = "SELECT "
                sSQL2 = sSQL2 & "  iRandomNo "
                sSQL2 = sSQL2 & ", iRandomNo "
                sSQL2 = sSQL2 & "  FROM tbSTEInterviewRoomProfile as ir "
                sSQL2 = sSQL2 & " WHERE ir.iInterviewerProfileId = " & Trim(str(oRs.Fields(0)))
                sSQL2 = sSQL2 & "   AND iNendo = " & g_int_CurrentNendo
                sSQL2 = sSQL2 & "   AND iDayFlag = " & Trim(str(Me.cboDay.ListIndex))
                sSQL2 = sSQL2 & "   AND iRandomNo is not null "
            End If

            Set oRs2 = g_obj_Conn.Execute(sSQL2)
            Do Until oRs2.EOF
                sWk = sWk & vbTab & oRs2.Fields(0)
                sWk = sWk & vbTab & oRs2.Fields(1)
                oRs2.MoveNext
            Loop
            oRs2.Close
            Set oRs2 = Nothing

            .Rows = lRow + 1
            .Col = 0
            .Row = lRow
            .ColSel = .cols - 1
            .RowSel = lRow
            .Clip = sWk

            oRs.MoveNext
            lRow = lRow + 1

        Loop

        oRs.Close
        Set oRs = Nothing

    End With

    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub lstRoom_Click()

    If vsfInterviewerRoom.Col > prvlRoomStartCol And ((vsfInterviewerRoom.Col - prvlRoomStartCol) Mod 2) = 1 Then
        If vsfInterviewerRoom.Row <> 0 Then
            If lstRoom.ListIndex >= 0 Then
                If lfbInterviewerRoomCheck Then
                    vsfInterviewerRoom.TextMatrix(vsfInterviewerRoom.Row, vsfInterviewerRoom.Col) = lstRoom.List(lstRoom.ListIndex)
                    vsfInterviewerRoom.TextMatrix(vsfInterviewerRoom.Row, vsfInterviewerRoom.Col - 1) = lstRoom.ItemData(lstRoom.ListIndex)
                    If vsfInterviewerRoom.Col = vsfInterviewerRoom.cols - 1 Then
                        If vsfInterviewerRoom.Row <> vsfInterviewerRoom.Rows - 1 Then
                            vsfInterviewerRoom.Row = vsfInterviewerRoom.Row + 1
                        End If
                        vsfInterviewerRoom.Col = prvlRoomStartCol
                    Else
                        vsfInterviewerRoom.Col = vsfInterviewerRoom.Col + 2
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Function lfbInterviewerRoomCheck()

Dim lCol As Long
Dim lID As Long
Dim bErr As Boolean
Dim lRow As Long

    bErr = False
    lID = lstRoom.ItemData(lstRoom.ListIndex)
    lRow = vsfInterviewerRoom.Row
    For lCol = prvlRoomStartCol To vsfInterviewerRoom.cols - 2 Step 2
        If vsfInterviewerRoom.TextMatrix(lRow, lCol) = "" Then Exit For
        If vsfInterviewerRoom.TextMatrix(lRow, lCol) = lID Then
            bErr = True
            Exit For
        End If
    Next

    lfbInterviewerRoomCheck = Not bErr

Exit Function
ErrProc:

    lfbInterviewerRoomCheck = False

End Function

Private Sub txtTeacher_Change()
f_void_PopulateGrid
End Sub

Private Sub vsfInterviewerRoom_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        vsfInterviewerRoom.TextMatrix(vsfInterviewerRoom.Row, vsfInterviewerRoom.Col) = ""
        vsfInterviewerRoom.TextMatrix(vsfInterviewerRoom.Row, vsfInterviewerRoom.Col - 1) = ""
    End If

End Sub

