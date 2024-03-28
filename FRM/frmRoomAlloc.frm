VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmRoomAlloc 
   BackColor       =   &H8000000A&
   Caption         =   "frmRoomAlloc : 会場入力"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmRoomAlloc.frx":0000
   ScaleHeight     =   9795
   ScaleWidth      =   13785
   WindowState     =   2  '最大化
   Begin VSFlex7LCtl.VSFlexGrid msfRoomAlloc 
      Height          =   6300
      Left            =   225
      TabIndex        =   17
      Top             =   1965
      Width           =   8925
      _cx             =   15743
      _cy             =   11112
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
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
      Height          =   400
      Left            =   240
      TabIndex        =   15
      Top             =   1470
      Width           =   765
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "削除"
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
      Left            =   4740
      TabIndex        =   9
      Top             =   8385
      Width           =   3400
   End
   Begin VB.TextBox txtUnallocatedExaminees 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   10905
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "xtUnallocatedExaminees"
      Top             =   2355
      Width           =   1530
   End
   Begin VB.TextBox txtTotalExaminees 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1890
      Width           =   1530
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
      Left            =   1095
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1470
      Width           =   1710
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
      Height          =   405
      Left            =   2985
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1470
      Width           =   1080
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
      Height          =   400
      Left            =   4380
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1470
      Width           =   1080
   End
   Begin VB.CommandButton cmdAddRoom 
      Caption         =   "グリッドに会場を追加"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6150
      TabIndex        =   6
      Top             =   1455
      Width           =   2985
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "会場の詳細を更新"
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
      Left            =   780
      TabIndex        =   8
      Top             =   8385
      Width           =   3400
   End
   Begin MSFlexGridLib.MSFlexGrid msfRoomAllocOld 
      Height          =   3855
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   8040
      _ExtentX        =   14182
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
   Begin VB.Label lblSerial 
      BackStyle       =   0  '透明
      Caption         =   "行番号"
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
      Left            =   270
      TabIndex        =   16
      Top             =   1185
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "未振分受験生"
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
      Left            =   9345
      TabIndex        =   12
      Top             =   2415
      Width           =   1530
   End
   Begin VB.Label lblTotalNoOfExaminees 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "合計受験生数"
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
      Left            =   9300
      TabIndex        =   10
      Top             =   1980
      Width           =   1560
   End
   Begin VB.Label lblRoomNo 
      BackStyle       =   0  '透明
      Caption         =   "会場名"
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
      Left            =   1485
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblRandomNo 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "乱数"
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
      Left            =   3180
      TabIndex        =   2
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lblCapacity 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "定員"
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
      Left            =   4515
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblErrorDetails 
      BackStyle       =   0  '透明
      Caption         =   "lblErrorDetails"
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
      Height          =   375
      Left            =   195
      TabIndex        =   14
      Top             =   9000
      Width           =   12075
   End
End
Attribute VB_Name = "frmRoomAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmRoomAlloc - 会場入力
'Author         :   Dileep Cherian
'Created On     :   13/9/01
'Update  On     :   2022.01.06
'Description    :   This form makes a provision for allocating examinee to room. This activity is
'                   required before appearing for the first examination.
'Reference      :   FunctionalSpecs OF ROOMALLOCATION.doc(ver 1.1)
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History   -   04/04/2002  -   Dileep Cherian
'User should be able to resize the coulmns, incase part of data is not visible in the normal display
'**************************************************************************************************
'Ammemdments    -   NyushiChangesSummary.doc(ver 1.0)
'Modification History   -   16/05/2002  -   Dileep Cherian
'Delete button should be provided for deleting rows from the grid
'**************************************************************************************************
Dim m_int_ToBeAllotted As Integer           ' examinees yet to be allocated
Dim m_int_CurrentRoom As String             ' to indentify the current room
Dim m_int_AllotedLastRow As Integer

Private m_bChangeOn As Boolean
Private m_bDirty As Boolean

Private prvsCurSerial As String '選択、入力中の訂正前シリアル番号

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Dim l_int_TotalExaminees As Integer                    ' to store the total number of examinees
    Dim l_int_TotalAllotted  As Integer                    ' total examinees already allocated
    Dim l_obj_Rst            As New ADODB.Recordset        ' recordset variable
    Dim sSQL                 As String                     ' SQL string variable

    m_bDirty = False

    LoadResStrings Me
''''Me.Caption = LoadResString(2001) '会場配置

    Call g_void_SetFontProperties(Me)     ' set the font properties


    sSQL = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
        " WHERE iNendo=" & g_int_CurrentNendo & _
        " AND iAbsentFlag = 0" & _
        " AND iExamineeStatus = " & gclExamineeStatus_Default

    l_obj_Rst.Open sSQL, g_obj_Conn, adOpenStatic, adLockReadOnly
    l_int_TotalExaminees = l_obj_Rst.RecordCount
        
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing

    cmdDelete.Enabled = False
    cmdFinish.Enabled = False
    txtUnallocatedExaminees.Text = ""

    If l_int_TotalExaminees > 0 Then
        txtRandomNo.Locked = True
        txtCapacity.Locked = True
        txtTotalExaminees.Text = l_int_TotalExaminees
        m_int_ToBeAllotted = l_int_TotalExaminees
        l_int_TotalAllotted = 0
        txtUnallocatedExaminees = m_int_ToBeAllotted
        Call f_void_InitGrid
    Else
        txtRandomNo.Locked = True
        txtCapacity.Locked = True
        txtTotalExaminees.Text = 0
        txtUnallocatedExaminees = 0
        lblErrorDetails.Caption = LoadResString(2009)
        cmdAddRoom.Enabled = False
    End If

    Call f_void_AddRooms

    If cboRooms.ListCount > 0 Then
        Call f_void_GetAllocation
    End If

'Commented to enable deleting row by making cap as 0 in update mode(data already there)
'    If cboRooms.ListCount > 0 And m_int_ToBeAllotted > 0 Then
'        cmdAddRoom.Enabled = True
'    Else
'        cmdAddRoom.Enabled = False
'    End If

'    txtCapacity.Enabled = False
'    txtRandomNo.Enabled = False


    If cboRooms.ListCount > 0 Then
        cboRooms.ListIndex = 0
        lblErrorDetails.Caption = ""
    Else
        lblErrorDetails.Caption = "割当可能な会場が見つかりませんでした。先に会場を定義してください。" ''''LoadResString(2010)
    End If


   Exit Sub

ErrorHandler:
   MsgBox Err.Description

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim Index As Integer

    fMainForm.mnuTools.Enabled = False  ' disable tools menu

    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub f_void_AddRooms()

    On Error GoTo ErrorHandler

    Dim l_obj_Rst As ADODB.Recordset
    Dim sSQL      As String
    Dim lCnt      As Long

'
'    Set l_obj_Rst = g_obj_Conn.Execute("SELECT count(*) FROM tbSTERoomProfile")
'    If Not l_obj_Rst.EOF Then
'        lCnt = l_obj_Rst.Fields(0)
'    End If
'    l_obj_Rst.Close
'    Set l_obj_Rst = Nothing

    sSQL = "SELECT vRoomName FROM tbSTERoomProfile" & _
        " WHERE iInterviewRoomFlag = 1"             ' changed on 31/07/02

    Set l_obj_Rst = g_obj_Conn.Execute(sSQL)

    Do While Not l_obj_Rst.EOF
        cboRooms.AddItem l_obj_Rst("vRoomName")
        l_obj_Rst.MoveNext
    Loop

''''    If cboRooms.ListCount > 0 Then
''''        cboRooms.ListIndex = 0
''''        lblErrorDetails.Caption = ""
''''    Else
''''        lblErrorDetails.Caption = LoadResString(2010)
''''    End If

    l_obj_Rst.Close
    Set l_obj_Rst = Nothing

    Exit Sub

ErrorHandler:
    MsgBox Err.Description

End Sub


Private Sub SetChange()
    lblErrorDetails.Caption = ""
    lblErrorDetails.Visible = False
'    If m_bChangeOn = False Then m_bDirty = True
End Sub

Private Sub cboRooms_GotFocus()
    m_int_CurrentRoom = cboRooms.ListIndex
End Sub

Private Sub cmdAddRoom_Click()

    On Error GoTo ErrorHandler
    Dim l_str_Sql As String                 ' SQL string variable
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset variable

    m_int_CurrentRoom = cboRooms.ListIndex
    lblErrorDetails.Caption = ""

    If Trim(txtRandomNo.Text) <> "" Then
        l_str_Sql = "SELECT iRoomProfileId FROM tbSTERoomProfile"
        l_str_Sql = l_str_Sql & " WHERE iRandom=" & CInt(Trim(txtRandomNo.Text))
        l_str_Sql = l_str_Sql & " AND vRoomName != '" & cboRooms.Text & "'"

        l_obj_Rst.Open l_str_Sql, g_obj_Conn
        If Not l_obj_Rst.EOF Then
            lblErrorDetails.Caption = LoadResString(2006)
            txtRandomNo.SetFocus
            txtRandomNo.SelStart = 0
            txtRandomNo.SelLength = Len(txtRandomNo.Text)
            Exit Sub
        End If
        l_obj_Rst.Close
        Set l_obj_Rst = Nothing
    Else
        lblErrorDetails.Caption = LoadResString(2007)
        txtRandomNo.SetFocus
        Exit Sub
    End If

    If Trim(txtCapacity.Text) = "" Then
        lblErrorDetails.Caption = "定員は入力必須です。" ''''LoadResString(2008)
        txtCapacity.SetFocus
        Exit Sub
    End If

    If txtSerial.Text <> "" Then
        If Not gf_IntCheck(txtSerial.Text) Then
            lblErrorDetails.Caption = "数値もしくは空白を入力してください。"
            txtCapacity.SetFocus
            Exit Sub
        End If
        If CInt(txtSerial.Text) < 1 Or CInt(txtSerial.Text) >= msfRoomAlloc.Rows Then
            txtSerial.Text = ""
'            lblErrorDetails.Caption = "シリアルの最大値以内で入力してください。"
'            txtCapacity.SetFocus
'            Exit Sub
        Else
Dim iRtn As Integer
            iRtn = MsgBox("シリアル番号" & txtSerial.Text & "番に指定データを挿入します。" & vbCrLf & _
                            "よろしければ「はい」" & vbCrLf & _
                            "最後尾に登録ならば「いいえ」" & vbCrLf & _
                            "登録を中止するならば「キャンセル」を押してください。", vbYesNoCancel, "登録確認")
            If iRtn = vbCancel Then Exit Sub
            If iRtn = vbNo Then txtSerial.Text = ""
        End If
    End If

    If txtCapacity.Text <> "0" Then
Dim lIndex As Long
Dim sItem As String
Dim lRow As Long
Dim lCurRow As Long
        lCurRow = -1

        With msfRoomAlloc
            For lRow = 1 To .Rows - 1
                .Row = lRow
                If UCase(Trim(.TextMatrix(lRow, 1))) = UCase(Trim(cboRooms.List(m_int_CurrentRoom))) Then
                    lCurRow = lRow
                    Exit For
                End If
            Next

            If lCurRow > 0 Then
            '登録済みエラー
                lblErrorDetails.Caption = "指定の会場は登録済みです。"
                txtRandomNo.SetFocus
                Exit Sub
            End If

            For lRow = 1 To .Rows - 1
                .Row = lRow
                If UCase(Trim(.TextMatrix(lRow, 0))) = UCase(Trim(txtSerial.Text)) Then
                    lCurRow = lRow
                    Exit For
                End If
            Next

'            If lCurRow > 0 Then
'                .Row = lCurRow
'                .RowSel = lCurRow
'                .Col = 0
'                .ColSel = .Cols - 1
'                sItem = .Clip
'                lIndex = txtSerial.Text
'                .AddItem sItem, lIndex
'            End If
        End With
    End If

'    g_obj_Conn.BeginTrans

'    l_str_Sql = "UPDATE tbSTERoomProfile SET iRandom=" & txtRandomNo.Text & "," & _
        " iMaxCapacity=" & txtCapacity.Text & "," & _
        " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
        " WHERE vRoomName='" & cboRooms.Text & "'"
'    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
'    g_obj_Conn.Execute l_str_Sql

'    Set l_obj_Rst = Nothing

    m_bDirty = True

    Call f_void_PopulateGrid
'    g_obj_Conn.CommitTrans
    If msfRoomAlloc.Rows > 1 Then cmdFinish.Enabled = True

    Exit Sub

ErrorHandler:
'    g_obj_Conn.RollbackTrans
    MsgBox Err.Description, vbInformation

End Sub

Private Sub cmdDelete_Click()
    ' delete the selected row and set its max capacity as 0
    ' bycalling the cmdAddRoom_click procedure
    Dim l_int_Answer As Integer

''''l_int_Answer = MsgBox(LoadResString(1122), vbQuestion + vbYesNo)
    l_int_Answer = MsgBox("選択しましたレコードを削除しますか？", vbQuestion + vbYesNo)

    If l_int_Answer = vbYes Then
'        txtCapacity.Text = 0
        m_bDirty = True
        msfRoomAlloc.RemoveItem msfRoomAlloc.Row
        Call l_void_ReNewJukenNo
'        Call cmdAddRoom_Click
'        Call cmdFinish_Click
        cmdFinish.Enabled = True
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub cmdFinish_Click()
    ' update the chages made in the grid
    Dim l_int_Counter As Integer            ' counter variable
    Dim l_str_RoomName As String            ' to store the roomname
    Dim l_int_startJuken As Integer         ' to store the starting juken number
    Dim l_int_EndJuken As Integer           ' to store the ending juken number
    Dim l_str_Sql As String                 ' SQL string
    
    On Error GoTo ErrorHandler
    
    With msfRoomAlloc
'        If .Rows <= 1 Then Exit Sub
        l_str_RoomName = ""
        l_int_startJuken = 0
        l_int_EndJuken = 0
        g_obj_Conn.BeginTrans
        For l_int_Counter = 1 To .Rows - 1
            .Row = l_int_Counter
            If .TextMatrix(.Row, 4) = "-" Then Exit For 'ここからは振り分ける人がいない設定

            .Col = 1
            l_str_RoomName = .Text

            .Col = 4
            l_int_startJuken = .Text

            .Col = 5
            l_int_EndJuken = .Text

            l_str_Sql = "UPDATE tbSTEExamineeProfile SET" & _
                " iRoomProfileId=(" & _
                " SELECT iRoomProfileId FROM tbSTERoomProfile" & _
                " WHERE vRoomName='" & l_str_RoomName & "')" & _
                " ,dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                " WHERE iJukenNumber BETWEEN " & l_int_startJuken & " AND " & l_int_EndJuken & _
                " AND iNendo=" & g_int_CurrentNendo & _
                " AND iExamineeStatus = " & gclExamineeStatus_Default & _
                " AND iAbsentFlag = 0"
            
           g_obj_Conn.Execute (l_str_Sql)
        Next
    End With
    
    l_str_Sql = "UPDATE tbSTEExamineeProfile SET" & _
                " iRoomProfileId=NULL" & _
                " ,dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                " WHERE iJukenNumber >" & l_int_EndJuken & _
                " AND iNendo=" & g_int_CurrentNendo & _
                " AND iExamineeStatus = " & gclExamineeStatus_Default & _
                " AND iAbsentFlag = 0"
    g_obj_Conn.Execute (l_str_Sql)

'ログを残すため、tbSTEexamineeRoomに書き込む。
'まずは以前のデータを削除
    l_str_Sql = " delete from tbSTEExamineeRoomProfile "
    l_str_Sql = l_str_Sql & " where exists ( select 1 from tbSTEExamineeProfile as ep"
    l_str_Sql = l_str_Sql & "               where ep.iNendo = " & g_int_CurrentNendo
    l_str_Sql = l_str_Sql & "                 and ep.iExamineeProfileId = tbSTEExamineeRoomProfile.iExamineeProfileId )  "
    l_str_Sql = l_str_Sql & " and exists ( select 1 from tbSTEroomProfile as rp"
    l_str_Sql = l_str_Sql & "               where rp.iInterviewRoomFlag = 1 "
    l_str_Sql = l_str_Sql & "                 and rp.iRoomProfileId = tbSTEExamineeRoomProfile.iRoomProfileId )  "
    g_obj_Conn.Execute (l_str_Sql)

'tbSTEexamineeRoomに書き込む。
    l_str_Sql = " uspSTEIns1stExamRoom " & g_int_CurrentNendo
    g_obj_Conn.Execute (l_str_Sql)

    lblErrorDetails.Caption = LoadResString(2404)
    g_obj_Conn.CommitTrans
    m_bDirty = False
    Exit Sub
ErrorHandler:
    g_obj_Conn.RollbackTrans
    MsgBox Err.Description
End Sub


Private Function f_long_getTotalExsamineeCnt() As Long

Dim l_str_Sql As String
Dim l_obj_Rst As ADODB.Recordset

    f_long_getTotalExsamineeCnt = -1

    l_str_Sql = "SELECT count(*) FROM tbSTEExamineeProfile" & _
        " WHERE iNendo=" & g_int_CurrentNendo & _
        " AND iAbsentFlag = 0" & _
        " AND iExamineeStatus = " & gclExamineeStatus_Default
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql, adOpenStatic, adLockReadOnly)

    If Not l_obj_Rst.EOF Then
        f_long_getTotalExsamineeCnt = l_obj_Rst.Fields(0)
    End If

    l_obj_Rst.Close
    Set l_obj_Rst = Nothing

End Function

Private Sub f_void_InitGrid()

    ' initializes the grid with it sheaders, column width etc
    On Error GoTo ErrorHandler

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
    
        .Rows = 2
 'Mah
        '.Cols = 6
        .cols = 8
'Mahe
        .FixedRows = 1
        .FixedCols = 0
        
        .Row = 0
        .Col = 0                                ' serial number
        .ColWidth(0) = 800
        .CellAlignment = flexAlignCenterCenter
        .Text = "行番号" ''''LoadResString(1756)
        
        .Col = .Col + 1                         ' room name
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(1) = 1800
        .Text = "会場名"    ''''LoadResString(1503)
        
        .Col = .Col + 1                         ' random number
        .ColWidth(2) = 1400
        .CellAlignment = flexAlignCenterCenter
        .Text = "乱数"    ''''LoadResString(1504)
        
        .Col = .Col + 1                         ' max capacity
        .ColWidth(3) = 1400
        .CellAlignment = flexAlignCenterCenter
        .Text = "定員" ''''LoadResString(1505)
        
        .Col = .Col + 1                         ' starting juken number
        .ColWidth(4) = 1600
        .CellAlignment = flexAlignCenterCenter
        .Text = "開始受験番号" ''''LoadResString(2012)
        
        .Col = .Col + 1                         ' ending juken number
        .ColWidth(5) = 1600
        .CellAlignment = flexAlignCenterCenter
        .Text = "終了受験番号" ''''LoadResString(2013)

        'Mah
        .Col = .Col + 1
        .ColWidth(.Col) = 0
        .Col = .Col + 1
        .ColWidth(.Col) = 0
        'Mahe
    End With

    Exit Sub

ErrorHandler:
    MsgBox Err.Description

End Sub

Private Sub cboRooms_Click()

    Dim l_str_Sql As String                 ' SQL string
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset variable
           
    On Error GoTo ErrorHandler
           
    l_str_Sql = "SELECT iRandom, iMaxCapacity FROM tbSTERoomProfile WHERE vRoomName='" & cboRooms.Text & "'"
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)

    If IsNull(l_obj_Rst("iMaxCapacity")) And m_int_ToBeAllotted = 0 Then
        lblErrorDetails.Caption = LoadResString(2011)
        cboRooms.ListIndex = m_int_CurrentRoom
        Exit Sub
    End If
    
    If Not l_obj_Rst.EOF Then
        If Trim(l_obj_Rst("iRandom")) <> "" Then
            txtRandomNo.Text = l_obj_Rst("iRandom")
        Else
            txtRandomNo.Text = 0
        End If
        
        If Trim(l_obj_Rst("iMaxCapacity")) <> "" Then
            txtCapacity.Text = l_obj_Rst("iMaxCapacity")
        Else
            txtCapacity.Text = 0
        End If
    End If
    
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing
    lblErrorDetails.Caption = ""    ' clear the error label once the room is changed

    Exit Sub

ErrorHandler:
    MsgBox Err.Description

End Sub

Private Sub f_void_PopulateGrid()
    'populates the grid with the resulting rows
    Dim l_str_sqlRoomDetails As String              ' sql string to pick the room details
    Dim l_obj_rsRoomDetails As New ADODB.Recordset  ' recordset variable to pick the room details
    Dim l_int_RowCounter As Integer                 ' counter variable
    Dim l_int_ColCounter As Integer                 ' counter variable
    Dim l_obj_rsJukenNo As New ADODB.Recordset      ' recordset variable to pick the valid examinees
    Dim l_str_sqlJukenNo As String                  ' sql string to pick the valid examinees
    Dim l_int_CurRoomMax As Integer                 ' store the max capacity of current room
    Dim l_int_startJuken As Integer                 ' store the current rooms start juken number
    Dim l_int_EndJuken As Integer                   ' store the current rooms end juken number
    Dim l_bln_RoomExists As Boolean                 ' check whether the room is already added to the grid or not
    Dim l_int_oldMaxCapacity As Integer             ' store the old max capacity before updating with the new one
    Dim l_str_sqlRooms As String                    ' sql string to pick the room details
    Dim l_obj_rsRooms As New ADODB.Recordset        ' recordset variable to pick the room details
    Dim l_str_sqlUpdate As String                   ' to update the removed roomid's
    Dim l_int_SrNo As Integer                   ' to store the starting serial number
    Dim l_bln_Delete As Boolean
    Dim m_int_cascade As Integer                'to store difference betn the new and old capacities of the room in question

    l_int_startJuken = 0
    l_str_sqlRoomDetails = "SELECT vRoomName,iRandom,iMaxCapacity FROM tbSTERoomProfile " & _
        " WHERE vRoomName='" & Trim(cboRooms.Text) & "'"

    l_obj_rsRoomDetails.Open l_str_sqlRoomDetails, g_obj_Conn, adOpenStatic, adLockReadOnly
    
    If l_obj_rsRoomDetails.EOF Then Exit Sub     ' exit, if the selected room does not exist

    With msfRoomAlloc
        ' loop through the grid rows to find whether the room is already
        ' allocated to the grid
        For l_int_RowCounter = 1 To .Rows - 1    'For all rows in grid
            If l_int_RowCounter > .Rows - 1 Then
                g_void_RefreshGrid msfRoomAlloc
                Exit Sub
            End If
            .Row = l_int_RowCounter
            .Col = 1
            'If the room in combo is the one in the grid?
            If UCase(Trim(.Text)) = UCase(Trim(cboRooms.List(m_int_CurrentRoom))) Or l_bln_RoomExists Then
                ' the room is already there in the grid, so update the room details in grid
                l_str_sqlRooms = "SELECT vRoomName,iRandom,iMaxCapacity FROM tbSTERoomProfile " & _
                    " WHERE vRoomName='" & Trim(.Text) & "'"
                l_obj_rsRooms.Open l_str_sqlRooms, g_obj_Conn, adOpenStatic, adLockReadOnly
                                                
                If Not l_bln_RoomExists Then   'If NOT for the first time, take juken from grid
                    .Col = 4
                    l_int_startJuken = .Text - 1
                End If
                                                                
                .Col = 2
                .Text = l_obj_rsRooms.Fields("iRandom").Value
                
                .Col = 3
                l_int_oldMaxCapacity = .Text
                l_int_CurRoomMax = l_obj_rsRooms.Fields("iMaxCapacity").Value
               
                .Text = l_int_CurRoomMax
                ' form the sql string to get the valid examinees to be allocated to this room
                l_str_sqlJukenNo = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber FROM tbSTEExamineeProfile" & _
                    " WHERE iNendo=" & g_int_CurrentNendo & _
                    " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iAbsentFlag = 0" & _
                    " AND iJukenNumber > " & l_int_startJuken & " ORDER BY iJukenNumber"
                    
                l_obj_rsJukenNo.Open l_str_sqlJukenNo, g_obj_Conn, adOpenStatic, adLockReadOnly
                
                If l_obj_rsJukenNo.EOF Then
                    Exit Sub    ' exit, if there are no valid examinees
                End If
                .Col = 4
                If l_bln_RoomExists Then
                    .Text = l_obj_rsJukenNo.Fields("iJukenNumber").Value
                End If

                .Col = 6
                If l_int_oldMaxCapacity <> l_int_CurRoomMax Then   'If the capacity has changed(only for the row where it was changed by user)
                    If CInt(.Text) = 0 And l_int_CurRoomMax >= l_obj_rsJukenNo.RecordCount Then    'If last row and the current capacity accomodates
                        m_int_ToBeAllotted = 0
                    ElseIf CInt(.Text) = 0 And l_int_CurRoomMax < l_obj_rsJukenNo.RecordCount Then
                        m_int_ToBeAllotted = l_obj_rsJukenNo.RecordCount - l_int_CurRoomMax
                    Else
                        m_int_ToBeAllotted = CInt(.Text) + (l_int_oldMaxCapacity - l_int_CurRoomMax)
                        m_int_cascade = l_int_oldMaxCapacity - l_int_CurRoomMax 'the difference in capacities
                    End If
                Else 'l_int_oldMaxCapacity <> l_int_CurRoomMax (this is next room)
                    .Col = 6
                    If CInt(.Text) = 0 Then
                        .Col = 7
                        If l_int_oldMaxCapacity > CInt(.Text) + m_int_cascade Then
                            m_int_ToBeAllotted = 0
                        Else
                            m_int_ToBeAllotted = (.Text + m_int_cascade) - l_int_oldMaxCapacity
                        End If
                    Else
                        m_int_ToBeAllotted = .Text + m_int_cascade
                    End If
                End If 'l_int_oldMaxCapacity <> l_int_CurRoomMax
                If l_int_CurRoomMax > 0 Then
                    If l_obj_rsJukenNo.RecordCount > l_int_CurRoomMax Then
                        ' allocate that amny examinees to the room, as its capacity
                        l_obj_rsJukenNo.Move l_int_CurRoomMax - 1
                    Else
                        ' allocate all remaining examinees to this room, as the capacity of this room is greater than total valid examinees to be allocated
                        l_obj_rsJukenNo.MoveLast
                        m_int_AllotedLastRow = l_obj_rsJukenNo.RecordCount
                        ' here the examinees to be allocated can go to a negative value also
                        ' we need to keep it as such, otherwise, if the capacity changes again,
                        ' this it might lead to inconsistency in the 'to be allocated' value.
                    End If
                    .Col = 5
                    l_int_EndJuken = l_obj_rsJukenNo.Fields("iJukenNumber").Value
                    .Text = l_int_EndJuken
                    .Col = 6
                    .Text = IIf(m_int_ToBeAllotted < 0, 0, m_int_ToBeAllotted)
                    .Col = 7
                    .Text = l_obj_rsJukenNo.RecordCount
                Else 'l_int_CurRoomMax > 0  (user made the cap 0)
                    l_str_sqlUpdate = "UPDATE tbSTEExamineeProfile SET" & _
                        " iRoomProfileId=NULL, " & _
                        " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                        " WHERE iJukenNumber BETWEEN "
                    .Col = 4
                    l_str_sqlUpdate = l_str_sqlUpdate & .Text & " AND "
                    .Col = 5
                    l_str_sqlUpdate = l_str_sqlUpdate & .Text & _
                        " AND iNendo=" & g_int_CurrentNendo & _
                        " AND iExamineeStatus = " & gclExamineeStatus_Default & _
                        " AND iAbsentFlag = 0"
                    g_obj_Conn.Execute (l_str_sqlUpdate)

                    If .Rows = 2 Then
                        .Rows = .Rows + 1
                        l_bln_Delete = True
                    End If
                    .Col = 4
                    l_int_EndJuken = .Text - 1
                    If .Row = .Rows - 1 Then
                        .Col = 7
                        m_int_ToBeAllotted = .Text
                    End If
                      .RemoveItem (l_int_RowCounter)
                    If l_bln_Delete Then
                        .Rows = .Rows - 1
                    End If
                    l_int_RowCounter = l_int_RowCounter - 1
                End If 'l_int_CurRoomMax > 0
                
                ' check if the 'to be allocated' value is less than 0 or not
                ' if less than 0, display it as 0 for the user
                If m_int_ToBeAllotted <= 0 Then
                    txtUnallocatedExaminees = "0"
                Else
                    txtUnallocatedExaminees = m_int_ToBeAllotted
                End If
                ' code to remove rows from grid, incase changing the capacity of a room in the grid
                ' is making the rows below that as ineffective
                If m_int_ToBeAllotted <= 0 Then
                    .Rows = .Row + 1
                    Exit Sub
                End If
                
                l_int_startJuken = l_int_EndJuken
                
                l_obj_rsJukenNo.Close
                Set l_obj_rsJukenNo = Nothing
                
                l_obj_rsRooms.Close
                Set l_obj_rsRooms = Nothing
                l_bln_RoomExists = True     ' set varibale to indicate that the room already exists
            End If 'UCase(Trim(.Text)) = UCase(Trim(cboRooms.List(m_int_CurrentRoom))) Or l_bln_RoomExists
        Next
        
        If m_int_ToBeAllotted <= 0 Then
            lblErrorDetails.Caption = LoadResString(2482)
            Exit Sub
        Else
            lblErrorDetails.Caption = ""
        End If
        
        If Trim(txtCapacity.Text) = 0 Then
            lblErrorDetails.Caption = LoadResString(2496)
        Else
            lblErrorDetails.Caption = ""
        End If
        
        If Not l_bln_RoomExists _
            And l_obj_rsRoomDetails.Fields("iMaxCapacity").Value <> 0 _
                And m_int_ToBeAllotted > 0 Then
            ' the room is not yet allocated in the grid
            ' add the room details to the grid and allocate examinees to that room
            If .Rows = 1 Then
                .Rows = 2
            End If
            .Row = .Rows - 1
            .Col = 0
            If Trim(.Text) = "" Then
                ' there are no rows in the grid
                ' this is going to be the first row
                l_int_SrNo = 1                  ' set the serial number as 1
                l_int_startJuken = 0            ' start from the very first jiken number available
                .Row = .Rows - 1
            Else
                ' some rows are already there
                ' add this as the last row
                l_int_SrNo = Trim(.Text) + 1    ' add 1 to the last existing serial number
                .Col = 5
                l_int_startJuken = .Text        ' start from the last juken number allocated + 1
Dim bMoveCheck As Boolean
                bMoveCheck = False
                If txtSerial.Text <> "" Then
                    If gf_IntCheck(txtSerial.Text) Then
                        If CInt(txtSerial.Text) Then
                            bMoveCheck = True
                        End If
                    End If
                End If
                If bMoveCheck Then
                    .AddItem "", CInt(txtSerial.Text)
                    .Row = CInt(txtSerial.Text)
                Else
                    .Rows = .Rows + 1               ' add one additional row for this room
                    .Row = .Rows - 1
                End If
            End If

            .Col = 0
            .Text = l_int_SrNo
            
            .Col = 1
            .Text = l_obj_rsRoomDetails.Fields("vRoomName").Value
            
            .Col = 2
            .Text = l_obj_rsRoomDetails.Fields("iRandom").Value
            
            .Col = 3
            l_int_CurRoomMax = l_obj_rsRoomDetails.Fields("iMaxCapacity").Value
            .Text = l_int_CurRoomMax
            
            ' form the sql string to get the valid examinees to be allocated to this room
            l_str_sqlJukenNo = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber FROM tbSTEExamineeProfile" & _
                " WHERE iNendo=" & g_int_CurrentNendo & _
                " AND iExamineeStatus = " & gclExamineeStatus_Default & _
                " AND iAbsentFlag = 0" & _
                " AND iJukenNumber > " & l_int_startJuken & _
                " ORDER BY iJukenNumber"
            l_obj_rsJukenNo.Open l_str_sqlJukenNo, g_obj_Conn, adOpenStatic, adLockReadOnly
            
            If l_obj_rsJukenNo.EOF Then
                Exit Sub    ' exit, if there are no valid examinees
            End If
            l_int_startJuken = l_obj_rsJukenNo.Fields("iJukenNumber").Value
                            
            .Col = 4
            .Text = l_obj_rsJukenNo.Fields("iJukenNumber").Value
            
            If l_obj_rsJukenNo.RecordCount > l_int_CurRoomMax Then
                ' allocate that amny examinees to the room, as its capacity
                l_obj_rsJukenNo.Move l_int_CurRoomMax - 1
                m_int_ToBeAllotted = m_int_ToBeAllotted - l_int_CurRoomMax  ' reduce the count of number of examinees to be allocated
            Else
                ' allocate all remaining examinees to this room, as the capacity of this room is greater than total valid examinees to be allocated
                l_obj_rsJukenNo.MoveLast
                m_int_AllotedLastRow = l_obj_rsJukenNo.RecordCount
                m_int_ToBeAllotted = m_int_ToBeAllotted - l_int_CurRoomMax
            End If
            
            If m_int_ToBeAllotted <= 0 Then
                txtUnallocatedExaminees = 0
            Else
                txtUnallocatedExaminees = m_int_ToBeAllotted
            End If
            
            .Col = 5
            .Text = l_obj_rsJukenNo.Fields("iJukenNumber").Value
            'Mah
            .Col = 6
            .Text = IIf(m_int_ToBeAllotted < 0, 0, m_int_ToBeAllotted)
            .Col = 7
            .Text = l_obj_rsJukenNo.RecordCount
            'Mahe
            l_obj_rsJukenNo.Close
            Set l_obj_rsJukenNo = Nothing
        End If
        If bMoveCheck Then
            Call l_void_ReNewJukenNo  'シリアル振りなおし
        End If
    End With

    l_obj_rsRoomDetails.Close
    Set l_obj_rsRoomDetails = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If m_bDirty Then
        If vbCancel = MsgBox("入力後、保存されていません。" & vbCrLf & "保存せず終了してもよろしいですか？", vbOKCancel) Then
            Cancel = 1
        Else
            Call g_void_CloseChildForm
        End If
    Else
        Call g_void_CloseChildForm
    End If
End Sub

Private Sub msfRoomAlloc_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim sSerial As String
Dim lRow As Long
Dim lCurRow As Long
Dim lChangeRow As Long
Dim sCurrent As String
Dim sChange As String
    If Col <> 0 Then Exit Sub 'シリアルのみ変更可
    sSerial = msfRoomAlloc.TextMatrix(Row, Col)
    If Not IsNumeric(sSerial) Then
        MsgBox "数値で入力してください。", vbOKOnly Or vbInformation, "入力エラー"
        msfRoomAlloc.TextMatrix(Row, Col) = prvsCurSerial
        Exit Sub
    End If
'入力されたシリアル番号をグリッドから探す
    lChangeRow = -1
    For lRow = 1 To msfRoomAlloc.Rows - 1
        If lRow <> Row And sSerial = msfRoomAlloc.TextMatrix(lRow, 0) Then
            lChangeRow = lRow
            Exit For
        End If
    Next
    If lChangeRow <> -1 Then
        lCurRow = Row
        With msfRoomAlloc
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
            Call l_void_ReNewJukenNo
        End With
'        Call cmdAddRoom_Click
    End If
    cmdFinish.Enabled = True
End Sub

Private Sub l_void_ReNewJukenNo()
    '受験番号範囲を再計算する
Dim sSQL As String
Dim oRs As ADODB.Recordset
Dim lStartNo As Long
Dim lRow As Long
Dim l_int_TotalExaminees As Long
Dim lsvRow As Long

    m_int_ToBeAllotted = f_long_getTotalExsamineeCnt   '残り人数に最大人数を再入
    For lRow = 1 To msfRoomAlloc.Rows - 1
        lsvRow = lRow
    'シリアルの書き換えもする
        msfRoomAlloc.TextMatrix(lRow, 0) = Trim(str(lRow))
        lStartNo = CLng(IIf(IsNumeric(msfRoomAlloc.TextMatrix(lRow - 1, 5)), msfRoomAlloc.TextMatrix(lRow - 1, 5), 0))
        sSQL = "SELECT min( dbo.usfMakeDispJukenNumber(iJukenNumber) ) as smJukenNumber ," & _
            " max( dbo.usfMakeDispJukenNumber(iJukenNumber) ) as mxJukenNumber " & _
            " FROM ( SELECT top " & Trim(str(msfRoomAlloc.TextMatrix(lRow, 3))) & _
            " iJukenNumber FROM tbSTEExamineeProfile" & _
            " WHERE iNendo=" & g_int_CurrentNendo & _
            " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iAbsentFlag = 0" & _
            " AND iJukenNumber > " & Trim(str(lStartNo)) & " ORDER BY iJukenNumber ) as t1 "
        Set oRs = g_obj_Conn.Execute(sSQL)
        If Not oRs.EOF Then
            lStartNo = oRs.Fields(0)
            msfRoomAlloc.TextMatrix(lRow, 4) = oRs.Fields(0)
            lStartNo = oRs.Fields(1)
            msfRoomAlloc.TextMatrix(lRow, 5) = oRs.Fields(1)
            oRs.Close
            m_int_ToBeAllotted = m_int_ToBeAllotted - msfRoomAlloc.TextMatrix(lRow, 3)  '残り人数を割振っただけ減算
        End If
        Set oRs = Nothing
        '残人数を書き換える
        txtUnallocatedExaminees.Text = IIf(m_int_ToBeAllotted <= 0, "0", Trim(str(m_int_ToBeAllotted)))
        If m_int_ToBeAllotted <= 0 Then Exit For
    Next

    If lsvRow <> msfRoomAlloc.Rows - 1 Then
        For lRow = lsvRow + 1 To msfRoomAlloc.Rows - 1
        'シリアルの書き換えもする
            msfRoomAlloc.TextMatrix(lRow, 0) = Trim(str(lRow))
            msfRoomAlloc.TextMatrix(lRow, 4) = "-"
            msfRoomAlloc.TextMatrix(lRow, 5) = "-"
        Next
    End If

End Sub

Private Sub msfRoomAlloc_Click()
    ' populate the form fields with the selected room
    ' also highlight the currently selected row
    Dim l_int_RowCounter As Integer
    Dim l_int_ColCounter As Integer
    Dim l_int_CurRow As Integer
    Dim l_int_CurCol As Integer
                
    With msfRoomAlloc
        If .Rows <= 1 Then Exit Sub     ' exit if there are no rows in the grid
        If .Row < 1 Then Exit Sub     ' exit if there are no rows in the grid
        l_int_CurRow = .Row
        l_int_CurCol = .Col
        .Col = 0
        txtSerial.Text = .Text
        .Col = 1
        cboRooms.Text = .Text
        .BackColorSel = &HFFFFFF  'white
        .FocusRect = 1   'flexFocusNone
        For l_int_RowCounter = 1 To .Rows - 1
            .Row = l_int_RowCounter
            For l_int_ColCounter = 0 To .cols - 1
                .Col = l_int_ColCounter
                If .Row <> l_int_CurRow Then
                    .CellBackColor = &HFFFFFF
                Else
                    .CellBackColor = &HC0C0FF
                End If
            Next
        Next
        .BackColorSel = &H800000  'normal (blue)
        .FocusRect = 0   'flexFocusHeavy
        .Row = l_int_CurRow
        .Col = l_int_CurCol
    End With
    cmdDelete.Enabled = True
End Sub

Private Sub msfRoomAlloc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then msfRoomAlloc_Click
End Sub

Private Sub msfRoomAlloc_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 1 Then
        Cancel = True
        Exit Sub
    End If
    prvsCurSerial = Trim(msfRoomAlloc.TextMatrix(Row, Col))
    If prvsCurSerial = "" Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub txtCapacity_KeyPress(KeyAscii As Integer)
    ' allow only integer values
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRandomNo_KeyPress(KeyAscii As Integer)
    ' allow only integer values
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub f_void_GetAllocation()

    On Error GoTo ErrorHandler


    Dim l_obj_rsRooms As New ADODB.Recordset            ' recordset to hold the roomprofileid
    Dim l_str_sqlRoomDetails As String                  ' to get the room details
    Dim l_obj_rsRoomDetails As New ADODB.Recordset      ' reocrdset to get the room details
    Dim l_str_JukenNo As String                         ' to get the range of juken numbers
    Dim l_obj_rsJukenNo As New ADODB.Recordset          ' recordset for the juken numbers
    Dim l_obj_rsTotalExaminees As New ADODB.Recordset   ' recordset for the entire examinees
    Dim l_int_AllocatedExaminees As Integer             ' store the number of examinees aslready allocated
    Dim l_int_SrNo As Integer                           ' to hold the serial nuber variable
    Dim l_int_Examinees As Integer                      ' total number of eligible examinees
    Dim l_int_ExamineesInRoomCount As Integer           ' total number of examinees in a particular room

    Dim sSQL_TotalExaminees As String                   ' to get the total number of examinees
    Dim sSQL_Rooms          As String                   ' to form the roomprofileid query
    
    sSQL_TotalExaminees = "SELECT iJukenNumber FROM tbSTEExamineeProfile" & _
                " WHERE iNendo=" & g_int_CurrentNendo & _
                " AND iExamineeStatus = " & gclExamineeStatus_Default & _
                " AND iAbsentFlag = 0" & _
                " ORDER BY iJukenNumber"
    l_obj_rsTotalExaminees.Open sSQL_TotalExaminees, g_obj_Conn, adOpenStatic, adLockReadOnly
    l_int_Examinees = l_obj_rsTotalExaminees.RecordCount
                
    If Not l_obj_rsTotalExaminees.EOF Then
    
        msfRoomAlloc.Rows = 2
        
        Do While Not l_obj_rsTotalExaminees.EOF

            With msfRoomAlloc
            
                ' get the room profile id where the examinee belongs
                sSQL_Rooms = "SELECT iRoomProfileId FROM tbSTEExamineeProfile" & _
                    " WHERE ijukenNumber =" & l_obj_rsTotalExaminees.Fields("iJukenNumber").Value & _
                    " AND iRoomProfileId IS NOT NULL" & _
                    " AND iNendo=" & g_int_CurrentNendo

                l_obj_rsRooms.Open sSQL_Rooms, g_obj_Conn
                
                If Not l_obj_rsRooms.EOF Then
                    l_int_SrNo = l_int_SrNo + 1
                    .Row = l_int_SrNo
                    
                    .Col = 0
                    .Text = l_int_SrNo
                                
                    ' get details of that room
                    l_str_sqlRoomDetails = " SELECT vRoomName, iRandom, iMaxCapacity FROM tbSTERoomProfile" & _
                        " WHERE iRoomProfileId=" & l_obj_rsRooms.Fields("iRoomProfileId").Value

                    l_obj_rsRoomDetails.Open l_str_sqlRoomDetails, g_obj_Conn
                    
                    .Col = 1
                    .Text = l_obj_rsRoomDetails.Fields("vRoomName").Value
                    
                    .Col = 2
                    .Text = l_obj_rsRoomDetails.Fields("iRandom").Value
                    
                    .Col = 3
                    .Text = l_obj_rsRoomDetails.Fields("iMaxCapacity").Value
                    
                    l_str_JukenNo = "SELECT iJukenNumber FROM tbSTEExamineeProfile" & _
                        " WHERE iRoomProfileId=" & l_obj_rsRooms.Fields("iRoomProfileId").Value & _
                        " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iAbsentFlag = 0" & _
                        " AND iNendo=" & g_int_CurrentNendo & _
                        " ORDER BY iJukenNumber"
                    l_obj_rsJukenNo.Open l_str_JukenNo, g_obj_Conn, adOpenStatic, adLockReadOnly
                    l_int_ExamineesInRoomCount = l_obj_rsJukenNo.RecordCount
                                
                    l_int_AllocatedExaminees = l_int_AllocatedExaminees + l_obj_rsJukenNo.RecordCount
                    
                    .Col = 4
                    .Text = l_obj_rsJukenNo.Fields("iJukenNumber").Value
                    
                    l_obj_rsJukenNo.MoveLast
                    .Col = 5
                    .Text = l_obj_rsJukenNo.Fields("iJukenNumber").Value
                    'MAh
                    'Total examinees-allocated in this room
                    .Col = 6
                    .Text = IIf(l_int_Examinees - l_int_AllocatedExaminees < 0, 0, l_int_Examinees - l_int_AllocatedExaminees)

                    .Col = 7
                    .Text = l_int_ExamineesInRoomCount
                    'Mahe

                    l_obj_rsJukenNo.Close
                    Set l_obj_rsJukenNo = Nothing
                    
                    l_obj_rsRoomDetails.Close
                    Set l_obj_rsRoomDetails = Nothing
                    
                    l_obj_rsTotalExaminees.Move l_int_ExamineesInRoomCount
                    
                    .Rows = .Rows + 1
                Else
                    l_obj_rsTotalExaminees.MoveNext
                End If
                
                l_obj_rsRooms.Close
                Set l_obj_rsRooms = Nothing
            End With
        Loop

        msfRoomAlloc.Rows = msfRoomAlloc.Rows - 1
        m_int_ToBeAllotted = l_int_Examinees - l_int_AllocatedExaminees
        txtUnallocatedExaminees.Text = IIf(m_int_ToBeAllotted <= 0, "0", Trim(str(m_int_ToBeAllotted)))
    End If
    
    l_obj_rsTotalExaminees.Close
    Set l_obj_rsTotalExaminees = Nothing

    Exit Sub

ErrorHandler:
    MsgBox Err.Description

End Sub

Private Sub txtSerial_KeyPress(KeyAscii As Integer)

    Call NumericOnly(Me, KeyAscii)

End Sub
