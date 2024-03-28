VERSION 5.00
Begin VB.Form frmPrepSecondExam 
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmPrepSecondExam.frx":0000
   ScaleHeight     =   9990
   ScaleWidth      =   13320
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.TextBox txtTotal 
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
      Left            =   9900
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   8385
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.ComboBox cboSourceRoom 
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
      Left            =   1470
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   2
      Top             =   1515
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.ListBox lstSource 
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
      Height          =   5460
      Left            =   240
      MultiSelect     =   2  'ägí£
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ComboBox cboSourceDay 
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
      Left            =   1485
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   1
      Top             =   1050
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.CommandButton cmdShikenbiFuriwake 
      Caption         =   "ìÒéüééå±ì˙Å@êUï™ é¿çs"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4500
      TabIndex        =   0
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label lblGuidance 
      BackStyle       =   0  'ìßñæ
      Caption         =   "éÛå±ê∂ÉfÅ[É^Ç…ÅAìÒéüééå±ì˙ÇPì˙ñ⁄ÅA2ì˙ñ⁄ÇÃì˙ïtÇê›íËÇµÇ‹Ç∑ÅB"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   3630
      TabIndex        =   12
      Top             =   2010
      Width           =   8085
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'ìßñæ
      Caption         =   "lblMsg"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3645
      TabIndex        =   11
      Top             =   2430
      Width           =   8445
   End
   Begin VB.Label lblTotalDayRoom 
      Caption         =   "çáåv"
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
      Left            =   5370
      TabIndex        =   10
      Top             =   8400
      Visible         =   0   'False
      Width           =   1530
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSourceCapacity 
      BackStyle       =   0  'ìßñæ
      Caption         =   "Label5"
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
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   2595
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label lblExcess 
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
      Height          =   5400
      Left            =   5400
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   5745
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSourceDay 
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
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1125
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'ìßñæ
      Caption         =   "éÛå±î‘çÜ"
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
      TabIndex        =   5
      Top             =   1575
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   1740
   End
End
Attribute VB_Name = "frmPrepSecondExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'Form Name      :   frmPrepSecondExam
'Author         :   Dileep Cherian
'Created On     :   22/10/01
'Description    :   This screen is used to provide mechanism for reshuffling the examinees to
'                   different days and room, which are initially, allocated using the
'                   distribution logic.
'Reference      :   FunctionalSpecs Of Normal Interview.doc(Ver1.0)
'**************************************************************************************************

Option Explicit
Dim f_dt_SourceDay As Date              ' to store the selected source day
Dim f_int_SourceDayMax As Long       ' to store the max capacity of the selected source dayon day
Dim f_int_SourceDayCount As Long     ' count of the existing examinees on the selected source day
Dim f_int_SourceRoomMax As Long      ' max capacity of selected source room
Dim f_int_SourceRoomCount As Long    ' count of existing examinees in selected source room
Dim f_bln_DataChange As Boolean         ' variable to indicate any change operations

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    LoadResStrings Me
    Me.Caption = "frmPrepSecondExam : ééå±ì˙êUï™" ''''LoadResString(2434)

''''Call g_void_SetFontProperties(Me)     ' set the font properties

    lblExcess.Height = 6000
    f_bln_DataChange = False

'    Call l_void_PopulateDayCombo
'    lblExcess.Alignment = 0
'    txtTotal.Text = lstSource.ListCount
'    lblTotalDayRoom.Caption = LoadResString(2488)


'å„ä˙ÇÃèÍçáÅ@ï\é¶ï∂åæÇïœÇ¶ÇÈ 2022.03.09 add jhi
#If zengo_kubun <> 1 Then
    cmdShikenbiFuriwake.Caption = "ìÒéüééå±ì˙ ämíË é¿çs"
    lblGuidance.Caption = "éÛå±ê∂ÉfÅ[É^Ç…ÅAìÒéüééå±ì˙ÇÃì˙ïtÇê›íËÇµÇ‹Ç∑ÅB"
#End If

    lblMsg.Caption = ""


    Exit Sub

ErrorHandler:
    MsgBox Err.Description

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim i As Integer


    fMainForm.mnuTools.Enabled = False  ' disable tools menu


    For i = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next

'    If g_bln_RunLogic Then  ' the distribution logic should be allowed to run only once
'        cmdShikenbiFuriwake.Enabled = False
'    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ÉGÉâÅ["    ''''LoadResString(1729)

End Sub

'*******************************************************************************
'* ìÒéüééå±ì˙êUï™ É{É^Éìèàóù                                                   *
'*******************************************************************************
Private Sub cmdShikenbiFuriwake_Click()

    On Error GoTo ErrorHandler
    Dim rinf   As Long
    Dim strMsg As String


    lblMsg.Caption = ""


#If zengo_kubun = 1 Then
    strMsg = "ìÒéüééå±ì˙ÇÃêUï™èàóùÇé¿çsÇµÇ‹Ç∑ÅBÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH"
#Else
    strMsg = "ìÒéüééå±ì˙ÇÃämíËèàóùÇé¿çsÇµÇ‹Ç∑ÅBÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH"
#End If


    ''''2021.12.15 add jhi
    rinf = myMsgBox(strMsg, "ämîF")
    If rinf = vbCancel Then
        Exit Sub
    End If


'    If Not g_bln_RunLogic Then  ' disable the button, once the logic is run either from normal interview or report

    g_bln_RunLogic = True
    cmdShikenbiFuriwake.Enabled = False

    DoEvents


    Call f_void_ShikenbiFuriwake    'ìÒéüééå±ì˙êUï™Å@èàóùä÷êî
    Call l_void_PopulateDayCombo    'populate the combos and list box with the new data

'    End If

    txtTotal.Text = lstSource.ListCount


#If zengo_kubun = 1 Then
    lblMsg.Caption = "ìÒéüééå±ì˙ÇÃêUï™èàóùÇ™ê≥èÌÇ…äÆóπÇµÇ‹ÇµÇΩÅB" ''''LoadResString(2404)
#Else
    lblMsg.Caption = "ìÒéüééå±ì˙ÇÃämíËèàóùÇ™ê≥èÌÇ…äÆóπÇµÇ‹ÇµÇΩÅB"
#End If



    cmdShikenbiFuriwake.Enabled = True

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ÉGÉâÅ["    ''''LoadResString(1729)

End Sub

'*******************************************************************************
'* ìÒéüééå±ì˙êUï™ main èàóù : uspSTEAutoAllocateExamineeSecondExamDay          *
'* 2021.12.15 cyosa jhi                                                        *
'*******************************************************************************
Private Sub f_void_ShikenbiFuriwake()

    Dim l_obj_Cmd            As New ADODB.Command         ' command object
    Dim l_obj_RstCompound    As New ADODB.Recordset       ' recordset object
    Dim l_obj_rstDay2        As New ADODB.Recordset       ' recordset object
    Dim l_obj_rstDay3        As New ADODB.Recordset       ' recordset object
    Dim l_obj_rstDay4        As New ADODB.Recordset       ' recordset object
    Dim l_obj_rstDay5        As New ADODB.Recordset       ' recordset object
    Dim l_obj_rstDay6        As New ADODB.Recordset       ' recordset object

    Dim l_int_Count          As Long                      ' counter
    Dim l_str_ExcessArray()  As String                    ' to store those examinees whose preference day was violated
    
    ' this stored procedure returns three different recordsets, which are retrieved using the nextrecordset method of ADO
    
    l_int_Count = 0
    Screen.MousePointer = vbHourglass   'change mouse pointer to busystate
        
    l_obj_Cmd.ActiveConnection = g_obj_Conn
    l_obj_Cmd.CommandType = adCmdStoredProc
    l_obj_Cmd.CommandText = "uspSTEAutoAllocateExamineeSecondExamDay"
    
    l_obj_RstCompound.CursorType = adOpenDynamic
    l_obj_RstCompound.LockType = adLockReadOnly
    
    Set l_obj_RstCompound = l_obj_Cmd.Execute
           


'    Do While Not l_obj_RstCompound.EOF
'        ReDim Preserve l_str_ExcessArray(l_int_Count)
'        l_str_ExcessArray(l_int_Count) = l_obj_RstCompound(0)
'        l_int_Count = l_int_Count + 1
'        l_obj_RstCompound.MoveNext
'    Loop
'
'    Set l_obj_rstDay2 = l_obj_RstCompound.NextRecordset    ' set the next recordset
'
'    Do While Not l_obj_rstDay2.EOF
'        ReDim Preserve l_str_ExcessArray(l_int_Count)
'        l_str_ExcessArray(l_int_Count) = l_obj_rstDay2(0)
'        l_int_Count = l_int_Count + 1
'        l_obj_rstDay2.MoveNext
'    Loop
'
'    Set l_obj_rstDay3 = l_obj_RstCompound.NextRecordset    ' set the next recordset
'
'    Do While Not l_obj_rstDay3.EOF
'        ReDim Preserve l_str_ExcessArray(l_int_Count)
'        l_str_ExcessArray(l_int_Count) = l_obj_rstDay3(0)
'        l_int_Count = l_int_Count + 1
'        l_obj_rstDay3.MoveNext
'    Loop
'
'    Set l_obj_rstDay4 = l_obj_RstCompound.NextRecordset    ' set the next recordset
'
'    Do While Not l_obj_rstDay4.EOF
'        ReDim Preserve l_str_ExcessArray(l_int_Count)
'        l_str_ExcessArray(l_int_Count) = l_obj_rstDay4(0)
'        l_int_Count = l_int_Count + 1
'        l_obj_rstDay4.MoveNext
'    Loop
'
'    Set l_obj_rstDay5 = l_obj_RstCompound.NextRecordset    ' set the next recordset
'
'    Do While Not l_obj_rstDay5.EOF
'        ReDim Preserve l_str_ExcessArray(l_int_Count)
'        l_str_ExcessArray(l_int_Count) = l_obj_rstDay5(0)
'        l_int_Count = l_int_Count + 1
'        l_obj_rstDay5.MoveNext
'    Loop
'
'    Set l_obj_rstDay6 = l_obj_RstCompound.NextRecordset    ' set the next recordset
'
'    Do While Not l_obj_rstDay6.EOF
'        ReDim Preserve l_str_ExcessArray(l_int_Count)
'        l_str_ExcessArray(l_int_Count) = l_obj_rstDay6(0)
'        l_int_Count = l_int_Count + 1
'        l_obj_rstDay6.MoveNext
'    Loop
'
'    If l_int_Count > 0 Then
'        lblExcess.Caption = LoadResString(2435)
'        lblExcess.Caption = lblExcess.Caption & vbCrLf & Join(l_str_ExcessArray, ",")
'    End If
'

    Set l_obj_RstCompound = Nothing

'    Set l_obj_rstDay2 = Nothing
'    Set l_obj_rstDay3 = Nothing
'    Set l_obj_rstDay4 = Nothing
'    Set l_obj_rstDay5 = Nothing
'    Set l_obj_rstDay6 = Nothing

    Screen.MousePointer = vbDefault 'restore mouse pointer

End Sub




Private Sub cboSourceDay_Click()    ' for normal interview/report
    Call l_void_PopulateRoomCombo(cboSourceDay.Text)
'    txtTotal.Text = lstSource.ListCount
    Call l_void_PopulateList(f_dt_SourceDay)
    lblSourceCapacity.Caption = CStr(f_int_SourceRoomMax)
    txtTotal.Text = lstSource.ListCount
End Sub

Private Sub cboSourceRoom_Click()   ' for normal interview/report
    Call l_void_PopulateList(f_dt_SourceDay)
    lblSourceCapacity.Caption = CStr(f_int_SourceRoomMax)
    txtTotal.Text = lstSource.ListCount
End Sub


Private Sub l_void_PopulateDayCombo()
    ' populate the sourceday and splday combos
    With cboSourceDay
        .Clear
        .AddItem LoadResString(2424)
        .AddItem LoadResString(2425)
        .AddItem LoadResString(2426)
        .ListIndex = 0
    End With
End Sub

Private Sub l_void_PopulateRoomCombo(ByVal l_str_vDay As String)

    ' fill the room combo based on the day selected in the day combo
    Dim l_obj_Rst As New ADODB.Recordset    ' recordset object
    Dim l_str_Sql As String                 ' SQL string
    Dim l_int_NoOfRooms As Long          ' to store the number of rooms
    Dim l_int_Counter As Long            ' counter
    
    On Error GoTo ErrorHandler
    
    cboSourceRoom.Clear
    
    ' get the current selected day and room, and their capacities
    l_str_Sql = "SELECT iNumberOfRoomDay1, iNumberOfRoomDay2, iNumberOfRoomDay3," & _
        " dtSecondExamDay1, dtSecondExamDay2, dtSecondExamDay3," & _
        " iNumberOfExamineeDay1, iNumberOfExamineeDay2, iNumberOfExamineeDay3" & _
        " FROM tbSTESecondExamProfile" & _
        " WHERE iSystemProfileId = (" & _
        " SELECT top 1 iSystemProfileId FROM tbSTESystemProfile" & _
        " WHERE iActiveFlag = 1)"

    l_obj_Rst.Open l_str_Sql, g_obj_Conn

    If Not l_obj_Rst.EOF Then

'       cmdShikenbiFuriwake.Enabled = True
'        Label4.Visible = True
'        lblSourceCapacity.Visible = True

        Select Case UCase(l_str_vDay)
        Case UCase(LoadResString(2424))     ' day 1
            f_dt_SourceDay = l_obj_Rst("dtSecondExamDay1")
            f_int_SourceDayMax = l_obj_Rst("iNumberOfExamineeDay1")
            l_int_NoOfRooms = l_obj_Rst("iNumberOfRoomDay1")
        Case UCase(LoadResString(2425))     ' day 2
            f_dt_SourceDay = l_obj_Rst("dtSecondExamDay2")
            f_int_SourceDayMax = l_obj_Rst("iNumberOfExamineeDay2")
            l_int_NoOfRooms = l_obj_Rst("iNumberOfRoomDay2")
        Case UCase(LoadResString(2426))     ' day 3
            f_dt_SourceDay = l_obj_Rst("dtSecondExamDay3")
            f_int_SourceDayMax = l_obj_Rst("iNumberOfExamineeDay3")
            l_int_NoOfRooms = l_obj_Rst("iNumberOfRoomDay3")
        End Select
    Else
        cmdShikenbiFuriwake.Enabled = False
        Label4.Visible = False
        lblSourceCapacity.Visible = False

        l_obj_Rst.Close
        Set l_obj_Rst = Nothing
        Exit Sub
    End If

    l_obj_Rst.Close
    Set l_obj_Rst = Nothing

    Exit Sub

ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub l_void_PopulateList(ByVal l_dt_dtDay As Date)

    ' populate the list box based on selection made in the day and room combos
    Dim l_obj_Rst As New ADODB.Recordset        ' recordset object
    Dim l_obj_rst1 As New ADODB.Recordset       ' recordset object
    Dim l_str_Sql As String                     ' SQL string
    
    On Error GoTo ErrorHandler

    lstSource.Clear
    f_int_SourceRoomCount = 0
    
    l_str_Sql = l_str_Sql & " SELECT iExamineeProfileId FROM tbSTEExamineeRoomProfile as er " & _
        " WHERE exists ( select 1 FROM tbSTESubjectProfile sp " & _
        " WHERE iSubType = 3 and er.iSubjectProfileId = sp.iSubjectProfileId )"

    l_obj_Rst.Open l_str_Sql, g_obj_Conn

    Do While Not l_obj_Rst.EOF
        l_str_Sql = "SELECT iJukenNumber, vExamineeName, iPreferenceDay1Flag, iPreferenceDay2Flag, iPreferenceDay3Flag" & _
            " FROM tbSTEExamineeProfile" & _
            " WHERE iExamineeProfileId=" & l_obj_Rst("iExamineeProfileId") & _
            " AND dtSecondExamDay='" & Format(l_dt_dtDay, "MM/DD/YYYY") & "'" & _
            " AND iExamineeStatus = " & gclExamineeStatus_1stPass

        l_obj_rst1.Open l_str_Sql, g_obj_Conn
        If Not l_obj_rst1.EOF Then
            lstSource.AddItem g_str_LPad(l_obj_rst1.Fields("iJukenNumber").Value, Len(l_obj_rst1.Fields("iJukenNumber").Value)) & _
                " - " & l_obj_rst1.Fields("vExamineeName").Value & _
                " - " & l_obj_rst1.Fields("iPreferenceDay1Flag").Value & _
                " - " & l_obj_rst1.Fields("iPreferenceDay2Flag").Value & _
                " - " & l_obj_rst1.Fields("iPreferenceDay3Flag").Value
            f_int_SourceRoomCount = f_int_SourceRoomCount + 1
        End If
        l_obj_rst1.Close
        Set l_obj_rst1 = Nothing
        
        l_obj_Rst.MoveNext
    Loop
    
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing

    f_int_SourceRoomMax = 0
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call g_void_CloseChildForm
End Sub
