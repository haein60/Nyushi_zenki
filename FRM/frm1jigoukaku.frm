VERSION 5.00
Begin VB.Form frm1jigoukaku 
   BackColor       =   &H00FFFFFF&
   Caption         =   "frm1jigoukaku : çáäié“ì¸óÕ(1éüééå±)"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm1jigoukaku.frx":0000
   ScaleHeight     =   9810
   ScaleWidth      =   14280
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.CommandButton cmdJukenList 
      Caption         =   "1éü éÛå±ê∂ÉäÉXÉgCSVèoóÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   210
      TabIndex        =   24
      Top             =   7770
      Width           =   2800
   End
   Begin VB.CommandButton cmdGoukakuList 
      Caption         =   "1éü çáäié“ÉäÉXÉgCSVèoóÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7080
      TabIndex        =   23
      Top             =   7770
      Width           =   2800
   End
   Begin VB.TextBox txtJuTotal 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   4680
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   21
      Top             =   7395
      Width           =   930
   End
   Begin VB.ComboBox cboRoom 
      Height          =   300
      Left            =   3795
      TabIndex        =   20
      Text            =   "cboRoom"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cboRoomID 
      Height          =   300
      Left            =   5400
      TabIndex        =   19
      Text            =   "cboRoomID"
      Top             =   1095
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cboSubject 
      Height          =   300
      Left            =   1785
      TabIndex        =   18
      Text            =   "cboSubject"
      Top             =   1110
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.TextBox txtSourceName 
      Height          =   375
      Left            =   10335
      TabIndex        =   17
      Text            =   "txtSourceName"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtDestName 
      Height          =   375
      Left            =   11625
      TabIndex        =   16
      Text            =   "txtDestName"
      Top             =   1065
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11520
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   15
      Top             =   7390
      Width           =   930
   End
   Begin VB.TextBox txtDestJuken 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   400
      Left            =   8625
      TabIndex        =   1
      Top             =   1650
      Width           =   1400
   End
   Begin VB.TextBox txtSourceJuken 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   400
      Left            =   2310
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1650
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "1éü çáäié“ ämíË"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5250
      TabIndex        =   8
      Top             =   8505
      Width           =   2205
   End
   Begin VB.CommandButton cmdDeselectall 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      TabIndex        =   7
      Top             =   5610
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      TabIndex        =   6
      Top             =   5010
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      TabIndex        =   5
      Top             =   4410
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectall 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      TabIndex        =   4
      Top             =   3810
      Width           =   1215
   End
   Begin VB.ListBox lstSelected 
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4935
      Left            =   7080
      MultiSelect     =   2  'ägí£
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2430
      Width           =   5370
   End
   Begin VB.ListBox lstExaminees 
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4935
      Left            =   240
      MultiSelect     =   2  'ägí£
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2430
      Width           =   5370
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'ìßñæ
      Caption         =   "éÛå±ê∂êî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   225
      TabIndex        =   22
      Top             =   7395
      Width           =   1200
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'ìßñæ
      Caption         =   "çáäié“êî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   7095
      TabIndex        =   14
      Top             =   7390
      Width           =   1200
   End
   Begin VB.Label Label4 
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
      Height          =   330
      Left            =   7095
      TabIndex        =   13
      Top             =   1725
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'ìßñæ
      Caption         =   "çáäié“Å@éÛå±î‘çÜ"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1725
      Width           =   1980
   End
   Begin VB.Label lblErrorDetails 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'ìßñæ
      Caption         =   "lblErrorDetails"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   540
      TabIndex        =   11
      Top             =   9270
      Width           =   12015
   End
   Begin VB.Label lblSelectFrom 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "1éü éÛå±ê∂ÉäÉXÉg"
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
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2145
      Width           =   5355
   End
   Begin VB.Label lblSelected 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "1éü çáäié“ÉäÉXÉg"
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
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   2145
      Width           =   5355
   End
End
Attribute VB_Name = "frm1jigoukaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*************************************************************************************************
'Form Name      : frm1jigoukaku
'Author         : jhi
'Created On     : 2021.12.29
'Description    : 1éüééå± - çáäié“ì¸óÕ
'Reference      :
'***************************************************************************************************

Private f_bln_SelectAll   As Boolean    'Shows the status of the Select All button
Private f_bln_Select      As Boolean    'Shows the status of the Select  button
Private f_bln_DeSelect    As Boolean    'Shows the status of the DeSelectAll button
Private f_bln_DeSelectAll As Boolean    'Shows the status of the DeSelect  button

Dim f_bln_DataChanged     As Boolean    'to enable/disable the save button
Dim f_bln_FormLoaded      As Boolean    'to check whether form is loaded or not

Public m_int_IntRpt       As Long       'form variable variable which indicated whether the form has to be instantiated for the "interview" or "report"
Public m_int_Action       As Long       'determine the action to be performed

''''Private Const prvcSubName_Language As String = "äOçëåÍ"     ''''2021.12.14 del jhi(äOçëåÍÇÕÇ»Ç¢)
Private Const prvcSubName_Language     As String = "âpåÍ"       ''''2021.12.14 add jhi(äOçëåÍ->âpåÍÇ…ïœçX)
Private Const prvcSubName_Science      As String = "óùâ»"
Private Const prvcSubName_SecondExam   As String = "ÇQéüééå±"

'*******************************************************************************
'* 1éüééå± - åáê»é“ì¸óÕ                                                        *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler


    LoadResStrings Me
''''Call g_void_SetFontProperties(Me)     ' set the font properties

    lblErrorDetails.Caption = ""

    f_bln_DataChanged = False
    f_bln_FormLoaded = False
    

    m_int_Action = 1    ''''1éüééå±:çáäié“ListÇã≠êßìIÇ…ê›íË

    '---------------------------------------------------------------------------
    '1éüééå±:çáäié“List
    '---------------------------------------------------------------------------
    Select Case m_int_Action
    Case 1
 
'       Me.Caption = "frm1jigoukaku : çáäié“ì¸óÕ"        ''''LoadResString(1013)
'       lblSelectFrom.Caption = "éÛå±ê∂ÉäÉXÉg"           ''''LoadResString(2408)
'       lblSelected.Caption = "çáäié“ÉäÉXÉg"             ''''LoadResString(2410)
 
        Label4.Visible = False
        txtDestJuken.Visible = False

'        Label1.Caption = "éÛå±î‘çÜ"
'        lblTotal.Caption = "çáäié“êî"
'        cmdOK.Caption = "1éü çáäié“ ämíË"

    End Select

    lstExaminees.Font = "ÇlÇr ÉSÉVÉbÉN"
    lstSelected.Font = "ÇlÇr ÉSÉVÉbÉN"

    lstExaminees.FontSize = 10
    lstSelected.FontSize = 10


    cmdDeselect.Enabled = False
    cmdSelect.Enabled = False

    Call f_void_CheckButtonStatus


    '---------------------------------------------------------------------------
    ' 1éüééå± çáäié“ì¸óÕÅ@ä÷êîÇåƒÇ—èoÇ∑
    '---------------------------------------------------------------------------
    Call f_void_Select


    txtJuTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount

    f_bln_FormLoaded = True

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ÉGÉâÅ[" ''''LoadResString(1729)

End Sub

Private Sub f_void_Select()

    Dim oRs As ADODB.Recordset    ' recordset object
    Dim sSQL As String             ' The SQL string
    Dim l_str_DisplayString As String   ' to form the display string in the list box
        
    lstExaminees.Clear
    lstSelected.Clear
        
    sSQL = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName,iSex FROM tbSTEExamineeProfile WHERE" & _
        " iNendo = " & g_int_CurrentNendo & _
        " AND iAbsentFlag = 0"
    
    Select Case m_int_Action
    Case 1   ' 1st exam
        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default
    Case 3, 4    ' 2nd exam
        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    Case 5  ' enter/refuse phase
        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_2ndWait
    Case 6  ' enter/refuse phase
        sSQL = sSQL & " AND (iExamineeStatus = " & gclExamineeStatus_2ndPass & " or iExamineeStatus = " & gclExamineeStatus_2ndWaitPass & ") and iRejectFlag = 0"
    End Select
        
    Set oRs = g_obj_Conn.Execute(sSQL)
    
'    If oRs.EOF Then
'        Set oRs = Nothing
'        Exit Sub
'    End If
    Do While Not oRs.EOF
        l_str_DisplayString = oRs.Fields("iJukenNumber").Value & _
            " - " & oRs.Fields("vExamineeName").Value
        If oRs.Fields("iSex").Value = 0 Then
            l_str_DisplayString = l_str_DisplayString & " - (*)"
        Else
            l_str_DisplayString = l_str_DisplayString & "      "
        End If

        lstExaminees.AddItem l_str_DisplayString
        oRs.MoveNext
    Loop
    Set oRs = Nothing
    
    sSQL = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName,iSex FROM tbSTEExamineeProfile  WHERE"
    sSQL = sSQL & " iNendo = " & g_int_CurrentNendo
    
    Select Case m_int_Action
    Case 1  ' input passed person data
        sSQL = sSQL & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    Case 3  ' passed person data for 2nd phase
        sSQL = sSQL & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_2ndPass
    Case 4  ' waiting list
        sSQL = sSQL & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_2ndWait
    Case 5  ' upliftment from waiting list
        sSQL = sSQL & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_2ndWaitPass
    Case 6  ' enter/refuse offer
        sSQL = sSQL & " AND iAbsentFlag = 0" & _
            " AND iRejectFlag = 1" & _
            " AND iExamineeStatus IN (" & gclExamineeStatus_2ndPass & "," & gclExamineeStatus_2ndWaitPass & ")"
    End Select
        
    Set oRs = g_obj_Conn.Execute(sSQL)
    
    If oRs.EOF Then
        Set oRs = Nothing
        Exit Sub
    End If
    Do While Not oRs.EOF
        l_str_DisplayString = oRs.Fields("iJukenNumber").Value & " - " & oRs.Fields("vExamineeName").Value
        If oRs.Fields("iSex").Value = 0 Then
            l_str_DisplayString = l_str_DisplayString & " - (*)"
        Else
            l_str_DisplayString = l_str_DisplayString & "      "
        End If
        lstSelected.AddItem l_str_DisplayString
        oRs.MoveNext
    Loop
    
    Set oRs = Nothing

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorHandler

    Dim l_int_Count              As Long                   ' counter
    Dim l_int_TempJuken          As Long                   ' to store the juken number
    Dim l_str_JukenNo            As String                 ' to store all the selected juken numbers as a string
    Dim l_str_NonSelected        As String                 ' to store all the non-selected juken numbers as a string
    Dim l_str_ExamineeID         As String                 ' string of examinee id's
    Dim l_obj_Rec                As ADODB.Recordset        ' recordset variable
    Dim l_str_Sql                As String                 ' to store the SQl string
    Dim l_str_MySql              As String
    Dim l_obj_Rst                As New ADODB.Recordset    ' recordset variable
    Dim l_obj_rst1               As New ADODB.Recordset
    Dim l_obj_rst2               As New ADODB.Recordset
    Dim l_obj_rst3               As New ADODB.Recordset
    Dim l_obj_rst4               As New ADODB.Recordset
    Dim l_str_ExamineeIDSql      As String                 ' to store the SQL string
    Dim l_int_subjectProfileId   As Long                   ' to store the subject profile Id
    Dim l_int_NewScoreProfileId  As Long                   ' to store the score profile Id
    Dim l_str_Sql1               As String                 ' to store the SQL string
    Dim l_str_sql2               As String

    Dim bRtn As Boolean



    ' É}ÉEÉXÉ|ÉCÉìÉ^ÇçªéûåvÇ…ÇµÇ‹Ç∑ÅB
    Screen.MousePointer = vbHourglass
    lblErrorDetails.Caption = "çáäié“ÇÃçXêVèàóùÇçsÇ¡ÇƒÇ¢Ç‹Ç∑ÅBÇµÇŒÇÁÇ≠Ç®éùÇøÇ≠ÇæÇ≥Ç¢ÅB"
    cmdOK.Enabled = False

    DoEvents


  
    ' get all the examinees in selected list box into a single string
    For l_int_Count = 0 To lstSelected.ListCount - 1
        l_int_TempJuken = Left(lstSelected.List(l_int_Count), 4)
        l_str_JukenNo = l_str_JukenNo & "," & l_int_TempJuken
    Next

    If Len(Trim(l_str_JukenNo)) > 0 Then
        l_str_JukenNo = Right(Trim(l_str_JukenNo), Len(Trim(l_str_JukenNo)) - 1)
    End If
    
    ' get all the examinees in non-selected examinees(left) list box into a single string
    For l_int_Count = 0 To lstExaminees.ListCount - 1
        l_int_TempJuken = Left(lstExaminees.List(l_int_Count), 4)
        l_str_NonSelected = l_str_NonSelected & "," & l_int_TempJuken
    Next

    If Len(Trim(l_str_NonSelected)) > 0 Then
        l_str_NonSelected = Right(Trim(l_str_NonSelected), Len(Trim(l_str_NonSelected)) - 1)
    End If
    
    If lstSelected.ListCount > 0 Or lstExaminees.ListCount > 0 Then
        
        g_obj_Conn.BeginTrans   ' start a transaction as there are multiple database table inserts/updates
        
        Select Case m_int_Action
        Case 1
            
            ' input passed person data for 1st exam
            If Len(l_str_JukenNo) > 0 Then
                l_str_Sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_1stPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_JukenNo & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = 0"
                
                g_obj_Conn.Execute l_str_Sql
            End If
            
            ' set the status back to 0, in case someone is moved from right to left
            If Len(l_str_NonSelected) > 0 Then
                l_str_Sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_Default & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_NonSelected & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = 1"
                
                g_obj_Conn.Execute l_str_Sql
            End If
            
        End Select
        
        g_obj_Conn.CommitTrans
        
        If f_bln_DataChanged Then
            f_bln_DataChanged = False
            cmdOK.Enabled = False
        End If

        lblErrorDetails.Caption = "çáäié“ÇÃçXêVèàóùÇ™ê≥èÌÇ…èIóπÇµÇ‹ÇµÇΩÅB"    ''''LoadResString(2404)

    End If

    ' É}ÉEÉXÉ|ÉCÉìÉ^ÇãKíËílÇ…ñﬂÇµÇ‹Ç∑ÅB
    Screen.MousePointer = vbDefault

    Exit Sub

ErrorHandler:
    g_obj_Conn.RollbackTrans
    lblErrorDetails.Caption = "1éüçáäié“ÇÃçXêVèàóùÇ≈ÉGÉâÅ[Ç™î≠ê∂ÇµÇ‹ÇµÇΩÅB" ''''LoadResString(2405)
    MsgBox Err.Description, vbInformation, "ÉGÉâÅ["


End Sub

'*******************************************************************************
'On the click of this button all the Examinees from the lstExaminees will be
'transfered to lstSelectedInterviewers
'*******************************************************************************
Private Sub cmdSelectAll_Click()

    On Error GoTo ErrorHandler

    Dim l_int_Examinees As Long

    
    f_bln_SelectAll = True
    
    lblErrorDetails.Caption = ""
    If lstExaminees.ListCount >= 1 Then
        For l_int_Examinees = lstExaminees.ListCount - 1 To 0 Step -1
            lstSelected.AddItem lstExaminees.List(l_int_Examinees)
            lstExaminees.ListIndex = l_int_Examinees
            lstExaminees.RemoveItem l_int_Examinees
        Next
    End If

    f_void_CheckButtonStatus
    f_bln_SelectAll = False
    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtJuTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'*******************************************************************************
'on the click of this button only the Examinee selected from the lstExaminees will be transfered to
'lstSelected
'*******************************************************************************
Private Sub cmdSelect_Click()

    On Error GoTo ErrorHandler

    Dim l_int_Count As Long

    
    f_bln_Select = True
    lblErrorDetails.Caption = ""

    If lstExaminees.SelCount > 0 Then
        For l_int_Count = lstExaminees.ListCount - 1 To 0 Step -1
            If lstExaminees.Selected(l_int_Count) Then
                lstSelected.AddItem lstExaminees.List(l_int_Count)
                lstExaminees.RemoveItem l_int_Count
            End If
        Next
    End If

    f_void_CheckButtonStatus
    f_bln_Select = False

    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtJuTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'on the click of this button only the interviewer selected from the lstSelected will be
'transfered to lstExaminees
Private Sub cmdDeselect_Click()

    On Error GoTo ErrorHandler

    Dim l_int_Count As Long
    
    lblErrorDetails.Caption = ""

    f_bln_DeSelect = True

    If lstSelected.SelCount > 0 Then
        For l_int_Count = lstSelected.ListCount - 1 To 0 Step -1
            If lstSelected.Selected(l_int_Count) Then
                lstExaminees.AddItem lstSelected.List(l_int_Count)
                lstSelected.RemoveItem l_int_Count
            End If
        Next
    End If

    f_void_CheckButtonStatus
    f_bln_DeSelect = True

    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtJuTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'on the click of this button all the Examinees from the lstSelectedInterviewers will be moved to
'LstAllinterviewers
Private Sub cmdDeselectAll_Click()

    On Error GoTo ErrorHandler
    Dim l_int_InterviewerCount As Long
  
  
    lblErrorDetails.Caption = ""

    f_bln_DeSelectAll = True

    If lstSelected.ListCount >= 1 Then
       For l_int_InterviewerCount = lstSelected.ListCount - 1 To 0 Step -1
            lstExaminees.AddItem lstSelected.List(l_int_InterviewerCount)
            lstSelected.RemoveItem l_int_InterviewerCount
        Next
    End If

    f_void_CheckButtonStatus
    f_bln_DeSelectAll = False

    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtJuTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub
'
'Private Sub l_SetUpdateButtonEnabled()
'
'    If dd Then
'    End If
'
'End Sub

Public Sub f_void_CheckButtonStatus()

    If lstExaminees.ListCount = 0 Then
        cmdSelectall.Enabled = False
        cmdSelect.Enabled = False
    Else
        cmdSelectall.Enabled = True
        If lstExaminees.SelCount > 0 Then
            cmdSelect.Enabled = True
        Else
            cmdSelect.Enabled = False
        End If
    End If
    
    If lstSelected.ListCount = 0 Then
        cmdDeselectall.Enabled = False
        cmdDeselect.Enabled = False
    Else
        cmdDeselectall.Enabled = True
        If lstSelected.SelCount > 0 Then
            cmdDeselect.Enabled = True
        Else
            cmdDeselect.Enabled = False
        End If
    End If

End Sub

Private Sub lstExaminees_Click()
    'Enables the cmdselect button when any element in the list box is selected else
    'button remains disabled
    f_void_CheckButtonStatus
End Sub

Private Sub lstExaminees_DblClick()
    cmdSelect_Click
    f_void_CheckButtonStatus
End Sub

Private Sub lstSelected_Click()
    'Enables the cmddeselect button when any element in the list box is selected else
    'button remains disabled
    f_void_CheckButtonStatus
End Sub

Private Sub lstSelected_DblClick()
    cmdDeselect_Click
    f_void_CheckButtonStatus
End Sub


Private Sub txtDestJuken_KeyPress(KeyAscii As Integer)
    ' move the input juken number from the non-selected listbox to the selected listbox
    Dim l_str_sqlExaminee As String             ' to form the examinee details query string
    Dim l_obj_rsExaminee As New ADODB.Recordset ' to hold the examinee details records
    Dim l_str_JukenNo As String                 ' to sotre the input juken number
    Dim l_int_counter1 As Long               ' to loop through the list box
    Dim l_int_counter2 As Long               ' to loop through the list box
    
    On Error GoTo ErrorHandler
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
        
    If KeyAscii = 13 Then
        
        If Trim(txtDestJuken.Text) = "" Then Exit Sub     ' exit if the textbox is empty
        
        ' enable the functionality only for the "Enter/Return key"
        l_str_sqlExaminee = "SELECT iJukenNumber, substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName FROM tbSTEExamineeProfile" & _
            " WHERE iJukenNumber=" & Trim(txtDestJuken.Text) & " AND iNendo=" & g_int_CurrentNendo
        l_obj_rsExaminee.Open l_str_sqlExaminee, g_obj_Conn
        
            
        If l_obj_rsExaminee.EOF Then
            ' the input juken number does not exist - display an error message
            lblErrorDetails.Caption = LoadResString(2473)
        Else
            lblErrorDetails.Caption = ""
            ' pad the input juken number with leading "0"
            l_str_JukenNo = g_str_LPad(Trim(txtDestJuken.Text), Len(Trim(txtDestJuken.Text)))
            
            For l_int_counter1 = 0 To lstSelected.ListCount - 1
                ' loop through the list box to check whether the juken number is present or not
                If Left(lstSelected.List(l_int_counter1), 4) = l_str_JukenNo Then
                    ' input juken is presnet
                    
                    ' display examinee name in the neme text box
                    txtDestName.Text = l_obj_rsExaminee.Fields("vExamineeName").Value
                    
                    ' make it the selected one
                    lstSelected.Selected(l_int_counter1) = True
                    
                    ' move it to the non-selected listbox
                    lblErrorDetails.Caption = ""
                    f_bln_DeSelect = True
                        
                    lstExaminees.AddItem lstSelected.List(l_int_counter1)
                    lstSelected.RemoveItem l_int_counter1
                                
                    f_void_CheckButtonStatus
                    f_bln_DeSelect = True
                    If Not f_bln_DataChanged Then
                        f_bln_DataChanged = True
                        cmdOK.Enabled = True
                    End If

                    txtJuTotal.Text = lstExaminees.ListCount
                    txtTotal.Text = lstSelected.ListCount
                    
                    ' loop thourh the nonselected listbox, and highlight the input juken number
                    For l_int_counter2 = 0 To lstExaminees.ListCount - 1
                        If Left(lstExaminees.List(l_int_counter2), 4) = l_str_JukenNo Then
                            lstExaminees.Selected(l_int_counter2) = True
                        Else
                            lstExaminees.Selected(l_int_counter2) = False
                        End If
                    Next
                    txtDestJuken.Text = ""
                    txtDestName.Text = ""
                    Exit Sub
                End If
            Next
        End If
        l_obj_rsExaminee.Close
        Set l_obj_rsExaminee = Nothing
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub txtSourceJuken_KeyPress(KeyAscii As Integer)
    ' move the input juken number from the selected listbox to the non-selected listbox
    Dim l_str_sqlExaminee As String             ' to form the examinee details query string
    Dim l_obj_rsExaminee As New ADODB.Recordset ' to hold the examinee details records
    Dim l_str_JukenNo As String                 ' to sotre the input juken number
    Dim l_int_counter1 As Long               ' to loop through the list box
    Dim l_int_counter2 As Long               ' to loop through the list box
    
    On Error GoTo ErrorHandler
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
        
    If KeyAscii = 13 Then
        
        If Trim(txtSourceJuken.Text) = "" Then Exit Sub     ' exit if the textbox is empty
        
        ' enable the functionality only for the "Enter/Return key"
        l_str_sqlExaminee = "SELECT iJukenNumber, substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName FROM tbSTEExamineeProfile" & _
            " WHERE iJukenNumber=" & Trim(txtSourceJuken.Text) & " AND iNendo=" & g_int_CurrentNendo
        l_obj_rsExaminee.Open l_str_sqlExaminee, g_obj_Conn
            
        If l_obj_rsExaminee.EOF Then
            ' the input juken number does not exist - display an error message
            lblErrorDetails.Caption = LoadResString(2473)
        Else
            lblErrorDetails.Caption = ""
            ' pad the input juken number with leading "0"
            l_str_JukenNo = g_str_LPad(Trim(txtSourceJuken.Text), Len(Trim(txtSourceJuken.Text)))
            
            ' loop through the list box to check whether the juken number is present or not
            For l_int_counter1 = 0 To lstExaminees.ListCount - 1
                If Left(lstExaminees.List(l_int_counter1), 4) = l_str_JukenNo Then
                    ' input juken is presnet
                    
                    ' display examinee name in the name text box
                    txtSourceName.Text = l_obj_rsExaminee.Fields("vExamineeName").Value
                    
                    ' make it the selected one
                    lstExaminees.Selected(l_int_counter1) = True
                    
                    ' move it to the selected listbox
                    f_bln_Select = True
                    lblErrorDetails.Caption = ""
                    
                    lstSelected.AddItem lstExaminees.List(l_int_counter1)
                    lstExaminees.RemoveItem l_int_counter1
                           
                    f_void_CheckButtonStatus
                    f_bln_Select = False
                    If Not f_bln_DataChanged Then
                        f_bln_DataChanged = True
                        cmdOK.Enabled = True
                    End If

                    txtJuTotal.Text = lstExaminees.ListCount
                    txtTotal.Text = lstSelected.ListCount
                    
                    ' loop thourh the selected listbox, and highlight the input juken number
                    For l_int_counter2 = 0 To lstSelected.ListCount - 1
                        If Left(lstSelected.List(l_int_counter2), 4) = l_str_JukenNo Then
                            lstSelected.Selected(l_int_counter2) = True
                        Else
                            lstSelected.Selected(l_int_counter2) = False
                        End If
                    Next
                    txtSourceJuken.Text = ""
                    txtSourceName.Text = ""
                End If
            Next
            
        End If
        l_obj_rsExaminee.Close
        Set l_obj_rsExaminee = Nothing
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_SelectAbsentee()

    Dim l_obj_Rst As ADODB.Recordset    ' recordset object
    Dim l_str_Sql As String             ' The SQL string
    Dim l_str_DisplayString As String   ' to form the display string in the list box
    Dim l_str_sqlRoomName As String
    Dim l_obj_rsRoomName As New ADODB.Recordset
    
    lstExaminees.Clear
    lstSelected.Clear

    l_str_Sql = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName,iSex,iRoomProfileId"
    l_str_Sql = l_str_Sql & " FROM tbSTEExamineeProfile WHERE iNendo = " & g_int_CurrentNendo
    l_str_Sql = l_str_Sql & " AND iExamineeProfileId NOT IN("
    l_str_Sql = l_str_Sql & " SELECT iExamineeProfileId FROM tbSTEScoreProfile"
    l_str_Sql = l_str_Sql & " WHERE iSubjectProfileId in (SELECT iSubjectProfileId FROM tbSTESubjectProfile"

    Select Case Trim(cboSubject.Text)
    Case prvcSubName_Science
        l_str_Sql = l_str_Sql & " WHERE vSubjectName in ('ï®óù' , 'âªäw' , 'ê∂ï®' ) ) "
    Case prvcSubName_Language
        l_str_Sql = l_str_Sql & " WHERE vSubjectName in ('ïßåÍ' , 'ì∆åÍ' , 'âpåÍ' ) ) "
    Case prvcSubName_SecondExam
        l_str_Sql = l_str_Sql & " WHERE vSubjectName in ('ñ ê⁄áT' , 'ñ ê⁄áU' , 'è¨ò_ï∂' ) ) "
    Case Else
        l_str_Sql = l_str_Sql & " WHERE vSubjectName='" & Trim(cboSubject.Text) & "' ) "
    End Select
    l_str_Sql = l_str_Sql & " AND tbSTEScoreProfile.iAbsentFlag=1) "
    If m_int_Action = 0 Then
        l_str_Sql = l_str_Sql & " AND iRoomProfileId=" & cboRoomId.Text & " "
    End If

    Select Case m_int_Action
    Case 0   ' 1st exam

        ' l_str_Sql = l_str_Sql & " AND iExamineeStatus = 0"
        ' modify form codesign 16/08/02
        '
        Select Case Trim(cboSubject.Text)
        Case "êîäw"
            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default
        Case "âpåÍ"
            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "')"
        Case "ì∆åÍ"
            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "')"
        Case "ïßåÍ"
            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "')"
        Case "ï®óù"
            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "') in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
        Case "âªäw"
            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "') in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
        Case "ê∂ï®"
            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "') in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
        Case prvcSubName_Science
            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & _
            " ( iScienceSubjProfileId1 in (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName in ('ï®óù' , 'âªäw' , 'ê∂ï®' ) ) " & _
            " OR iScienceSubjProfileId2 in (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName in ('ï®óù' , 'âªäw' , 'ê∂ï®' ) ) ) "
        Case prvcSubName_Language
            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & _
            " iLanguageSubjProfileId in (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName in ('ïßåÍ' , 'ì∆åÍ' , 'âpåÍ' ) ) "
        End Select
    Case 2    ' 2nd exam
        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    End Select

    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)

    If l_obj_Rst.EOF Then
 
        txtTotal.Text = lstSelected.ListCount

'        Set l_obj_Rst = Nothing
'        Exit Sub
    End If

    Do While Not l_obj_Rst.EOF
        ' form the string to be displayed in the listbox
        l_str_DisplayString = l_obj_Rst.Fields("iJukenNumber").Value & _
            " - " & l_obj_Rst.Fields("vExamineeName").Value

        If l_obj_Rst.Fields("iSex").Value = 0 Then
            l_str_DisplayString = l_str_DisplayString & " (*)"
        Else
            l_str_DisplayString = l_str_DisplayString & "    "
        End If
            
        ' check whether the examinee is allocated to any room or not
        If Trim(l_obj_Rst.Fields("iRoomProfileId").Value) <> "" Then
            
            l_str_sqlRoomName = "SELECT vRoomName FROM tbSTERoomProfile" & _
                " WHERE iRoomProfileId=" & l_obj_Rst.Fields("iRoomProfileId").Value
            l_obj_rsRoomName.Open l_str_sqlRoomName, g_obj_Conn
            
            If Not l_obj_rsRoomName.EOF Then
                l_str_DisplayString = l_str_DisplayString & " - " & l_obj_rsRoomName.Fields("vRoomName").Value
            End If
            
            l_obj_rsRoomName.Close
            Set l_obj_rsRoomName = Nothing
        End If

        lstExaminees.AddItem l_str_DisplayString
        l_obj_Rst.MoveNext
    Loop
 
    l_obj_Rst.Close
    Set l_obj_Rst = Nothing

    l_str_Sql = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName,iSex,iRoomProfileId"
    l_str_Sql = l_str_Sql & " FROM tbSTEExamineeProfile WHERE iNendo = " & g_int_CurrentNendo
    l_str_Sql = l_str_Sql & " AND exists ( SELECT 1 FROM tbSTEScoreProfile"
    l_str_Sql = l_str_Sql & " WHERE tbSTEScoreProfile.iExamineeProfileId = tbSTEExamineeProfile.iExamineeProfileId "
    l_str_Sql = l_str_Sql & " AND iSubjectProfileId in (SELECT iSubjectProfileId FROM tbSTESubjectProfile"
    Select Case cboSubject.Text
    Case prvcSubName_Science
        l_str_Sql = l_str_Sql & " WHERE vSubjectName in ('ï®óù' , 'âªäw' , 'ê∂ï®'  ) ) "
    Case prvcSubName_Language
        l_str_Sql = l_str_Sql & " WHERE vSubjectName in ('ïßåÍ' , 'ì∆åÍ' , 'âpåÍ' ) ) "
    Case prvcSubName_SecondExam
        l_str_Sql = l_str_Sql & " WHERE vSubjectName in ('ñ ê⁄áT' , 'ñ ê⁄áU' , 'è¨ò_ï∂' ) ) "
    Case Else
        l_str_Sql = l_str_Sql & " WHERE vSubjectName = '" & cboSubject.Text & "' ) "
    End Select
    l_str_Sql = l_str_Sql & " AND iAbsentFlag=1)"
    If m_int_Action = 0 Then
        l_str_Sql = l_str_Sql & " AND iRoomProfileId=" & cboRoomId.Text & " "
    End If

    Select Case m_int_Action
    Case 0  ' input absentee in the 1st exam phase
        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default
    Case 2  ' input absentee in the 2nd exam phase
        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    End Select
        
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If l_obj_Rst.EOF Then
        txtTotal.Text = lstSelected.ListCount
        Set l_obj_Rst = Nothing
        Exit Sub
    End If
    
    Do While Not l_obj_Rst.EOF
        l_str_DisplayString = l_obj_Rst.Fields("iJukenNumber").Value & _
            " - " & l_obj_Rst.Fields("vExamineeName").Value
        

        If l_obj_Rst.Fields("iSex").Value = 0 Then
            l_str_DisplayString = l_str_DisplayString & " (*)"
        Else
            l_str_DisplayString = l_str_DisplayString & "    "
        End If
                
        If Trim(l_obj_Rst.Fields("iRoomProfileId").Value) <> "" Then
            l_str_sqlRoomName = "SELECT vRoomName FROM tbSTERoomProfile" & _
                " WHERE iRoomProfileId=" & l_obj_Rst.Fields("iRoomProfileId").Value
            l_obj_rsRoomName.Open l_str_sqlRoomName, g_obj_Conn
            
            If Not l_obj_rsRoomName.EOF Then
                l_str_DisplayString = l_str_DisplayString & " - " & l_obj_rsRoomName.Fields("vRoomName").Value
            End If
            
            l_obj_rsRoomName.Close
            Set l_obj_rsRoomName = Nothing
        End If
        
        lstSelected.AddItem l_str_DisplayString
        l_obj_Rst.MoveNext
    Loop

    l_obj_Rst.Close
    Set l_obj_Rst = Nothing

    txtJuTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount

End Sub


Public Sub f_void_LoadRoom()        'populate the room names

    On Error GoTo ErrorHandler

    Dim adoRs    As New ADODB.Recordset
    Dim sSQL     As String
    
    sSQL = "SELECT iRoomProfileid,vRoomName FROM tbSTERoomProfile" & _
        " WHERE iMaxCapacity > 0 "
    
    If m_int_IntRpt = 0 Then    ' change made on 31/07/02
        sSQL = sSQL & " AND iInterviewRoomFlag = 0"
    Else
        sSQL = sSQL & " AND iInterviewRoomFlag = 1"
    End If
    
    sSQL = sSQL & " ORDER BY iRoomProfileId"

'-------------------------------------------------------------------------------
'2021.12.14 add jhi
'SELECT
'    iRoomProfileid
'   ,vRoomName
'From
'    tbSTERoomProfile
'Where
'        iMaxCapacity > 0
'    AND iInterviewRoomFlag = 1
'Order By
'    iRoomProfileid
'-------------------------------------------------------------------------------
    
    adoRs.Open sSQL, g_obj_Conn

    Do While Not adoRs.EOF
        cboRoomId.AddItem adoRs.Fields("iRoomProfileid").Value    'hidden combo to keep the id's of rooms
        cboRoom.AddItem adoRs.Fields("vRoomName").Value           'combo which displays the rooms names
        adoRs.MoveNext
    Loop
    
    If cboRoom.ListCount > 0 Then
        cboRoom.ListIndex = 0
        cboRoomId.ListIndex = 0
        lblErrorDetails.Caption = ""
    Else
        lblErrorDetails.Caption = LoadResString(2010)
        Unload Me
    End If

    adoRs.Close
    Set adoRs = Nothing

    Exit Sub

ErrorHandler:
        MsgBox Err.Description, vbInformation, "ÉGÉâÅ[" ''''LoadResString(1729)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    f_bln_DataChanged = False
    Call g_void_CloseChildForm
    Unload Me

End Sub

'*******************************************************************************
'* 1éüéÛå±ê∂ List                                                              *
'* 2022.01.15 update jhi                                                       *
'*******************************************************************************
Private Sub cmdJukenList_Click()

    On Error GoTo ErrorHandler

    Dim fso                   As Object
    Dim objText               As Object
    Dim strFile               As String
    Dim blnOpenFile           As Boolean

    Dim sJukenNo              As String
    Dim sJukenNm              As String
    Dim icnt                  As Integer

    Dim strLine               As String


    If lstExaminees.ListCount < 1 Then
        cmdJukenList.Enabled = False
        Exit Sub
    End If

    cmdJukenList.Enabled = True

    blnOpenFile = False

    'FSOÉIÉuÉWÉFÉNÉbÉgÇèâä˙âª
    strFile = App.Path & "\Report\1éüéÛå±ê∂àÍóó" & Format(Now(), "YYYYMMDDHHmmSS") & ".csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.CreateTextFile(strFile, True, False)

    blnOpenFile = True

    sJukenNo = ""    'éÛå±ê∂No
    sJukenNm = ""    'éÛå±ñº


    '---------------------------------------------------------------------------
    'HeaderÇÉtÉ@ÉCÉãÇ…èoóÕ
    '---------------------------------------------------------------------------
    strLine = "No,îNìx,éÛå±î‘çÜ,éÛå±ê∂ñº"
    objText.WriteLine (strLine)


    'ÉtÉ@ÉCÉãÇèoóÕ
    For icnt = 0 To lstExaminees.ListCount - 1
        sJukenNo = Left(lstExaminees.List(icnt), 4)

        sJukenNm = Mid(lstExaminees.List(icnt), 7, 8)
        sJukenNm = Trim(sJukenNm)
        strLine = icnt + 1 & "," & g_int_CurrentNendo & "," & sJukenNo & "," & sJukenNm
        objText.WriteLine (strLine)
    Next

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If

    ShellExecute Me.hwnd, "open", strFile, vbNullString, vbNullString, 1

    Exit Sub


ErrorHandler:

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If
    MsgBox Err.Description, vbInformation, "1éüéÛå±ê∂àÍóóï\"


End Sub

'*******************************************************************************
'* 1éü çáäié“ List                                                             *
'* 2022.01.15 add jhi                                                          *
'*******************************************************************************
Private Sub cmdGoukakuList_Click()

    On Error GoTo ErrorHandler

    Dim fso                   As Object
    Dim objText               As Object
    Dim strFile               As String
    Dim blnOpenFile           As Boolean

    Dim sJukenNo              As String
    Dim sJukenNm              As String
    Dim icnt                  As Integer

    Dim strLine               As String


    If lstSelected.ListCount < 1 Then
        cmdGoukakuList.Enabled = False
        Exit Sub
    End If

    cmdGoukakuList.Enabled = True

    blnOpenFile = False

    'FSOÉIÉuÉWÉFÉNÉbÉgÇèâä˙âª
    strFile = App.Path & "\Report\1éüçáäié“àÍóó_" & Format(Now(), "YYYYMMDDHHmmSS") & ".csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.CreateTextFile(strFile, True, False)

    blnOpenFile = True

    sJukenNo = ""    'éÛå±ê∂No
    sJukenNm = ""    'çáäié“ñº


    '---------------------------------------------------------------------------
    'ê›íËÉpÉâÉÅÅ[É^ÇÉtÉ@ÉCÉãÇ…èoóÕ
    '---------------------------------------------------------------------------
''''    strLine = "â»ñ⁄: " & cboSubject.Text & "," & ",,âÔèÍñº: " & cboRoom.Text
''''    objText.WriteLine (strLine)


    '---------------------------------------------------------------------------
    'HeaderÇÉtÉ@ÉCÉãÇ…èoóÕ
    '---------------------------------------------------------------------------
    strLine = "No,îNìx,éÛå±î‘çÜ,çáäié“ñº"
    objText.WriteLine (strLine)


    '---------------------------------------------------------------------------
    'ListBoxÇÃì‡óeÇÉtÉ@ÉCÉãÇ…èoóÕ
    '---------------------------------------------------------------------------
    For icnt = 0 To lstSelected.ListCount - 1
        sJukenNo = Left(lstSelected.List(icnt), 4)

        sJukenNm = Mid(lstSelected.List(icnt), 7, 8)
        sJukenNm = Trim(sJukenNm)
        strLine = icnt + 1 & "," & g_int_CurrentNendo & "," & sJukenNo & "," & sJukenNm
        objText.WriteLine (strLine)
    Next

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If

    ShellExecute Me.hwnd, "open", strFile, vbNullString, vbNullString, 1

    Exit Sub


ErrorHandler:

    If blnOpenFile = True Then
        blnOpenFile = False
        objText.Close
        Set objText = Nothing
        Set fso = Nothing
    End If
    MsgBox Err.Description, vbInformation, "1éüçáäié“àÍóóï\"

End Sub
