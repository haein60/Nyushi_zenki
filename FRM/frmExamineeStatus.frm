VERSION 5.00
Begin VB.Form frmExamineeStatus 
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmExamineeStatus.frx":0000
   ScaleHeight     =   9795
   ScaleWidth      =   12420
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.ComboBox cboRoom 
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
      Left            =   7245
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   22
      Top             =   1080
      Width           =   2355
   End
   Begin VB.ComboBox cboRoomID 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   9660
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   24
      Top             =   1155
      Visible         =   0   'False
      Width           =   1485
   End
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
      Left            =   11190
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   7785
      Width           =   1230
   End
   Begin VB.TextBox txtDestJuken 
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
      Height          =   400
      Left            =   8625
      TabIndex        =   2
      Top             =   1920
      Width           =   1400
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
      ItemData        =   "frmExamineeStatus.frx":3AD3
      Left            =   2205
      List            =   "frmExamineeStatus.frx":3AD5
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   0
      Top             =   1080
      Width           =   2355
   End
   Begin VB.TextBox txtSourceJuken 
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
      Left            =   1770
      TabIndex        =   1
      Top             =   1920
      Width           =   1400
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "1éü åáê»é“ ämíË"
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
      TabIndex        =   9
      Top             =   8745
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   3810
      Width           =   1215
   End
   Begin VB.ListBox lstSelected 
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
      Height          =   4935
      Left            =   7080
      MultiSelect     =   2  'ägí£
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2790
      Width           =   5370
   End
   Begin VB.ListBox lstExaminees 
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
      Height          =   4935
      Left            =   240
      MultiSelect     =   2  'ägí£
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2790
      Width           =   5370
   End
   Begin VB.TextBox txtDestName 
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   9405
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   9225
      Width           =   2355
   End
   Begin VB.TextBox txtSourceName 
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
      Left            =   2115
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   9165
      Width           =   2355
   End
   Begin VB.Label lblRoom 
      Alignment       =   1  'âEëµÇ¶
      Caption         =   "âÔèÍñº"
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
      Left            =   5280
      TabIndex        =   23
      Top             =   1140
      Width           =   1755
   End
   Begin VB.Label lblTotal 
      Caption         =   "åáê»é“êî"
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
      Left            =   7095
      TabIndex        =   20
      Top             =   7800
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
      TabIndex        =   19
      Top             =   1965
      Width           =   1365
   End
   Begin VB.Label Label1 
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
      Height          =   300
      Left            =   240
      TabIndex        =   15
      Top             =   1965
      Width           =   1365
   End
   Begin VB.Label lblErrorDetails 
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
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   8280
      Width           =   12015
   End
   Begin VB.Label lblSubject 
      Caption         =   "â»ñ⁄ëIë"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1140
      Width           =   1755
   End
   Begin VB.Label lblSelectFrom 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "éÛå±é“ÉäÉXÉg"
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
      Left            =   255
      TabIndex        =   11
      Top             =   2445
      Width           =   5355
   End
   Begin VB.Label lblSelected 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "åáê»é“ÉäÉXÉg"
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
      Left            =   7080
      TabIndex        =   10
      Top             =   2445
      Width           =   5355
   End
   Begin VB.Label Label3 
      Caption         =   "ñºèÃ"
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
      Left            =   7680
      TabIndex        =   18
      Top             =   9240
      Width           =   1665
   End
   Begin VB.Label Label2 
      Caption         =   "ñºèÃ"
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
      TabIndex        =   16
      Top             =   9150
      Width           =   1860
   End
End
Attribute VB_Name = "frmExamineeStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmExamineeStatus
'Author         :   Vishal Kamath
'Created On     :   10/10/01
'Description    :   To Uplift from waiting List OR refuse Offer
'Reference      :   FunctionalSpecs OF ExamineeStatus.doc(Ver1.1)
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History   -   04/04/2002  -   Dileep Cherian
'While updating the examinee status, check whether there are any records to be updated or not
'In case of insert/update of more than one table, it should be within a transaction
'**************************************************************************************************
'Ammemdments    -   NyushiChangesSummary.doc(ver 1.0)
'Modification History   -   29/05/2002  -   Dileep Cherian
'in absentee record screen, examinees absent for the specific selected subject should
'be displayed. If an examinee is absent for a single exam, he is considered absent for the
'entire examination
'**************************************************************************************************

Private f_bln_SelectAll   As Boolean    'Shows the status of the Select All button
Private f_bln_Select      As Boolean    'Shows the status of the Select  button
Private f_bln_DeSelect    As Boolean    'Shows the status of the DeSelectAll button
Private f_bln_DeSelectAll As Boolean    'Shows the status of the DeSelect  button
Public m_int_Action       As Long       'determine the action to be performed
Dim f_bln_DataChanged     As Boolean    'to enable/disable the save button
Dim f_bln_FormLoaded      As Boolean    'to check whether form is loaded or not
Public m_int_IntRpt       As Long       'form variable variable which indicated whether the form has to be instantiated for the "interview" or "report"

''''Private Const prvcSubName_Language As String = "äOçëåÍ" ''''2021.12.14 del jhi(äOçëåÍÇÕÇ»Ç¢)
Private Const prvcSubName_Language As String = "âpåÍ"       ''''2021.12.14 add jhi(äOçëåÍ->âpåÍÇ…ïœçX)
Private Const prvcSubName_Science As String = "óùâ»"
Private Const prvcSubName_SecondExam As String = "ÇQéüééå±"

Private Sub Form_Load()

    On Error GoTo ErrorHandler


    LoadResStrings Me
    Call g_void_SetFontProperties(Me)     ' set the font properties


    f_bln_DataChanged = False
    f_bln_FormLoaded = False
    
    lblRoom.Visible = False
    cboRoom.Visible = False

    Select Case m_int_Action

    '---------------------------------------------------------------------------
    '1éüééå±:åáê»é“List
    '---------------------------------------------------------------------------
    Case 0
        Me.Caption = "frmExamineeStatus : åáê»é“ì¸óÕ"  ''''LoadResString(1010)
        lblSelectFrom.Caption = "éÛå±ê∂ÉäÉXÉg"         ''''LoadResString(2408)
        lblSelected.Caption = "åáê»é“ÉäÉXÉg"           ''''LoadResString(2409)
        lblTotal.Caption = "åáê»é“êî"                  ''''LoadResString(2489)


        lblRoom.Visible = True
        cboRoom.Visible = True

        Label1.Visible = False         'sourceéÛå±î‘çÜlabel ''''2021.12.14 add
        txtSourceJuken.Visible = False 'sourceéÛå±î‘çÜtext  ''''2021.12.14 add

        Label3.Visible = False
        Label2.Visible = False
        Label4.Visible = False

        txtDestJuken.Visible = False
        txtSourceName.Visible = False
        txtDestName.Visible = False

        cmdOK.Caption = "1éü åáê»é“ ämíË"

        Call f_void_LoadRoom            'populate room combo

    '---------------------------------------------------------------------------
    '1éüééå±:çáäié“List
    '---------------------------------------------------------------------------
    Case 1
        ' input passed person data for 1st exam
        Me.Caption = "frmExamineeStatus : çáäié“ì¸óÕ"    ''''LoadResString(1013)
        lblSelectFrom.Caption = "éÛå±ê∂ÉäÉXÉg"           ''''LoadResString(2408)
        lblSelected.Caption = "çáäié“ÉäÉXÉg"             ''''LoadResString(2410)
        cboSubject.Visible = False
        lblSubject.Visible = False
        Label4.Visible = False
        txtDestJuken.Visible = False
        Label1.Caption = "éÛå±î‘çÜ"
        lblTotal.Caption = "çáäié“êî"                    ''''LoadResString(2490)
        cmdOK.Caption = "1éü çáäié“ ämíË"

    '---------------------------------------------------------------------------
    '2éüééå±:åáê»é“List
    '---------------------------------------------------------------------------
    Case 2
        Me.Caption = LoadResString(1010)
        lblSelectFrom.Caption = LoadResString(2408)
        lblSelected.Caption = LoadResString(2409)
        lblTotal.Caption = LoadResString(2489)
        Label3.Visible = False
        Label2.Visible = False
        Label4.Visible = False
        txtDestJuken.Visible = False
        txtSourceName.Visible = False
        txtDestName.Visible = False

        cmdOK.Caption = "2éü åáê»é“ ämíË"

    '---------------------------------------------------------------------------
    '2éüééå±:çáäié“List
    '---------------------------------------------------------------------------
    Case 3
        Me.Caption = LoadResString(1013)

        lblSelectFrom.Caption = LoadResString(2408)
        lblSelected.Caption = LoadResString(2410)
        cboSubject.Visible = False
        lblSubject.Visible = False
        lblTotal.Caption = LoadResString(2490)
        Label4.Visible = False
        txtDestJuken.Visible = False
        Label1.Caption = "çáäié“î‘çÜ"

        cmdOK.Caption = "2éü çáäié“ ämíË"

    '---------------------------------------------------------------------------
    '2éüééå±:ï‚åáé“List
    '---------------------------------------------------------------------------
    Case 4
        Me.Caption = "frmExamineeStatus : ï‚åáé“ì¸óÕ"      ''''LoadResString(1022) ï‚åáé“ì¸óÕ

        cboSubject.Visible = False
        lblSubject.Visible = False
        Label4.Visible = False
        txtDestJuken.Visible = False


        lblSelectFrom.Caption = "éÛå±ê∂ÉäÉXÉg"             ''''LoadResString(2408)
        lblSelected.Caption = "ï‚åáé“ÉäÉXÉg"               ''''LoadResString(2411)
        lblTotal.Caption = "ï‚åáé“êî"                      ''''LoadResString(2491)


        Label1.Caption = "ï‚åáé“î‘çÜ"

        cmdOK.Caption = "2éü ï‚åáé“ ämíË"

    '---------------------------------------------------------------------------
    'ï‚åáé“çáäié“åJè„Ç∞èàóù
    '---------------------------------------------------------------------------
    Case 5
        Me.Caption = "frmExamineeStatus : ï‚åáé“çáäié“åJè„Ç∞èàóù"  ''''LoadResString(1025)
        cboSubject.Visible = False
        lblSubject.Visible = False

        lblSelectFrom.Caption = "ï‚åáé“ÉäÉXÉg"                     ''''LoadResString(2411)
        lblSelected.Caption = "çáäié“ÉäÉXÉg"                       ''''LoadResString(2410)
        lblTotal.Caption = "ï‚åáçáäié“êî"                          ''''LoadResString(2492)

    Case 6
        ' input refuse offer
        Me.Caption = LoadResString(1026)
        lblSelectFrom.Caption = LoadResString(2410)
        lblSelected.Caption = LoadResString(2412)
        cboSubject.Visible = False
        lblSubject.Visible = False
        lblTotal.Caption = LoadResString(2493)
    End Select

    Label3.Visible = False
    Label2.Visible = False
    txtSourceName.Visible = False
    txtDestName.Visible = False

    lstExaminees.Font = "ÇlÇr ÉSÉVÉbÉN"
    lstSelected.Font = "ÇlÇr ÉSÉVÉbÉN"
'    lstThisTimeSelected.Font = "ÇlÇr ÉSÉVÉbÉN"

    lstExaminees.FontSize = 10
    lstSelected.FontSize = 10



    '---------------------------------------------------------------------------
    '- â»ñ⁄ÇëIëcomboÉZÉbÉg                                                   -
    '---------------------------------------------------------------------------
    ''''1éüééå±ÅA2éüééå±Åuåáê»é“ListÅvì¸óÕÇ≈ÇÕâ»ñ⁄Çï\é¶Ç∑ÇÈ
    If m_int_Action = 0 Or m_int_Action = 2 Then
        'input absentee record
        Call f_void_cboSubject_Get
        cboSubject.ListIndex = 0

    End If



    '---------------------------------------------------------------------------
    ' 1éüééå± åáê»é“ì¸óÕ, 2éüééå± åáê»é“ì¸óÕ ÉfÅ[É^ê›íË
    '---------------------------------------------------------------------------
    If m_int_Action = 0 Or m_int_Action = 2 Then
        Call f_void_SelectAbsentee
    Else
        Call f_void_Select
    End If

    cmdDeselect.Enabled = False
    cmdSelect.Enabled = False

    Call f_void_CheckButtonStatus

    txtTotal.Text = lstSelected.ListCount
    f_bln_FormLoaded = True

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ÉGÉâÅ[" ''''LoadResString(1729)

End Sub

Private Sub cboRoom_Click()

    On Error GoTo ErrorHandler
    Dim L_str_temp As String

    cboRoomId.ListIndex = cboRoom.ListIndex
'    If f_bln_FormLoaded Then Call f_void_SelectAbsentee
    
    L_str_temp = UCase(LoadResString(2474)) & "*"
    lblErrorDetails.Caption = ""

    If m_int_Action = 2 Then
        If UCase(cboSubject) Like L_str_temp Then
            g_int_ExamType = 2
        Else
            g_int_ExamType = 3
        End If
    End If

    If f_bln_FormLoaded Then Call f_void_SelectAbsentee

    Exit Sub

ErrorHandler:
     MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub f_void_cboSubject_Get()

    On Error GoTo ErrorHandler
    Dim sSQL          As String                 ' SQL string
    Dim adoRs         As New ADODB.Recordset    ' recordset object
    Dim l_int_Counter As Long
 
   
    ' select all subjects that come under the selected exam type
    sSQL = "SELECT iSubjectprofileID,vSubjectName FROM tbSTESubjectProfile "

    If m_int_Action = 0 Then
        sSQL = sSQL & "WHERE iExamType =" & g_int_ExamType
    ElseIf m_int_Action = 2 Or m_int_Action = 3 Then
        sSQL = sSQL & "WHERE iExamType IN(2,3,4,5)"
    End If
    sSQL = sSQL & " ORDER BY iDispOrder"

'-------------------------------------------------------------------------------
'2021.12.14 add jhi
'SELECT
'    --iSubjectprofileID
'   --,vSubjectName
'   *
'From
'    tbSTESubjectProfile
'Where
'    iExamType = 1
'ORDER BY iDispOrder
'-------------------------------------------------------------------------------
    
    Set adoRs = g_obj_Conn.Execute(sSQL)
    
    ' add the subjects to combo box
    Do While Not adoRs.EOF
        l_int_Counter = l_int_Counter + 1
        cboSubject.AddItem adoRs("vSubjectName")
        adoRs.MoveNext
    Loop
    
    ' release the object variables
    adoRs.Close
    Set adoRs = Nothing

    '1éüééå±ÇÃåáê»é“List
''''2021.12.28 del jhi Ç»Ç∫Ç±ÇÍÇí«â¡Ç∑ÇÈÇÃÇ©ÅH
'    If m_int_Action = 0 Then
'        cboSubject.AddItem prvcSubName_Science, 0  'óùâ»
'        cboSubject.AddItem prvcSubName_Language, 0 'âpåÍ
'    End If

    '2éüééå±ÇÃåáê»é“List
    If m_int_Action = 2 Then
        cboSubject.AddItem prvcSubName_SecondExam, 0 '2éüééå±
    End If

'    If l_int_Counter > 0 Then
'        cboSubject.ListIndex = 0
'    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ÉGÉâÅ[" ''''LoadResString(1729)
End Sub


' The different values of m_int_action and what they stand for
'   0   -   Input Absentee Record for 1st exam
'   1   -   Input Passed Person data for 1st exam
'   2   -   Input absentee record for 2nd exam
'   3   -   Input Passed Person data for 2nd exam
'   4   -   Input waiting list for 2nd exam
'   5   -   upliftment from waiting list for Enter/Refuse phase
'   6   -   Input Refuse offer for Enter/Refuse phase

Private Sub cboSubject_Click()

    On Error GoTo ErrorHandler
    Dim L_str_temp As String
    
    L_str_temp = UCase(LoadResString(2474)) & "*"
    lblErrorDetails.Caption = ""

    If m_int_Action = 2 Then
        If UCase(cboSubject) Like L_str_temp Then
            g_int_ExamType = 2
        Else
            g_int_ExamType = 3
        End If
    End If
    If f_bln_FormLoaded Then Call f_void_SelectAbsentee

    Exit Sub

ErrorHandler:
     MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub cmdOK_Click()

    Dim l_int_Count As Long                  ' counter
    Dim l_int_TempJuken As Long              ' to store the juken number
    Dim l_str_JukenNo As String                 ' to store all the selected juken numbers as a string
    Dim l_str_NonSelected As String             ' to store all the non-selected juken numbers as a string
    Dim l_str_ExamineeID As String              ' string of examinee id's
    Dim l_obj_Rec As ADODB.Recordset            ' recordset variable
    Dim l_str_Sql As String                     ' to store the SQl string
    Dim l_str_MySql As String
    Dim l_obj_Rst As New ADODB.Recordset        ' recordset variable
    Dim l_obj_rst1 As New ADODB.Recordset
    Dim l_obj_rst2 As New ADODB.Recordset
    Dim l_obj_rst3 As New ADODB.Recordset
    Dim l_obj_rst4 As New ADODB.Recordset
    Dim l_str_ExamineeIDSql As String           ' to store the SQL string
    Dim l_int_subjectProfileId As Long       ' to store the subject profile Id
    Dim l_int_NewScoreProfileId As Long      ' to store the score profile Id
    Dim l_str_Sql1 As String                    ' to store the SQL string
    Dim l_str_sql2 As String

    Dim bRtn As Boolean
    
    On Error GoTo ErrorHandler
    
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
        Case 0
            ' input absentee record for 1st exam

            ' get the selected subject
            If Trim(cboSubject.Text) = prvcSubName_Language Or Trim(cboSubject.Text) = prvcSubName_Science Then
                l_obj_rst3.Open "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
                    " WHERE iSubType = " & IIf(Trim(cboSubject.Text) = prvcSubName_Language, "1", "2"), g_obj_Conn
            ElseIf Trim(cboSubject.Text) = prvcSubName_SecondExam Then
                l_obj_rst3.Open "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
                    " WHERE iExamType = 2 ", g_obj_Conn
            Else
                l_obj_rst3.Open "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
                    " WHERE vSubjectName='" & Trim(cboSubject.Text) & "'", g_obj_Conn
            End If

            Do Until l_obj_rst3.EOF

                l_int_subjectProfileId = l_obj_rst3("isubjectprofileid")

'                l_int_TempJuken = Left(lstSelected.List(l_int_Count), 4)

                ' insert/update details of selected examinees
                For l_int_Count = 0 To lstSelected.ListCount - 1

                    l_int_TempJuken = Left(lstSelected.List(l_int_Count), 4)

'ì¸ééééå±é¿é{éûïsãÔçáNo6ëŒâû  2005/01/22 êîäwÇÃÇ∆Ç´ÅAåáê»Ç…Ç»ÇÁÇ»Ç¢ïsãÔçáèCê≥ÅB
'                    l_str_sql2 = " SELECT COUNT(*) FROM tbSTEExamineeProfile as ep where " & l_int_subjectProfileId & IIf(Trim(cboSubject.Text) = prvcSubName_Language, " = iLanguageSubjProfileId ", " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) ")
                    l_str_sql2 = " SELECT COUNT(*) FROM tbSTEExamineeProfile as ep inner join tbSTEExamineeRoomProfile as erp on erp.iExamineeProfileId = ep.iExamineeProfileId "
                    l_str_sql2 = l_str_sql2 & " where erp.iSubjectProfileId = " & l_int_subjectProfileId
                    l_str_sql2 = l_str_sql2 & " AND ep.iNendo = " & g_int_CurrentNendo
                    l_str_sql2 = l_str_sql2 & " AND ep.iJukenNumber = " & l_int_TempJuken
                    l_obj_rst2.Open l_str_sql2, g_obj_Conn, adOpenStatic, adLockReadOnly
                    If l_obj_rst2.Fields(0) = 0 Then
                        l_obj_rst2.Close
                        GoTo LoopEnd
                    Else
                        l_obj_rst2.Close
                    End If


'                    l_str_sql2 = "SELECT max(iScoreProfileId) as iScoreProfileId FROM tbSTEScoreProfile"
'                    l_obj_rst2.Open l_str_sql2, g_obj_Conn, adOpenStatic, adLockReadOnly
'                    If Not l_obj_rst2.EOF Then
'                        l_obj_rst2.MoveLast
'                        l_int_NewScoreProfileId = l_obj_rst2("iScoreProfileId") + 1
'                    Else
'                        l_str_Sql1 = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='tbSTEScoreProfile'"
'                        l_obj_rst1.Open l_str_Sql1, g_obj_Conn, adOpenStatic, adLockReadOnly
'                        If Not l_obj_rst1.EOF Then
'                            l_int_NewScoreProfileId = l_obj_rst1("iTableCounterIdMapping")
'                        Else
'                            l_int_NewScoreProfileId = 1
'                        End If
'                        l_obj_rst1.Close
'                        Set l_obj_rst1 = Nothing
'                    End If
'                    ' release the object variable
'                    l_obj_rst2.Close
'                    Set l_obj_rst2 = Nothing
                    bRtn = getNewId("tbSTEScoreProfile", "iScoreProfileId", l_int_NewScoreProfileId)

                    l_str_ExamineeIDSql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                        " WHERE iNendo = " & g_int_CurrentNendo & _
                        " AND iJukenNumber = " & l_int_TempJuken
                    If m_int_Action = 0 Then
                        ' input absentee record for 1st exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                    Else
                        ' input absentee record for 2nd exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                    End If
                    
                    l_obj_rst4.Open l_str_ExamineeIDSql, g_obj_Conn, adOpenStatic, adLockReadOnly
                                           
                    l_str_Sql = "SELECT iScoreProfileId FROM tbSTEScoreProfile" & _
                        " WHERE iExamineeProfileId=" & l_obj_rst4("iExamineeProfileId") & _
                        " AND iSubjectProfileId=" & l_int_subjectProfileId & _
                        " AND iAbsentFlag = 1"
                    l_obj_rst2.Open l_str_Sql, g_obj_Conn
                    If l_obj_rst2.EOF Then
                        l_str_Sql = "INSERT INTO tbSTEScoreProfile (iScoreProfileId,iSubjectProfileId,iExamineeProfileId,iAbsentFlag,dtCreate,dtUpdate) VALUES(" & _
                            l_int_NewScoreProfileId & "," & _
                            l_int_subjectProfileId & "," & _
                            l_obj_rst4("iExamineeProfileId") & ", 1,'" & _
                            Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
                    End If
                    l_obj_rst2.Close
                    Set l_obj_rst2 = Nothing
                    
                    g_obj_Conn.Execute l_str_Sql
                    
                    l_str_Sql = "UPDATE tbSTEExamineeProfile SET iAbsentFlag = 1, dtUpdate='" & Format(Date, "MM/DD/YYYY") & "' WHERE" & _
                        " iNendo = " & g_int_CurrentNendo & _
                        " AND iExamineeProfileId = " & l_obj_rst4("iExamineeProfileId")
                    If m_int_Action = 0 Then
                        ' input absentee record for 1st exam
                        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                    Else
                        ' input absentee record for 2nd exam
                        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                    End If
                    
                    g_obj_Conn.Execute l_str_Sql
                    
                    Set l_obj_rst4 = Nothing

LoopEnd:

                Next
                
                ' insert/update details of non-selected examinees
                For l_int_Count = 0 To lstExaminees.ListCount - 1
                    l_int_TempJuken = Left(lstExaminees.List(l_int_Count), 4)

'ì¸ééééå±é¿é{éûïsãÔçáNo6ëŒâû  2005/01/22 êîäwÇÃÇ∆Ç´ÅAåáê»Ç…Ç»ÇÁÇ»Ç¢ïsãÔçáèCê≥ÅB
'                    l_str_sql2 = " SELECT COUNT(*) FROM tbSTEExamineeProfile as ep where " & l_int_subjectProfileId & IIf(Trim(cboSubject.Text) = prvcSubName_Language, " = iLanguageSubjProfileId ", " in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) ")
                    l_str_sql2 = " SELECT COUNT(*) FROM tbSTEExamineeProfile as ep inner join tbSTEExamineeRoomProfile as erp on erp.iExamineeProfileId = ep.iExamineeProfileId "
                    l_str_sql2 = l_str_sql2 & " where erp.iSubjectProfileId = " & l_int_subjectProfileId
                    l_str_sql2 = l_str_sql2 & " AND iNendo = " & g_int_CurrentNendo
                    l_str_sql2 = l_str_sql2 & " AND iJukenNumber = " & l_int_TempJuken
                    l_obj_rst2.Open l_str_sql2, g_obj_Conn, adOpenStatic, adLockReadOnly
                    If l_obj_rst2.Fields(0) = 0 Then
                        l_obj_rst2.Close
                        GoTo LoopEnd2
                    Else
                        l_obj_rst2.Close
                    End If

                    l_str_ExamineeIDSql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                        " WHERE iNendo = " & g_int_CurrentNendo & _
                        " AND iJukenNumber = " & l_int_TempJuken
                    If m_int_Action = 0 Then
                        ' input absentee record for 1st exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                    Else
                        ' input absentee record for 2nd exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                    End If
                    
                    l_obj_rst4.Open l_str_ExamineeIDSql, g_obj_Conn, adOpenStatic, adLockReadOnly
                                            
                    l_str_Sql = "DELETE FROM tbSTEScoreProfile WHERE iAbsentFlag = 1" & _
                        " AND iExamineeProfileId=" & l_obj_rst4("iExamineeProfileId") & _
                        " AND iSubjectProfileId=" & l_int_subjectProfileId
                      
                    g_obj_Conn.Execute l_str_Sql
                    
                    ' check whether the examinee is present for all other subjects
                    l_str_Sql1 = "SELECT iSubjectProfileId FROM tbSTEScoreProfile" & _
                        " WHERE iSubjectProfileId <>" & l_int_subjectProfileId & _
                        " AND iExamineeProfileId=" & l_obj_rst4("iExamineeProfileId") & _
                        " and iAbsentFlag=1"
                    l_obj_rst1.Open l_str_Sql1, g_obj_Conn
                    If l_obj_rst1.EOF Then
                                        
                        l_str_Sql = "UPDATE tbSTEExamineeProfile SET iAbsentFlag = 0," & _
                        " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                        " WHERE iNendo = " & g_int_CurrentNendo & _
                        " AND iExamineeProfileId = " & l_obj_rst4("iExamineeProfileId")
                        
                        If m_int_Action = 0 Then
                            ' input absentee record for 1st exam
                            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                        Else
                            ' input absentee record for 2nd exam
                            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                        End If
                        
                        g_obj_Conn.Execute l_str_Sql
                    End If
                    
                    l_obj_rst1.Close
                    Set l_obj_rst1 = Nothing
                    Set l_obj_rst4 = Nothing

LoopEnd2:

                Next

                l_obj_rst3.MoveNext

            Loop

            Set l_obj_rst3 = Nothing

        Case 2
            ' input absentee record for 2nd exam
            
            ' get the selected subject
            If Trim(cboSubject.Text) = prvcSubName_SecondExam Then
                l_obj_rst3.Open "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
                    " WHERE iExamType = 2 ", g_obj_Conn
            Else
                l_obj_rst3.Open "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
                    " WHERE vSubjectName='" & Trim(cboSubject.Text) & "'", g_obj_Conn
            End If

            Do Until l_obj_rst3.EOF

                l_int_subjectProfileId = l_obj_rst3("isubjectprofileid")

                ' insert/update details of selected examinees
                For l_int_Count = 0 To lstSelected.ListCount - 1
'                    l_str_sql2 = "SELECT ISNULL( MAX(iScoreProfileId) , -1 ) FROM tbSTEScoreProfile"
'                    l_obj_rst2.Open l_str_sql2, g_obj_Conn, adOpenStatic, adLockReadOnly
'                    If l_obj_rst2.Fields(0) > -1 Then
'                        l_int_NewScoreProfileId = l_obj_rst2.Fields(0) + 1
'                    Else
'                        l_str_Sql1 = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='tbSTEScoreProfile'"
'                        l_obj_rst1.Open l_str_Sql1, g_obj_Conn, adOpenStatic, adLockReadOnly
'                        If Not l_obj_rst1.EOF Then
'                            l_int_NewScoreProfileId = l_obj_rst1("iTableCounterIdMapping")
'                        Else
'                            l_int_NewScoreProfileId = 1
'                        End If
'                        l_obj_rst1.Close
'                        Set l_obj_rst1 = Nothing
'                    End If
'                    ' release the object variable
'                    l_obj_rst2.Close
'                    Set l_obj_rst2 = Nothing
                    bRtn = getNewId("tbSTEScoreProfile", "iScoreProfileId", l_int_NewScoreProfileId)

                    l_int_TempJuken = Left(lstSelected.List(l_int_Count), 4)

                    l_str_ExamineeIDSql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                        " WHERE iNendo = " & g_int_CurrentNendo & _
                        " AND iJukenNumber = " & l_int_TempJuken
                    If m_int_Action = 0 Then
                        ' input absentee record for 1st exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                    Else
                        ' input absentee record for 2nd exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                    End If

                    l_obj_rst4.Open l_str_ExamineeIDSql, g_obj_Conn, adOpenStatic, adLockReadOnly

                    l_str_Sql = "SELECT iScoreProfileId FROM tbSTEScoreProfile" & _
                        " WHERE iExamineeProfileId=" & l_obj_rst4("iExamineeProfileId") & _
                        " AND iSubjectProfileId=" & l_int_subjectProfileId & _
                        " AND iAbsentFlag = 1"
                    l_obj_rst2.Open l_str_Sql, g_obj_Conn
                    If l_obj_rst2.EOF Then
                        l_str_Sql = "INSERT INTO tbSTEScoreProfile (iScoreProfileId,iSubjectProfileId,iExamineeProfileId,iAbsentFlag,dtCreate,dtUpdate) VALUES(" & _
                            l_int_NewScoreProfileId & "," & _
                            l_int_subjectProfileId & "," & _
                            l_obj_rst4("iExamineeProfileId") & ", 1,'" & _
                            Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
                    End If
                    l_obj_rst2.Close
                    Set l_obj_rst2 = Nothing
                    
                    g_obj_Conn.Execute l_str_Sql
                    
                    l_str_Sql = "UPDATE tbSTEExamineeProfile SET iAbsentFlag = 1, dtUpdate='" & Format(Date, "MM/DD/YYYY") & "' WHERE" & _
                        " iNendo = " & g_int_CurrentNendo & _
                        " AND iExamineeProfileId = " & l_obj_rst4("iExamineeProfileId")
                    If m_int_Action = 0 Then
                        ' input absentee record for 1st exam
                        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                    Else
                        ' input absentee record for 2nd exam
                        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                    End If
                    
                    g_obj_Conn.Execute l_str_Sql
                    
                    Set l_obj_rst4 = Nothing
                Next

                ' insert/update details of non-selected examinees
                For l_int_Count = 0 To lstExaminees.ListCount - 1
                    l_int_TempJuken = Left(lstExaminees.List(l_int_Count), 4)
                    
                    l_str_ExamineeIDSql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                        " WHERE iNendo = " & g_int_CurrentNendo & _
                        " AND iJukenNumber = " & l_int_TempJuken
                    If m_int_Action = 0 Then
                        ' input absentee record for 1st exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                    Else
                        ' input absentee record for 2nd exam
                        l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                    End If
                    
                    l_obj_rst4.Open l_str_ExamineeIDSql, g_obj_Conn, adOpenStatic, adLockReadOnly
                                            
                    l_str_Sql = "DELETE FROM tbSTEScoreProfile WHERE iAbsentFlag = 1" & _
                        " AND iExamineeProfileId=" & l_obj_rst4("iExamineeProfileId") & _
                        " AND iSubjectProfileId=" & l_int_subjectProfileId
                      
                    g_obj_Conn.Execute l_str_Sql
                    
                    ' check whether the examinee is present for all other subjects
                    l_str_Sql1 = "SELECT iSubjectProfileId FROM tbSTEScoreProfile" & _
                        " WHERE iSubjectProfileId <>" & l_int_subjectProfileId & _
                        " AND iExamineeProfileId=" & l_obj_rst4("iExamineeProfileId") & _
                        " and iAbsentFlag=1"
                    l_obj_rst1.Open l_str_Sql1, g_obj_Conn
                    If l_obj_rst1.EOF Then
                                        
                        l_str_Sql = "UPDATE tbSTEExamineeProfile SET iAbsentFlag = 0," & _
                        " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                        " WHERE iNendo = " & g_int_CurrentNendo & _
                        " AND iExamineeProfileId = " & l_obj_rst4("iExamineeProfileId")
                        
                        If m_int_Action = 0 Then
                            ' input absentee record for 1st exam
                            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                        Else
                            ' input absentee record for 2nd exam
                            l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                        End If
                        
                        g_obj_Conn.Execute l_str_Sql
                    End If
                    
                    l_obj_rst1.Close
                    Set l_obj_rst1 = Nothing
                    Set l_obj_rst4 = Nothing
                Next

                l_obj_rst3.MoveNext

            Loop

            l_obj_rst3.Close
            Set l_obj_rst3 = Nothing

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
        Case 3
            ' input passed person data for 2nd exam
            If Len(l_str_JukenNo) > 0 Then
                l_str_Sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_2ndPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_JukenNo & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = 1"
                
                g_obj_Conn.Execute l_str_Sql
            End If
            
            ' set the status back to 1, in case someone is moved from right to left
            If Len(l_str_NonSelected) > 0 Then
                l_str_Sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_1stPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_NonSelected & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = 2"
                
                g_obj_Conn.Execute l_str_Sql
            End If
        Case 4
            ' input waiting list for 2nd exam
            If Len(l_str_JukenNo) > 0 Then
                l_str_Sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_2ndWait & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_JukenNo & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = 1"
                
                g_obj_Conn.Execute l_str_Sql
            End If
            
            ' set the status back to 1, in case someone is moved from right to left
            If Len(l_str_NonSelected) > 0 Then
                l_str_Sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_1stPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_NonSelected & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = 3"
                
                g_obj_Conn.Execute l_str_Sql
            End If
        
        Case 5
            ' upliftment from waiting list
            If Len(l_str_JukenNo) > 0 Then
                l_str_Sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_2ndWaitPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_JukenNo & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = 3"
                
                g_obj_Conn.Execute l_str_Sql
            End If
            
            ' set the status back to 3, in case someone is moved from right to left
            If Len(l_str_NonSelected) > 0 Then
                l_str_Sql = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_2ndWait & ","
                l_str_Sql = l_str_Sql & " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'"
                l_str_Sql = l_str_Sql & " WHERE iNendo = " & g_int_CurrentNendo
                l_str_Sql = l_str_Sql & " AND iJukenNumber IN (" & l_str_NonSelected & ")"
                l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0"
                l_str_Sql = l_str_Sql & " AND iExamineeStatus = 6"
                
                g_obj_Conn.Execute l_str_Sql
            End If
            
        Case 6
            ' input refuse offer
            If Len(l_str_JukenNo) > 0 Then
                l_str_Sql = "UPDATE tbSTEExamineeProfile SET iRejectFlag = 1," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_JukenNo & ")" '& _
'                    " AND iAbsentFlag = 0" & _
'                    " AND iExamineeStatus IN(2,6)"
                
                g_obj_Conn.Execute l_str_Sql
            End If
            
            ' set the rejectflag back to 0, in case someone is moved from right to left
            If Len(l_str_NonSelected) > 0 Then
                l_str_Sql = "UPDATE tbSTEExamineeProfile SET iRejectFlag = 0," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_NonSelected & ")" '& _
'                    " AND iAbsentFlag = 0" & _
'                    " AND iExamineeStatus IN(2,6)"

                g_obj_Conn.Execute l_str_Sql
            End If
            
        End Select
        
        g_obj_Conn.CommitTrans
        
        If f_bln_DataChanged Then
            f_bln_DataChanged = False
            cmdOK.Enabled = False
        End If
        lblErrorDetails.Caption = LoadResString(2404)
    End If

    Exit Sub

ErrorHandler:
    g_obj_Conn.RollbackTrans
    lblErrorDetails.Caption = LoadResString(2405)
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler

    Dim Index As Long
    
    fMainForm.mnuTools.Enabled = False
    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
        ' disable the toolbar buttons
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next

'    If m_int_Action = 0 Or m_int_Action = 2 Then
'        Call f_void_SelectAbsentee
'    Else
'        Call f_void_Select
'    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub


Private Sub cmdSelectAll_Click()
    'On the click of this button all the Examinees from the lstExaminees will be transfered to lstSelectedInterviewers
    Dim l_int_Examinees As Long
    On Error GoTo ErrorHandler
    
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
    txtTotal.Text = lstSelected.ListCount
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdSelect_Click()
    'on the click of this button only the Examinee selected from the lstExaminees will be transfered to
    'lstSelected
    Dim l_int_Count As Long
    On Error GoTo ErrorHandler
    
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
    txtTotal.Text = lstSelected.ListCount
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdDeselect_Click()
    'on the click of this button only the interviewer selected from the lstSelected will be
    'transfered to lstExaminees
    Dim l_int_Count As Long
    On Error GoTo ErrorHandler
    
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
    txtTotal.Text = lstSelected.ListCount
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdDeselectAll_Click()
    'on the click of this button all the Examinees from the lstSelectedInterviewers will be moved to
    'LstAllinterviewers
    Dim l_int_InterviewerCount As Long
    On Error GoTo ErrorHandler
    
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
    'Procedure to check the status of the buttons
    'i.e enabling and disabling the buttons based on the presense
    'and selection of data in the list boxes

    If lstExaminees.ListCount = 0 Then
        cmdSelectAll.Enabled = False
        cmdSelect.Enabled = False
    Else
        cmdSelectAll.Enabled = True
        If lstExaminees.SelCount > 0 Then
            cmdSelect.Enabled = True
        Else
            cmdSelect.Enabled = False
        End If
    End If
    
    If lstSelected.ListCount = 0 Then
        cmdDeselectAll.Enabled = False
        cmdDeselect.Enabled = False
    Else
        cmdDeselectAll.Enabled = True
        If lstSelected.SelCount > 0 Then
            cmdDeselect.Enabled = True
        Else
            cmdDeselect.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    f_bln_DataChanged = False
    Call g_void_CloseChildForm
    Unload Me
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

Private Sub f_void_Select()

    Dim l_obj_Rst As ADODB.Recordset    ' recordset object
    Dim l_str_Sql As String             ' The SQL string
    Dim l_str_DisplayString As String   ' to form the display string in the list box
        
    lstExaminees.Clear
    lstSelected.Clear
        
    l_str_Sql = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName,iSex FROM tbSTEExamineeProfile WHERE" & _
        " iNendo = " & g_int_CurrentNendo & _
        " AND iAbsentFlag = 0"
    
    Select Case m_int_Action
    Case 1   ' 1st exam
        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_Default
    Case 3, 4    ' 2nd exam
        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    Case 5  ' enter/refuse phase
        l_str_Sql = l_str_Sql & " AND iExamineeStatus = " & gclExamineeStatus_2ndWait
    Case 6  ' enter/refuse phase
        l_str_Sql = l_str_Sql & " AND (iExamineeStatus = " & gclExamineeStatus_2ndPass & " or iExamineeStatus = " & gclExamineeStatus_2ndWaitPass & ") and iRejectFlag = 0"
    End Select
        
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
'    If l_obj_Rst.EOF Then
'        Set l_obj_Rst = Nothing
'        Exit Sub
'    End If
    Do While Not l_obj_Rst.EOF
        l_str_DisplayString = l_obj_Rst.Fields("iJukenNumber").Value & _
            " - " & l_obj_Rst.Fields("vExamineeName").Value
        If l_obj_Rst.Fields("iSex").Value = 0 Then
            l_str_DisplayString = l_str_DisplayString & " - (*)"
        Else
            l_str_DisplayString = l_str_DisplayString & "      "
        End If

        lstExaminees.AddItem l_str_DisplayString
        l_obj_Rst.MoveNext
    Loop
    Set l_obj_Rst = Nothing
    
    l_str_Sql = "SELECT dbo.usfMakeDispJukenNumber(iJukenNumber) as iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName,iSex FROM tbSTEExamineeProfile  WHERE"
    l_str_Sql = l_str_Sql & " iNendo = " & g_int_CurrentNendo
    
    Select Case m_int_Action
    Case 1  ' input passed person data
        l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    Case 3  ' passed person data for 2nd phase
        l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_2ndPass
    Case 4  ' waiting list
        l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_2ndWait
    Case 5  ' upliftment from waiting list
        l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_2ndWaitPass
    Case 6  ' enter/refuse offer
        l_str_Sql = l_str_Sql & " AND iAbsentFlag = 0" & _
            " AND iRejectFlag = 1" & _
            " AND iExamineeStatus IN (" & gclExamineeStatus_2ndPass & "," & gclExamineeStatus_2ndWaitPass & ")"
    End Select
        
    Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
    
    If l_obj_Rst.EOF Then
        Set l_obj_Rst = Nothing
        Exit Sub
    End If
    Do While Not l_obj_Rst.EOF
        l_str_DisplayString = l_obj_Rst.Fields("iJukenNumber").Value & " - " & l_obj_Rst.Fields("vExamineeName").Value
        If l_obj_Rst.Fields("iSex").Value = 0 Then
            l_str_DisplayString = l_str_DisplayString & " - (*)"
        Else
            l_str_DisplayString = l_str_DisplayString & "      "
        End If
        lstSelected.AddItem l_str_DisplayString
        l_obj_Rst.MoveNext
    Loop
    
    Set l_obj_Rst = Nothing
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

