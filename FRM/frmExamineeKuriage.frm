VERSION 5.00
Begin VB.Form frmExamineeKuriage 
   Caption         =   "frmExamineeKuriage : ï‚åáé“çáäiåJè„Ç∞èàóù"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15435
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmExamineeKuriage.frx":0000
   ScaleHeight     =   9900
   ScaleWidth      =   15435
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.TextBox txtGoTotal 
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
      Left            =   11220
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   31
      Top             =   5130
      Width           =   930
   End
   Begin VB.CommandButton cmdJukenList 
      Caption         =   "ï‚åáé“ÉäÉXÉgCSVèoóÕ"
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
      Left            =   240
      TabIndex        =   30
      Top             =   8175
      Width           =   2130
   End
   Begin VB.CommandButton cmdGoukakuList 
      Caption         =   "çáäié“ÉäÉXÉgCSVèoóÕ"
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
      Left            =   12180
      TabIndex        =   29
      Top             =   4620
      Width           =   2130
   End
   Begin VB.TextBox txtHoTotal 
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
      Left            =   4125
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   28
      Top             =   7875
      Width           =   930
   End
   Begin VB.TextBox txtDestJuken 
      Height          =   300
      Left            =   11790
      TabIndex        =   27
      Text            =   "txtDestJuken"
      Top             =   630
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox cboRoomID 
      Height          =   300
      Left            =   7335
      TabIndex        =   26
      Text            =   "cboRoomID"
      Top             =   525
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboRoom 
      Height          =   300
      Left            =   6105
      TabIndex        =   25
      Text            =   "cboRoom"
      Top             =   525
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboSubject 
      Height          =   300
      Left            =   4425
      TabIndex        =   24
      Text            =   "cboSubject"
      Top             =   510
      Visible         =   0   'False
      Width           =   1605
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   1450
   End
   Begin VB.TextBox txtDestName 
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   9675
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   495
      Visible         =   0   'False
      Width           =   1450
   End
   Begin VB.TextBox txtKuriage 
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
      Left            =   8640
      MaxLength       =   2
      TabIndex        =   17
      Top             =   7875
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.ListBox lstThisTimeSelected 
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
      Height          =   2010
      Left            =   7320
      MultiSelect     =   2  'ägí£
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   5835
      Width           =   4820
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
      Left            =   11205
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   7875
      Width           =   930
   End
   Begin VB.TextBox txtSourceJuken 
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
      Height          =   405
      Left            =   1650
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1080
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "çáäié“ ämíË"
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
      Left            =   5265
      TabIndex        =   7
      Top             =   8520
      Width           =   1905
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
      Left            =   5565
      TabIndex        =   6
      Top             =   4965
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
      Left            =   5565
      TabIndex        =   5
      Top             =   4365
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
      Left            =   5565
      TabIndex        =   4
      Top             =   3765
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
      Left            =   5565
      TabIndex        =   3
      Top             =   3165
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
      Height          =   3180
      ItemData        =   "frmExamineeKuriage.frx":3AD3
      Left            =   7335
      List            =   "frmExamineeKuriage.frx":3ADA
      MultiSelect     =   2  'ägí£
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   4820
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
      Height          =   5910
      ItemData        =   "frmExamineeKuriage.frx":3AEB
      Left            =   255
      List            =   "frmExamineeKuriage.frx":3AF2
      MultiSelect     =   2  'ägí£
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1920
      Width           =   4820
   End
   Begin VB.Label lblKo 
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
      Height          =   285
      Left            =   7320
      TabIndex        =   33
      Top             =   5145
      Width           =   1080
   End
   Begin VB.Label lblHo 
      BackStyle       =   0  'ìßñæ
      Caption         =   "ï‚åáé“êî"
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
      Left            =   240
      TabIndex        =   32
      Top             =   7875
      Width           =   1200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'ìßñæ
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
      Left            =   2130
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'ìßñæ
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
      Left            =   9075
      TabIndex        =   22
      Top             =   555
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblGuidance2 
      BackStyle       =   0  'ìßñæ
      Caption         =   "åJè„Ç∞âÒêîÇïtÇØÇƒämíËÇµÇ‹Ç∑ÅB"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7290
      TabIndex        =   19
      Top             =   8865
      Width           =   2805
   End
   Begin VB.Label lblGuidance1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   """ç°âÒåJè„çáäié“""ëãÇÃéÛå±ê∂Ç ""çáäié“ÉäÉXÉg""ÇÃì¸ÇÍÅA"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7185
      TabIndex        =   18
      Top             =   8655
      Width           =   4845
   End
   Begin VB.Label Label5 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "åJè„Ç∞âÒêî"
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
      Height          =   255
      Left            =   7305
      TabIndex        =   16
      Top             =   7905
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblThisTimeSelected 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "ç°âÒåJè„çáäié“"
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
      Left            =   7320
      TabIndex        =   15
      Top             =   5550
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'ìßñæ
      Caption         =   "ï‚åáçáäié“êî"
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
      Left            =   9750
      TabIndex        =   12
      Top             =   7905
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Height          =   300
      Left            =   270
      TabIndex        =   11
      Top             =   1155
      Width           =   1335
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'ìßñæ
      Caption         =   "lblMsg"
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
      Left            =   285
      TabIndex        =   10
      Top             =   9270
      Width           =   14235
   End
   Begin VB.Label lblSelectFrom 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "ï‚åáé“ÉäÉXÉg"
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
      Left            =   255
      TabIndex        =   9
      Top             =   1635
      Width           =   4800
   End
   Begin VB.Label lblSelected 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "çáäié“ÉäÉXÉg"
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
      Left            =   7335
      TabIndex        =   8
      Top             =   1635
      Width           =   4800
   End
End
Attribute VB_Name = "frmExamineeKuriage"
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

Private f_bln_SelectAll   As Boolean          'Shows the status of the Select All button
Private f_bln_Select      As Boolean          'Shows the status of the Select  button
Private f_bln_DeSelect    As Boolean          'Shows the status of the DeSelectAll button
Private f_bln_DeSelectAll As Boolean          'Shows the status of the DeSelect  button
Public m_int_Action       As Long             'determine the action to be performed
Dim f_bln_DataChanged     As Boolean          'to enable/disable the save button
Dim f_bln_FormLoaded      As Boolean          'to check whether form is loaded or not
Public m_int_IntRpt       As Long             'form variable variable which indicated whether the form has to be instantiated for the "interview" or "report"


' The different values of m_int_action and what they stand for
'   0   -   Input Absentee Record for 1st exam
'   1   -   Input Passed Person data for 1st exam
'   2   -   Input absentee record for 2nd exam
'   3   -   Input Passed Person data for 2nd exam
'   4   -   Input waiting list for 2nd exam
'   5   -   upliftment from waiting list for Enter/Refuse phase
'   6   -   Input Refuse offer for Enter/Refuse phase

'*******************************************************************************
'* Form_Load  ï‚åáé“çáäiåJè„Ç∞èàóù                                             *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler

    LoadResStrings Me
''''Call g_void_SetFontProperties(Me)     ' set the font properties

    lblMsg.ForeColor = &HFF&              ''''ê‘

    f_bln_DataChanged = False


    m_int_Action = 5 '2021.12.29 ã≠êßìI ï‚åáé“çáäiåJè„Ç∞èàóù flagÇÉZÉbÉg



    Select Case m_int_Action
    '---------------------------------------------------------------------------
    'ï‚åáé“çáäié“åJè„Ç∞èàóù
    '2021.12.29 cyosa jhi
    '---------------------------------------------------------------------------
    Case 5
        ' upliftment from waiting list

        Me.Caption = "frmExamineeKuriage : ï‚åáé“çáäiåJè„Ç∞èàóù"      ''''LoadResString(1025)

'        lblSelectFrom.Caption = "ï‚åáé“ÉäÉXÉg"                        ''''LoadResString(2411)
'        lblSelected.Caption = "çáäié“ÉäÉXÉg"                          ''''LoadResString(2410)
'        lblTotal.Caption = "ï‚åáçáäié“êî"                             ''''LoadResString(2492)
'        Label1.Caption = "ï‚åáé“î‘çÜ"
        lblThisTimeSelected.Visible = True

        'add,xzg,2008/04/08,S--------
        Label5.Visible = True
        txtKuriage.Visible = True
        'add,xzg,2008/04/08,E--------

    '---------------------------------------------------------------------------
    'é´ëﬁ
    '2021.12.22 cyosa jhi
    '---------------------------------------------------------------------------
'    Case 6
'        ' input refuse offer
'        Me.Caption = "é´ëﬁ"                           ''''LoadResString(1026)
'        lblSelectFrom.Caption = "çáäié“ÉäÉXÉg"        ''''LoadResString(2410)
'        lblSelected.Caption = "é´ëﬁé“ÉäÉXÉg"          ''''LoadResString(2412)
'        cboSubject.Visible = False
'        lblSubject.Visible = False
'        lblTotal.Caption = "ì¸äwé´ëﬁé“êî"             ''''LoadResString(2493)
'        Label3.Visible = False
'        Label2.Visible = False
'        txtSourceName.Visible = False
'        txtDestName.Visible = False
'        Label1.Caption = "é´ëﬁé“î‘çÜ"
'        Label4.Visible = False
'        txtDestJuken.Visible = False
'        lblThisTimeSelected.Caption = "ç°âÒé´ëﬁé“"
'        lblThisTimeSelected.Visible = True
'
'        cmdOK.Caption = "é´ëﬁé“ ämíË"
    End Select

    lblMsg.Caption = ""

    lstExaminees.Font = "ÇlÇr ÉSÉVÉbÉN"
    lstSelected.Font = "ÇlÇr ÉSÉVÉbÉN"
    lstThisTimeSelected.Font = "ÇlÇr ÉSÉVÉbÉN"

    lstExaminees.FontSize = 10
    lstSelected.FontSize = 10
    lstThisTimeSelected.FontSize = 10


'    If m_int_Action = 0 Or m_int_Action = 2 Then
'        Call f_void_SelectAbsentee
'    Else
'        Call f_void_Select
'    End If


    '---------------------------------------------------------------------------
    ' 3Ç¬ÇÃListboxÇ…ÉfÅ[É^ï\é¶Ç∑ÇÈ
    '---------------------------------------------------------------------------
    Call f_void_Select



    cmdDeselect.Enabled = False
    cmdSelect.Enabled = False

    Call f_void_CheckButtonStatus

    txtHoTotal.Text = lstExaminees.ListCount    'ï‚åáé“List countêî
    txtGoTotal.Text = lstSelected.ListCount     'çáäié“List countêî
    txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount 'çáäié“List countêî+åJè„Ç∞êî

    f_bln_FormLoaded = True

    Me.Visible = True     '2022.01.16 add jhi Ç±ÇÍÇ™Ç»Ç¢Ç∆â∫ãLSetFocusä÷êîÇ™ÉGÉâÅ[Ç…Ç»ÇÈ
    txtKuriage.SetFocus   '2022.01.16 add jhi

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ÉGÉâÅ[" ''''LoadResString(1729)

End Sub

'*******************************************************************************
'*Åy>ÅzÉ{É^Éìèàóù                                                              *
'*******************************************************************************
Private Sub cmdSelect_Click()

    On Error GoTo ErrorHandler
    Dim i As Long
    
    lblMsg.Caption = ""
    f_bln_Select = True

    If lstExaminees.SelCount > 0 Then
        For i = lstExaminees.ListCount - 1 To 0 Step -1
            If lstExaminees.Selected(i) Then
                lstThisTimeSelected.AddItem lstExaminees.List(i)
                lstExaminees.RemoveItem i
            End If
        Next i
    End If

    f_void_CheckButtonStatus
    f_bln_Select = False

    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtHoTotal.Text = lstExaminees.ListCount
    txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount

    'åJè„Ç∞âÒêîÇÃì¸óÕèÄîı
    Me.Visible = True     '2022.01.16 add jhi
    txtKuriage.SetFocus   '2022.01.16 add jhi


    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "ÉGÉâÅ[" ''''LoadResString(1729)

End Sub

'*******************************************************************************
'*Åyçáäié“Å@ämíËÅzÉ{É^Éìèàóù                                                   *
'*******************************************************************************
Private Sub cmdOK_Click()
    
    On Error GoTo ErrorHandler

    Dim l_str_NonlstThisTimeSelected As String                 ' to store all the non-lstThisTimeSelected juken numbers as a string
    Dim l_str_ExamineeID             As String                 ' string of examinee id's
    Dim l_str_MySql                  As String
    Dim oRs1                         As New ADODB.Recordset
    Dim oRs2                         As New ADODB.Recordset
    Dim oRs3                         As New ADODB.Recordset
    Dim oRs4                         As New ADODB.Recordset
    Dim l_str_ExamineeIDSql          As String                 ' to store the SQL string
    Dim l_int_subjectProfileId       As Long                   ' to store the subject profile Id
    Dim l_int_NewScoreProfileId      As Long                   ' to store the score profile Id
    Dim sSQL1                        As String                 ' to store the SQL string
    Dim sSQL2                        As String

    Dim bRtn As Boolean

    Dim oRs                          As New ADODB.Recordset    ' recordset variable
    Dim i                            As Long
    Dim sSQL                         As String
    Dim sJukenNo                     As String
    Dim iTempJuken                   As Long
    Dim rinf                         As Long



    ''''2021.12.15 add jhi
    rinf = myMsgBox("çáäié“ ämíËèàóùÇé¿çsÇµÇ‹Ç∑ÅBÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH", gTit)
    If rinf = vbCancel Then
        Exit Sub
    End If

    
    'ç°âÒåJè„çáäié“ListBoxÇ≈ëIëÇµÇΩéÛå±î‘çÜÇéÊÇËèoÇµÇƒÅAcommaãÊÇ´ÇËÇ≈ï∂éöóÒÇçÏê¨Ç∑ÇÈ
    sJukenNo = ""
    For i = 0 To lstThisTimeSelected.ListCount - 1
        iTempJuken = Left(lstThisTimeSelected.List(i), 4)
        sJukenNo = sJukenNo & "," & iTempJuken
    Next

    If Len(Trim(sJukenNo)) > 0 Then
        sJukenNo = Right(Trim(sJukenNo), Len(Trim(sJukenNo)) - 1)
    End If
    
    ' get all the examinees in non-lstThisTimeSelected examinees(left) list box into a single string
    For i = 0 To lstExaminees.ListCount - 1
        iTempJuken = Left(lstExaminees.List(i), 4)
        l_str_NonlstThisTimeSelected = l_str_NonlstThisTimeSelected & "," & iTempJuken
    Next

    If Len(Trim(l_str_NonlstThisTimeSelected)) > 0 Then
        l_str_NonlstThisTimeSelected = Right(Trim(l_str_NonlstThisTimeSelected), Len(Trim(l_str_NonlstThisTimeSelected)) - 1)
    End If
    
    If lstThisTimeSelected.ListCount > 0 Or lstExaminees.ListCount > 0 Then
        
        g_obj_Conn.BeginTrans   ' start a transaction as there are multiple database table inserts/updates
        
        Select Case m_int_Action
        Case 5
            ' upliftment from waiting list
            If Len(sJukenNo) > 0 Then

                'add,xzg,2008/04/08,S---------------
                'â»ñ⁄åJè„Ç∞âÒêîÇí«â¡
                'check
                If Len(txtKuriage.Text) < 1 Then
                    g_obj_Conn.RollbackTrans
                    MsgBox "åJè„Ç∞âÒêîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                    txtKuriage.SetFocus
                    Exit Sub
                End If

                Dim strKuriage As String
                strKuriage = Trim(txtKuriage.Text)
                If Not IsNumeric(strKuriage) Then
                    g_obj_Conn.RollbackTrans
                    MsgBox "åJè„Ç∞âÒêî(1Å`99)ÇÃîÕàÕÇÇ≤éwíËÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                    txtKuriage.SetFocus
                    Exit Sub
                Else
                    If Val(strKuriage) > 100 Or Val(strKuriage) < 1 Then
                        g_obj_Conn.RollbackTrans
                        MsgBox "åJè„Ç∞âÒêî(1Å`99)ÇÃîÕàÕÇÇ≤éwíËÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                        txtKuriage.SetFocus
                        Exit Sub
                    End If
                End If
                
                'åJè„Ç∞êîÇÃí«â¡ÅiiKuriageÅj
                sSQL = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_2ndWaitPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " ,iKuriage=" & strKuriage & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & sJukenNo & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = " & gclExamineeStatus_2ndWait
                    'add,xzg,2008/04/08,E---------------

                g_obj_Conn.Execute sSQL
                 
            End If
            
            ' set the status back to 3, in case someone is moved from right to left
            If Len(l_str_NonlstThisTimeSelected) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_2ndWait & ","
                sSQL = sSQL & " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'"
                sSQL = sSQL & " WHERE iNendo = " & g_int_CurrentNendo
                sSQL = sSQL & " AND iJukenNumber IN (" & l_str_NonlstThisTimeSelected & ")"
                sSQL = sSQL & " AND iAbsentFlag = 0"
                sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_2ndWaitPass
                
                g_obj_Conn.Execute sSQL
            End If
            
        Case 6
            ' input refuse offer
            If Len(sJukenNo) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iRejectFlag = 1," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & sJukenNo & ")" '& _
'                    " AND iAbsentFlag = 0" & _
'                    " AND iExamineeStatus IN(2,6)"
                
                g_obj_Conn.Execute sSQL
            End If
            
            ' set the rejectflag back to 0, in case someone is moved from right to left
            If Len(l_str_NonlstThisTimeSelected) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iRejectFlag = 0," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_NonlstThisTimeSelected & ")" '& _
'                    " AND iAbsentFlag = 0" & _
'                    " AND iExamineeStatus IN(2,6)"
                
                g_obj_Conn.Execute sSQL
            End If
            
        End Select
        
        g_obj_Conn.CommitTrans
        
        If f_bln_DataChanged Then
            f_bln_DataChanged = False
            cmdOK.Enabled = False
        End If

        lblMsg.Caption = "çáäié“ämíËèàóùÇ™äÆóπÇµÇ‹ÇµÇΩÅB" ''''LoadResString(2404):çXêVèàóùÇÕê≥èÌÇ…äÆóπÇµÇ‹ÇµÇΩÅB

    End If


    ''''ListBox 3Ç¬Ççƒï\é¶Ç∑ÇÈ
    Call f_void_Select

    txtKuriage.Text = "" 'ê›íËåJè„Ç∞âÒêîÇclear

    txtHoTotal.Text = lstExaminees.ListCount
    txtGoTotal.Text = lstSelected.ListCount
    txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount


    Exit Sub

ErrorHandler:
    g_obj_Conn.RollbackTrans
    lblMsg.Caption = LoadResString(2405)
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler

    Dim i As Long

    
    fMainForm.mnuTools.Enabled = False

    For i = 1 To fMainForm.Toolbar1.Buttons.Count
        ' disable the toolbar buttons
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next
 
''''    If m_int_Action = 0 Or m_int_Action = 2 Then
''''        Call f_void_SelectAbsentee
''''    Else
''''        Call f_void_Select
''''    End If

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub cmdSelectAll_Click()

    'On the click of this button all the Examinees from the lstExaminees will be transfered to lstThisTimeSelectedInterviewers
    Dim l_int_Examinees As Long
    On Error GoTo ErrorHandler

    
    f_bln_SelectAll = True
    
    lblMsg.Caption = ""
    If lstExaminees.ListCount >= 1 Then
        For l_int_Examinees = lstExaminees.ListCount - 1 To 0 Step -1
            lstThisTimeSelected.AddItem lstExaminees.List(l_int_Examinees)
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

    txtHoTotal.Text = lstExaminees.ListCount
    txtGoTotal.Text = lstSelected.ListCount
    txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdDeselect_Click()

    'on the click of this button only the interviewer selected from the lstThisTimeSelected will be
    'transfered to lstExaminees
    Dim i As Long
    On Error GoTo ErrorHandler
    
    lblMsg.Caption = ""
    f_bln_DeSelect = True
        If lstThisTimeSelected.SelCount > 0 Then
            For i = lstThisTimeSelected.ListCount - 1 To 0 Step -1
                If lstThisTimeSelected.Selected(i) Then
                    lstExaminees.AddItem lstThisTimeSelected.List(i)
                    lstThisTimeSelected.RemoveItem i
                End If
            Next
        End If
    f_void_CheckButtonStatus
    f_bln_DeSelect = True
    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtHoTotal.Text = lstExaminees.ListCount
    txtGoTotal.Text = lstSelected.ListCount
    txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Private Sub cmdDeselectAll_Click()

    'on the click of this button all the Examinees from the lstThisTimeSelectedInterviewers will be moved to
    'LstAllinterviewers
    Dim l_int_InterviewerCount As Long
    On Error GoTo ErrorHandler
    
    lblMsg.Caption = ""
    f_bln_DeSelectAll = True
    If lstThisTimeSelected.ListCount >= 1 Then
       For l_int_InterviewerCount = lstThisTimeSelected.ListCount - 1 To 0 Step -1
            lstExaminees.AddItem lstThisTimeSelected.List(l_int_InterviewerCount)
            lstThisTimeSelected.RemoveItem l_int_InterviewerCount
        Next
    End If
    f_void_CheckButtonStatus
    f_bln_DeSelectAll = False
    If Not f_bln_DataChanged Then
        f_bln_DataChanged = True
        cmdOK.Enabled = True
    End If

    txtHoTotal.Text = lstExaminees.ListCount
    txtGoTotal.Text = lstSelected.ListCount
    txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Public Sub f_void_CheckButtonStatus()
    'Procedure to check the status of the buttons
    'i.e enabling and disabling the buttons based on the presense
    'and selection of data in the list boxes

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
    
    If lstThisTimeSelected.ListCount = 0 Then
        cmdDeselectall.Enabled = False
        cmdDeselect.Enabled = False
    Else
        cmdDeselectall.Enabled = True
        If lstThisTimeSelected.SelCount > 0 Then
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

Private Sub cboSubject_Click()

    Dim L_str_temp As String
    On Error GoTo ErrorHandler
    
    L_str_temp = UCase(LoadResString(2474)) & "*"
    lblMsg.Caption = ""
    If m_int_Action = 2 Then
        If UCase(cboSubject) Like L_str_temp Then
            g_int_ExamType = 2
        Else
            g_int_ExamType = 3
        End If
    End If

'    If f_bln_FormLoaded Then Call f_void_SelectAbsentee

    Exit Sub

ErrorHandler:
     MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cboRoom_Click()

    cboRoomId.ListIndex = cboRoom.ListIndex
    If f_bln_FormLoaded Then Call f_void_SelectAbsentee

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
'    f_void_CheckButtonStatus
End Sub

Private Sub lstSelected_DblClick()
'    cmdDeselect_Click
'    f_void_CheckButtonStatus
End Sub

'***************************************************************************
'* ï‚åáé“ÉäÉXÉgÅAçáäié“ÉäÉXÉgÇListBoxÇ…ï\é¶Ç∑ÇÈèàóù                       *
'***************************************************************************
Private Sub f_void_Select()

    Dim oRs         As ADODB.Recordset    ' recordset object
    Dim sSQL        As String             ' The SQL string
    Dim sTmp        As String             ' listboxÇ…ï\é¶Ç∑ÇÈï∂éöóÒ
    Dim sKuriage    As String             ' åJè„Ç∞âÒêî
  
      
    lstExaminees.Clear           'ï‚åáé“ÉäÉXÉg
    lstSelected.Clear            'çáäié“ÉäÉXÉg
    lstThisTimeSelected.Clear    'ç°âÒåJè„çáäié“


    '***************************************************************************
    '* ï‚åáé“ÉäÉXÉgÇÃÉfÅ[É^íäèo  gclExamineeStatus_2ndWait=3Ç…Ç»ÇÈ             *
    '***************************************************************************
    sSQL = "SELECT iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName" & _
           " ,iSex,iKuriage FROM tbSTEExamineeProfile WHERE" & _
           " iNendo = " & g_int_CurrentNendo & _
           " AND iAbsentFlag = 0"
    
    Select Case m_int_Action

   '----------------------------------------------------------------------------
   'ï‚åáé“(=3) ëŒè€
   '----------------------------------------------------------------------------
    Case 5
        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_2ndWait

   '----------------------------------------------------------------------------
   'é´ëﬁëŒè€(iExamineeStatus=2:2éüçáäié“ or iExamineeStatus=6:åJè„Ç∞çáäié“)
   '----------------------------------------------------------------------------
    Case 6
        sSQL = sSQL & " AND (iExamineeStatus = " & gclExamineeStatus_2ndPass & " or iExamineeStatus = " & gclExamineeStatus_2ndWaitPass & ") and iRejectFlag = 0"

    End Select

'------------------------------------------------
'2021.12.16 add jhi
'SELECT
'    iJukenNumber
'-- ,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName
'   ,vExamineeName
'   ,iSex
'   ,iKuriage
'FROM
'    tbSTEExamineeProfile
'WHERE
'        iNendo = 2020
'    AND iAbsentFlag     = 0
'    AND iExamineeStatus = 3
'------------------------------------------------
       
    Set oRs = g_obj_Conn.Execute(sSQL)
    
'    If oRs.EOF Then
'        Set oRs = Nothing
'        Exit Sub
'    End If

    Do While Not oRs.EOF
        sTmp = g_str_LPad(oRs.Fields("iJukenNumber").Value, Len(oRs.Fields("iJukenNumber").Value)) & _
            " - " & oRs.Fields("vExamineeName").Value

        ''''íjê´ÇÃÇµÇÈÇµÇïtÇØÇÈ( - (*))
        If oRs.Fields("iSex").Value = 0 Then
            sTmp = sTmp & " - (*)"    ''''fPadLeft(" - (*)", 8, " ")
        Else
            sTmp = sTmp & "      "    ''''fPadLeft("      ", 8, " ")
        End If


        If IsNull(oRs.Fields("iKuriage").Value) = True Then
            sTmp = sTmp & "  " & " 0"
        Else
            sKuriage = CStr(oRs.Fields("iKuriage").Value)
            sTmp = sTmp & "  " & PadLeft(sKuriage, 2, " ")
        End If

        lstExaminees.AddItem sTmp
        oRs.MoveNext
    Loop

    oRs.Close
    Set oRs = Nothing


    '***************************************************************************
    '* çáäié“ÉäÉXÉgÇÃÉfÅ[É^íäèo                                                *
    '***************************************************************************
    sSQL = ""
    sSQL = sSQL & "SELECT" & vbCrLf
    sSQL = sSQL & "    iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName" & vbCrLf
    sSQL = sSQL & "   ,iSex" & vbCrLf
    sSQL = sSQL & "   ,iKuriage" & vbCrLf
    sSQL = sSQL & "FROM" & vbCrLf
    sSQL = sSQL & "    tbSTEExamineeProfile" & vbCrLf
    sSQL = sSQL & "WHERE" & vbCrLf
    sSQL = sSQL & "        iNendo = " & g_int_CurrentNendo & vbCrLf
    
    '---------------------------------------------------------------------------
    ' çáäié“ÉäÉXÉgÇÃÉfÅ[É^íäèoèåè SQL
    '---------------------------------------------------------------------------
    Select Case m_int_Action
    Case 5
        sSQL = sSQL & " AND iAbsentFlag = 0" & _
            " AND iRejectFlag = 0" & _
            " AND iExamineeStatus = " & gclExamineeStatus_2ndWaitPass & _
            " order by iKuriage,iJukenNumber"

    Case 6  ' enter/refuse offer
        sSQL = sSQL & " AND iAbsentFlag = 0" & _
            " AND iRejectFlag = 1" & _
            " AND iExamineeStatus IN (" & gclExamineeStatus_2ndPass & "," & gclExamineeStatus_2ndWaitPass & ")" & _
            " order by iJukenNumber"
    End Select
        

    Set oRs = g_obj_Conn.Execute(sSQL)
    
    If oRs.EOF Then
        Set oRs = Nothing
        Exit Sub
    End If

    Do While Not oRs.EOF

        sTmp = g_str_LPad(oRs.Fields("iJukenNumber").Value, Len(oRs.Fields("iJukenNumber").Value)) & _
            " - " & oRs.Fields("vExamineeName").Value

        ''''íjê´ÇÃÇµÇÈÇµÇïtÇØÇÈ( - (*))
        If oRs.Fields("iSex").Value = 0 Then
            sTmp = sTmp & " - (*)"    ''''fPadLeft(" - (*)", 8, " ")
        Else
            sTmp = sTmp & "      "    ''''fPadLeft("      ", 8, " ")
        End If

        If IsNull(oRs.Fields("iKuriage").Value) = True Then
            sTmp = sTmp & "  "
        Else
            sKuriage = CStr(oRs.Fields("iKuriage").Value)
            sTmp = sTmp & "  " & PadLeft(sKuriage, 2, " ")
        End If

        lstSelected.AddItem sTmp

        oRs.MoveNext

    Loop
    
    oRs.Close
    Set oRs = Nothing


End Sub
Private Sub lstThisTimeSelected_Click()
    f_void_CheckButtonStatus
End Sub

Private Sub lstThisTimeSelected_DblClick()
    cmdDeselect_Click
    f_void_CheckButtonStatus
End Sub

Private Sub txtDestJuken_KeyPress(KeyAscii As Integer)
    ' move the input juken number from the non-selected listbox to the selected listbox
    Dim sSQLExaminee As String             ' to form the examinee details query string
    Dim l_obj_rsExaminee As New ADODB.Recordset ' to hold the examinee details records
    Dim sJukenNo As String                 ' to sotre the input juken number
    Dim ier1 As Long               ' to loop through the list box
    Dim ier2 As Long               ' to loop through the list box
    
    On Error GoTo ErrorHandler
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
        
    If KeyAscii = 13 Then
        
        If Trim(txtDestJuken.Text) = "" Then Exit Sub     ' exit if the textbox is empty
        
        ' enable the functionality only for the "Enter/Return key"
        sSQLExaminee = "SELECT iJukenNumber, vExamineeName FROM tbSTEExamineeProfile" & _
            " WHERE iJukenNumber=" & Trim(txtDestJuken.Text) & " AND iNendo=" & g_int_CurrentNendo
        l_obj_rsExaminee.Open sSQLExaminee, g_obj_Conn
        
            
        If l_obj_rsExaminee.EOF Then
            ' the input juken number does not exist - display an error message
            lblMsg.Caption = LoadResString(2473)
        Else
            lblMsg.Caption = ""
            ' pad the input juken number with leading "0"
            sJukenNo = g_str_LPad(Trim(txtDestJuken.Text), Len(Trim(txtDestJuken.Text)))
            
            For ier1 = 0 To lstSelected.ListCount - 1
                ' loop through the list box to check whether the juken number is present or not
                If Left(lstSelected.List(ier1), 4) = sJukenNo Then
                    ' input juken is presnet
                    
                    ' display examinee name in the neme text box
                    txtDestName.Text = l_obj_rsExaminee.Fields("vExamineeName").Value
                    
                    ' make it the selected one
                    lstSelected.Selected(ier1) = True
                    
                    ' move it to the non-selected listbox
                    lblMsg.Caption = ""
                    f_bln_DeSelect = True
                        
                    lstExaminees.AddItem lstSelected.List(ier1)
                    lstSelected.RemoveItem ier1
                                
                    f_void_CheckButtonStatus
                    f_bln_DeSelect = True
                    If Not f_bln_DataChanged Then
                        f_bln_DataChanged = True
                        cmdOK.Enabled = True
                    End If

                    txtHoTotal.Text = lstExaminees.ListCount
                    txtGoTotal.Text = lstSelected.ListCount
                    txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount
                    
                    ' loop thourh the nonselected listbox, and highlight the input juken number
                    For ier2 = 0 To lstExaminees.ListCount - 1
                        If Left(lstExaminees.List(ier2), 4) = sJukenNo Then
                            lstExaminees.Selected(ier2) = True
                        Else
                            lstExaminees.Selected(ier2) = False
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
    Dim sSQLExaminee As String             ' to form the examinee details query string
    Dim l_obj_rsExaminee As New ADODB.Recordset ' to hold the examinee details records
    Dim sJukenNo As String                 ' to sotre the input juken number
    Dim ier1 As Long               ' to loop through the list box
    Dim ier2 As Long               ' to loop through the list box
    
    On Error GoTo ErrorHandler
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
        
    If KeyAscii = 13 Then
        
        If Trim(txtSourceJuken.Text) = "" Then Exit Sub     ' exit if the textbox is empty
        
        ' enable the functionality only for the "Enter/Return key"
        sSQLExaminee = "SELECT iJukenNumber, vExamineeName FROM tbSTEExamineeProfile" & _
            " WHERE iJukenNumber=" & Trim(txtSourceJuken.Text) & " AND iNendo=" & g_int_CurrentNendo
        l_obj_rsExaminee.Open sSQLExaminee, g_obj_Conn
            
        If l_obj_rsExaminee.EOF Then
            ' the input juken number does not exist - display an error message
            lblMsg.Caption = LoadResString(2473)
        Else
            lblMsg.Caption = ""
            ' pad the input juken number with leading "0"
            sJukenNo = g_str_LPad(Trim(txtSourceJuken.Text), Len(Trim(txtSourceJuken.Text)))
            
            ' loop through the list box to check whether the juken number is present or not
            For ier1 = 0 To lstExaminees.ListCount - 1
                If Left(lstExaminees.List(ier1), 4) = sJukenNo Then
                    ' input juken is presnet
                    
                    ' display examinee name in the name text box
                    txtSourceName.Text = l_obj_rsExaminee.Fields("vExamineeName").Value
                    
                    ' make it the selected one
                    lstExaminees.Selected(ier1) = True
                    
                    ' move it to the selected listbox
                    f_bln_Select = True
                    lblMsg.Caption = ""
                    
                    lstThisTimeSelected.AddItem lstExaminees.List(ier1)
                    lstExaminees.RemoveItem ier1
                           
                    f_void_CheckButtonStatus
                    f_bln_Select = False
                    If Not f_bln_DataChanged Then
                        f_bln_DataChanged = True
                        cmdOK.Enabled = True
                    End If

                    txtHoTotal.Text = lstExaminees.ListCount
                    txtGoTotal.Text = lstSelected.ListCount
                    txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount
                    
                    ' loop thourh the selected listbox, and highlight the input juken number
                    For ier2 = 0 To lstSelected.ListCount - 1
                        If Left(lstSelected.List(ier2), 4) = sJukenNo Then
                            lstSelected.Selected(ier2) = True
                        Else
                            lstSelected.Selected(ier2) = False
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
    Dim oRs As ADODB.Recordset    ' recordset object
    Dim sSQL As String             ' The SQL string
    Dim l_str_DisplayString As String   ' to form the display string in the list box
    Dim sSQLRoomName As String
    Dim l_obj_rsRoomName As New ADODB.Recordset
    
    lstExaminees.Clear
    lstSelected.Clear
        
    sSQL = "SELECT iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName,iSex,iRoomProfileId" & _
        " FROM tbSTEExamineeProfile WHERE iNendo = " & g_int_CurrentNendo & _
        " AND iExamineeProfileId NOT IN(" & _
        " SELECT iExamineeProfileId FROM tbSTEScoreProfile" & _
        " WHERE iSubjectProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "' " & _
        " AND iAbsentFlag=1)"
        If m_int_Action = 0 Then
            sSQL = sSQL & " AND iRoomProfileId=" & cboRoomId.Text & " "
        End If

    Select Case m_int_Action
    Case 0   ' 1st exam
        ' sSQL = sSQL & " AND iExamineeStatus = 0"
        ' modify form codesign 16/08/02
        '
        Select Case Trim(cboSubject.Text)
        Case "êîäw"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default
        Case "âpåÍ"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "')"
        Case "ì∆åÍ"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "')"
        Case "ïßåÍ"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND iLanguageSubjProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "')"
        Case "ï®óù"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "') in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
        Case "âªäw"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "') in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
        Case "ê∂ï®"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "') in ( iScienceSubjProfileId1 , iScienceSubjProfileId2 ) "
        Case "óùâ»"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & _
            " ( iScienceSubjProfileId1 in (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName in ('ï®óù' , 'âªäw' , 'ê∂ï®' ) ) " & _
            " OR iScienceSubjProfileId2 in (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName in ('ï®óù' , 'âªäw' , 'ê∂ï®' ) ) ) "
        Case "äOçëåÍ"
            sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default & " AND " & _
            " iLanguageSubjProfileId in (SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
            " WHERE vSubjectName in ('ïßåÍ' , 'ì∆åÍ' , 'âpåÍ' ) ) "
        End Select
    Case 2    ' 2nd exam
        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass & " ) "
    End Select
        
    Set oRs = g_obj_Conn.Execute(sSQL)
    
    If oRs.EOF Then
        txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount
        
'        Set oRs = Nothing
'        Exit Sub
    End If

    Do While Not oRs.EOF
        ' form the string to be displayed in the listbox
        l_str_DisplayString = g_str_LPad(oRs.Fields("iJukenNumber").Value, Len(oRs.Fields("iJukenNumber").Value)) & _
            " - " & oRs.Fields("vExamineeName").Value
        
        If oRs.Fields("iSex").Value = 0 Then
            l_str_DisplayString = l_str_DisplayString & " (*)"
        End If
            
        ' check whether the examinee is allocated to any room or not
        If Trim(oRs.Fields("iRoomProfileId").Value) <> "" Then
            
            sSQLRoomName = "SELECT vRoomName FROM tbSTERoomProfile" & _
                " WHERE iRoomProfileId=" & oRs.Fields("iRoomProfileId").Value
            l_obj_rsRoomName.Open sSQLRoomName, g_obj_Conn
            
            If Not l_obj_rsRoomName.EOF Then
                l_str_DisplayString = l_str_DisplayString & " - " & l_obj_rsRoomName.Fields("vRoomName").Value
            End If
            
            l_obj_rsRoomName.Close
            Set l_obj_rsRoomName = Nothing
        End If
        
        lstExaminees.AddItem l_str_DisplayString
        oRs.MoveNext
    Loop
    Set oRs = Nothing
    
    sSQL = "SELECT iJukenNumber,substring(vExamineeName + 'Å@Å@Å@Å@Å@Å@Å@Å@',1,10) as vExamineeName,iSex,iRoomProfileId" & _
        " FROM tbSTEExamineeProfile WHERE iNendo = " & g_int_CurrentNendo & _
        " AND iExamineeProfileId IN(" & _
        " SELECT iExamineeProfileId FROM tbSTEScoreProfile" & _
        " WHERE iSubjectProfileId=(SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
        " WHERE vSubjectName='" & Trim(cboSubject.Text) & "')" & _
        " AND iAbsentFlag=1)"
    
    Select Case m_int_Action
    Case 0  ' input absentee in the 1st exam phase
        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default
    Case 2  ' input absentee in the 2nd exam phase
        sSQL = sSQL & " AND iAbsentFlag = 1"
        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
    End Select
        
    Set oRs = g_obj_Conn.Execute(sSQL)
    
    If oRs.EOF Then
        txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount
        Set oRs = Nothing
        Exit Sub
    End If
    
    Do While Not oRs.EOF
        l_str_DisplayString = g_str_LPad(oRs.Fields("iJukenNumber").Value, Len(oRs.Fields("iJukenNumber").Value)) & _
            " - " & oRs.Fields("vExamineeName").Value
        
        If oRs.Fields("iSex").Value = 0 Then
            l_str_DisplayString = l_str_DisplayString & " (*)"
        End If
        
        If Trim(oRs.Fields("iRoomProfileId").Value) <> "" Then
            sSQLRoomName = "SELECT vRoomName FROM tbSTERoomProfile" & _
                " WHERE iRoomProfileId=" & oRs.Fields("iRoomProfileId").Value
            l_obj_rsRoomName.Open sSQLRoomName, g_obj_Conn
            
            If Not l_obj_rsRoomName.EOF Then
                l_str_DisplayString = l_str_DisplayString & " - " & l_obj_rsRoomName.Fields("vRoomName").Value
            End If
            
            l_obj_rsRoomName.Close
            Set l_obj_rsRoomName = Nothing
        End If
        
        lstSelected.AddItem l_str_DisplayString
        oRs.MoveNext
    Loop
    
    Set oRs = Nothing

    txtHoTotal.Text = lstExaminees.ListCount
    txtGoTotal.Text = lstSelected.ListCount
    txtTotal.Text = lstSelected.ListCount + lstThisTimeSelected.ListCount

End Sub

'*******************************************************************************
'populate the room names
'*******************************************************************************
Public Sub f_void_LoadRoom()

    Dim l_obj_RsRoom As New ADODB.Recordset
    Dim sSQLRoom As String
    
    On Error GoTo ErrorHandler
    
    sSQLRoom = "SELECT iRoomProfileid,vRoomName FROM tbSTERoomProfile" & _
        " WHERE iMaxCapacity > 0 "
    
    If m_int_IntRpt = 0 Then    ' change made on 31/07/02
        sSQLRoom = sSQLRoom & " AND iInterviewRoomFlag = 0"
    Else
        sSQLRoom = sSQLRoom & " AND iInterviewRoomFlag = 1"
    End If
    
    sSQLRoom = sSQLRoom & " ORDER BY iRoomProfileId"
    
    l_obj_RsRoom.Open sSQLRoom, g_obj_Conn
    Do While Not l_obj_RsRoom.EOF
        cboRoomId.AddItem l_obj_RsRoom.Fields("iRoomProfileid").Value       'hidden combo to keep the id's of rooms
        cboRoom.AddItem l_obj_RsRoom.Fields("vRoomName").Value              'combo which displays the rooms names
        l_obj_RsRoom.MoveNext
    Loop
    
    If cboRoom.ListCount > 0 Then
        cboRoom.ListIndex = 0
        cboRoomId.ListIndex = 0
        lblMsg.Caption = ""
    Else
        lblMsg.Caption = LoadResString(2010)
        Unload Me
    End If
    l_obj_RsRoom.Close
    Set l_obj_RsRoom = Nothing
    Exit Sub
ErrorHandler:
        MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub


Private Sub cmdOK_Click_BK()

    
    On Error GoTo ErrorHandler

    Dim i                            As Long                   ' counter
    Dim iTempJuken                   As Long                   ' to store the juken number
    Dim sJukenNo                     As String                 ' to store all the lstThisTimeSelected juken numbers as a string
    Dim l_str_NonlstThisTimeSelected As String                 ' to store all the non-lstThisTimeSelected juken numbers as a string
    Dim l_str_ExamineeID             As String                 ' string of examinee id's
    Dim sSQL                         As String                 ' to store the SQl string
    Dim l_str_MySql                  As String
    Dim oRs                          As New ADODB.Recordset    ' recordset variable
    Dim oRs1                         As New ADODB.Recordset
    Dim oRs2                         As New ADODB.Recordset
    Dim oRs3                         As New ADODB.Recordset
    Dim oRs4                         As New ADODB.Recordset
    Dim l_str_ExamineeIDSql          As String                 ' to store the SQL string
    Dim l_int_subjectProfileId       As Long                   ' to store the subject profile Id
    Dim l_int_NewScoreProfileId      As Long                   ' to store the score profile Id
    Dim sSQL1                        As String                 ' to store the SQL string
    Dim sSQL2                        As String

    Dim bRtn                         As Boolean


    
    ' get all the examinees in lstThisTimeSelected list box into a single string
    For i = 0 To lstThisTimeSelected.ListCount - 1
        iTempJuken = Left(lstThisTimeSelected.List(i), 4)
        sJukenNo = sJukenNo & "," & iTempJuken
    Next

    If Len(Trim(sJukenNo)) > 0 Then
        sJukenNo = Right(Trim(sJukenNo), Len(Trim(sJukenNo)) - 1)
    End If
    
    ' get all the examinees in non-lstThisTimeSelected examinees(left) list box into a single string
    For i = 0 To lstExaminees.ListCount - 1
        iTempJuken = Left(lstExaminees.List(i), 4)
        l_str_NonlstThisTimeSelected = l_str_NonlstThisTimeSelected & "," & iTempJuken
    Next

    If Len(Trim(l_str_NonlstThisTimeSelected)) > 0 Then
        l_str_NonlstThisTimeSelected = Right(Trim(l_str_NonlstThisTimeSelected), Len(Trim(l_str_NonlstThisTimeSelected)) - 1)
    End If
    
    If lstThisTimeSelected.ListCount > 0 Or lstExaminees.ListCount > 0 Then
        
        g_obj_Conn.BeginTrans   ' start a transaction as there are multiple database table inserts/updates
        
        Select Case m_int_Action
        Case 0, 2
            ' input absentee record for 1st exam and 2nd exam
            
            ' get the lstThisTimeSelected subject
            oRs3.Open "SELECT iSubjectProfileId FROM tbSTESubjectProfile" & _
                " WHERE vSubjectName='" & Trim(cboSubject.Text) & "'", g_obj_Conn
            If Not oRs3.EOF Then
                l_int_subjectProfileId = oRs3("isubjectprofileid")
            End If
            Set oRs3 = Nothing
                        
            ' insert/update details of lstThisTimeSelected examinees
            For i = 0 To lstThisTimeSelected.ListCount - 1
'                sSQL2 = "SELECT iScoreProfileId FROM tbSTEScoreProfile"
'                oRs2.Open sSQL2, g_obj_Conn, adOpenStatic, adLockReadOnly
'                If Not oRs2.EOF Then
'                    oRs2.MoveLast
'                    l_int_NewScoreProfileId = oRs2("iScoreProfileId") + 1
'                Else
'                    sSQL1 = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='tbSTEScoreProfile'"
'                    oRs1.Open sSQL1, g_obj_Conn, adOpenStatic, adLockReadOnly
'                    If Not oRs1.EOF Then
'                        l_int_NewScoreProfileId = oRs1("iTableCounterIdMapping")
'                    Else
'                        l_int_NewScoreProfileId = 1
'                    End If
'                    oRs1.Close
'                    Set oRs1 = Nothing
'                End If
'                ' release the object variable
'                oRs2.Close
'                Set oRs2 = Nothing


                bRtn = getNewId("tbSTEScoreProfile", "iScoreProfileId", l_int_NewScoreProfileId)

                iTempJuken = Left(lstThisTimeSelected.List(i), 4)
                
                l_str_ExamineeIDSql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber = " & iTempJuken
                If m_int_Action = 0 Then
                    ' input absentee record for 1st exam
                    l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                Else
                    ' input absentee record for 2nd exam
                    l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                End If
                
                oRs4.Open l_str_ExamineeIDSql, g_obj_Conn, adOpenStatic, adLockReadOnly
                                       
                sSQL = "SELECT iScoreProfileId FROM tbSTEScoreProfile" & _
                    " WHERE iExamineeProfileId=" & oRs4("iExamineeProfileId") & _
                    " AND iSubjectProfileId=" & l_int_subjectProfileId & _
                    " AND iAbsentFlag = 1"
                oRs2.Open sSQL, g_obj_Conn
                If oRs2.EOF Then
                    sSQL = "INSERT INTO tbSTEScoreProfile (iScoreProfileId,iSubjectProfileId,iExamineeProfileId,iAbsentFlag,dtCreate,dtUpdate) VALUES(" & _
                        l_int_NewScoreProfileId & "," & _
                        l_int_subjectProfileId & "," & _
                        oRs4("iExamineeProfileId") & ", 1,'" & _
                        Format(Date, "MM/DD/YYYY") & "','" & Format(Date, "MM/DD/YYYY") & "')"
                End If
                oRs2.Close
                Set oRs2 = Nothing
                
                g_obj_Conn.Execute sSQL
                
                sSQL = "UPDATE tbSTEExamineeProfile SET iAbsentFlag = 1, dtUpdate='" & Format(Date, "MM/DD/YYYY") & "' WHERE" & _
                    " iNendo = " & g_int_CurrentNendo & _
                    " AND iExamineeProfileId = " & oRs4("iExamineeProfileId")
                If m_int_Action = 0 Then
                    ' input absentee record for 1st exam
                    sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default
                Else
                    ' input absentee record for 2nd exam
                    sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                End If
                
                g_obj_Conn.Execute sSQL
                
                Set oRs4 = Nothing
            Next
            
            ' insert/update details of non-lstThisTimeSelected examinees
            For i = 0 To lstExaminees.ListCount - 1
                iTempJuken = Left(lstExaminees.List(i), 4)
                
                l_str_ExamineeIDSql = "SELECT iExamineeProfileId FROM tbSTEExamineeProfile" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber = " & iTempJuken
                If m_int_Action = 0 Then
                    ' input absentee record for 1st exam
                    l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_Default
                Else
                    ' input absentee record for 2nd exam
                    l_str_ExamineeIDSql = l_str_ExamineeIDSql & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                End If
                
                oRs4.Open l_str_ExamineeIDSql, g_obj_Conn, adOpenStatic, adLockReadOnly
                                        
                sSQL = "DELETE FROM tbSTEScoreProfile WHERE iAbsentFlag = 1" & _
                    " AND iExamineeProfileId=" & oRs4("iExamineeProfileId") & _
                    " AND iSubjectProfileId=" & l_int_subjectProfileId
                  
                g_obj_Conn.Execute sSQL
                
                ' check whether the examinee is present for all other subjects
                sSQL1 = "SELECT iSubjectProfileId FROM tbSTEScoreProfile" & _
                    " WHERE iSubjectProfileId <>" & l_int_subjectProfileId & _
                    " AND iExamineeProfileId=" & oRs4("iExamineeProfileId") & _
                    " and iAbsentFlag=1"
                oRs1.Open sSQL1, g_obj_Conn
                If oRs1.EOF Then
                                    
                    sSQL = "UPDATE tbSTEExamineeProfile SET iAbsentFlag = 0," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iExamineeProfileId = " & oRs4("iExamineeProfileId")
                    
                    If m_int_Action = 0 Then
                        ' input absentee record for 1st exam
                        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_Default
                    Else
                        ' input absentee record for 2nd exam
                        sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                    End If
                    
                    g_obj_Conn.Execute sSQL
                End If
                
                oRs1.Close
                Set oRs1 = Nothing
                Set oRs4 = Nothing
            Next
            
        Case 1
            
            ' input passed person data for 1st exam
            If Len(sJukenNo) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_1stPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & sJukenNo & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = " & gclExamineeStatus_Default
                
                g_obj_Conn.Execute sSQL
            End If
            
            ' set the status back to 0, in case someone is moved from right to left
            If Len(l_str_NonlstThisTimeSelected) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_Default & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_NonlstThisTimeSelected & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                
                g_obj_Conn.Execute sSQL
            End If
        Case 3
            ' input passed person data for 2nd exam
            If Len(sJukenNo) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_2ndPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & sJukenNo & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                
                g_obj_Conn.Execute sSQL
            End If
            
            ' set the status back to 1, in case someone is moved from right to left
            If Len(l_str_NonlstThisTimeSelected) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_1stPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_NonlstThisTimeSelected & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = " & gclExamineeStatus_2ndPass
                
                g_obj_Conn.Execute sSQL
            End If
        Case 4
            ' input waiting list for 2nd exam
            If Len(sJukenNo) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_2ndWait & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & sJukenNo & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = " & gclExamineeStatus_1stPass
                
                g_obj_Conn.Execute sSQL
            End If
            
            ' set the status back to 1, in case someone is moved from right to left
            If Len(l_str_NonlstThisTimeSelected) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_1stPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_NonlstThisTimeSelected & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = " & gclExamineeStatus_2ndWait
                
                g_obj_Conn.Execute sSQL
            End If
        
        Case 5
            ' upliftment from waiting list
            If Len(sJukenNo) > 0 Then
                'add,xzg,2008/04/08,S---------------
                'â»ñ⁄åJè„Ç∞âÒêîÇí«â¡
                'check
                If Len(txtKuriage.Text) < 1 Then
                    g_obj_Conn.RollbackTrans
                    MsgBox "åJè„Ç∞âÒêîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                    txtKuriage.SetFocus
                    Exit Sub
                End If
                Dim strKuriage As String
                strKuriage = Trim(txtKuriage.Text)
                If Not IsNumeric(strKuriage) Then
                    g_obj_Conn.RollbackTrans
                    MsgBox "åJè„Ç∞âÒêîÅi1Å`ÇQÇOÅjÇê≥ÇµÇ≠ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                    txtKuriage.SetFocus
                    Exit Sub
                Else
                    If Val(strKuriage) > 20 Or Val(strKuriage) < 1 Then
                        g_obj_Conn.RollbackTrans
                        MsgBox "åJè„Ç∞âÒêîÅi1Å`ÇQÇOÅjÇê≥ÇµÇ≠ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                        txtKuriage.SetFocus
                        Exit Sub
                    End If
                End If
                
                'åJè„Ç∞êîÇÃí«â¡ÅiiKuriageÅj
                sSQL = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_2ndWaitPass & "," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " ,iKuriage=" & strKuriage & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & sJukenNo & ")" & _
                    " AND iAbsentFlag = 0" & _
                    " AND iExamineeStatus = " & gclExamineeStatus_2ndWait
                    'add,xzg,2008/04/08,E---------------
                g_obj_Conn.Execute sSQL
                 
            End If
            
            ' set the status back to 3, in case someone is moved from right to left
            If Len(l_str_NonlstThisTimeSelected) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iExamineeStatus = " & gclExamineeStatus_2ndWait & ","
                sSQL = sSQL & " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'"
                sSQL = sSQL & " WHERE iNendo = " & g_int_CurrentNendo
                sSQL = sSQL & " AND iJukenNumber IN (" & l_str_NonlstThisTimeSelected & ")"
                sSQL = sSQL & " AND iAbsentFlag = 0"
                sSQL = sSQL & " AND iExamineeStatus = " & gclExamineeStatus_2ndWaitPass
                
                g_obj_Conn.Execute sSQL
            End If
            
        Case 6
            ' input refuse offer
            If Len(sJukenNo) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iRejectFlag = 1," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & sJukenNo & ")" '& _
'                    " AND iAbsentFlag = 0" & _
'                    " AND iExamineeStatus IN(2,6)"
                
                g_obj_Conn.Execute sSQL
            End If
            
            ' set the rejectflag back to 0, in case someone is moved from right to left
            If Len(l_str_NonlstThisTimeSelected) > 0 Then
                sSQL = "UPDATE tbSTEExamineeProfile SET iRejectFlag = 0," & _
                    " dtUpdate='" & Format(Date, "MM/DD/YYYY") & "'" & _
                    " WHERE iNendo = " & g_int_CurrentNendo & _
                    " AND iJukenNumber IN (" & l_str_NonlstThisTimeSelected & ")" '& _
'                    " AND iAbsentFlag = 0" & _
'                    " AND iExamineeStatus IN(2,6)"
                
                g_obj_Conn.Execute sSQL
            End If
            
        End Select
        
        g_obj_Conn.CommitTrans
        
        If f_bln_DataChanged Then
            f_bln_DataChanged = False
            cmdOK.Enabled = False
        End If
        lblMsg.Caption = LoadResString(2404)
    End If

    Exit Sub

ErrorHandler:
    g_obj_Conn.RollbackTrans
    lblMsg.Caption = LoadResString(2405)
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub


'*******************************************************************************
'* ï‚åáé“ List                                                                 *
'* 2022.01.16 update jhi                                                       *
'*******************************************************************************
Private Sub cmdJukenList_Click()

    On Error GoTo ErrorHandler

    Dim fso                   As Object
    Dim objText               As Object
    Dim strFile               As String
    Dim blnOpenFile           As Boolean

    Dim icnt                  As Integer

    Dim sJukenNo              As String
    Dim sJukenNm              As String
    Dim sKurikaisu            As String

    Dim strLine               As String



    If lstExaminees.ListCount < 1 Then
        cmdJukenList.Enabled = False
        Exit Sub
    End If

    cmdJukenList.Enabled = True


    blnOpenFile = False

    'FSOÉIÉuÉWÉFÉNÉbÉgÇèâä˙âª
    strFile = App.Path & "\Report\ï‚åáé“àÍóó" & Format(Now(), "YYYYMMDDHHmmSS") & ".csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.CreateTextFile(strFile, True, False)

    blnOpenFile = True

    sJukenNo = ""    'éÛå±ê∂No
    sJukenNm = ""    'éÛå±ñº
    sKurikaisu = ""  'åJè„Ç∞âÒêî

    '---------------------------------------------------------------------------
    'HeaderÇÉtÉ@ÉCÉãÇ…èoóÕ
    '---------------------------------------------------------------------------
    strLine = "No,îNìx,éÛå±î‘çÜ,éÛå±ê∂ñº,åJè„Ç∞âÒêî"
    objText.WriteLine (strLine)


    'ÉtÉ@ÉCÉãÇèoóÕ
    For icnt = 0 To lstExaminees.ListCount - 1

        sJukenNo = Left(lstExaminees.List(icnt), 4)
        sJukenNm = Mid(lstExaminees.List(icnt), 7, 8)
        sJukenNm = Trim(sJukenNm)
        sKurikaisu = Trim(Mid(lstExaminees.List(icnt), 26, 2))

        strLine = icnt + 1 & "," & g_int_CurrentNendo & "," & sJukenNo & "," & sJukenNm & "," & sKurikaisu
        objText.WriteLine (strLine)

    Next icnt

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
    MsgBox Err.Description, vbInformation, "ï‚åáé“àÍóóï\"


End Sub

Private Sub cmdGoukakuList_Click()

    On Error GoTo ErrorHandler

    Dim fso                   As Object
    Dim objText               As Object
    Dim strFile               As String
    Dim blnOpenFile           As Boolean

    Dim icnt                  As Integer

    Dim sJukenNo              As String
    Dim sJukenNm              As String
    Dim sKurikaisu            As String

    Dim strLine               As String



    If lstSelected.ListCount < 1 Then
        cmdGoukakuList.Enabled = False
        Exit Sub
    End If

    cmdGoukakuList.Enabled = True


    blnOpenFile = False

    'FSOÉIÉuÉWÉFÉNÉbÉgÇèâä˙âª
    strFile = App.Path & "\Report\ï‚åáçáäié“àÍóó" & Format(Now(), "YYYYMMDDHHmmSS") & ".csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.CreateTextFile(strFile, True, False)

    blnOpenFile = True

    sJukenNo = ""      'éÛå±ê∂No
    sJukenNm = ""      'éÛå±ñº
    sKurikaisu = ""    'åJè„Ç∞âÒêî

    '---------------------------------------------------------------------------
    'HeaderÇÉtÉ@ÉCÉãÇ…èoóÕ
    '---------------------------------------------------------------------------
    strLine = "No,îNìx,éÛå±î‘çÜ,éÛå±ê∂ñº,åJè„Ç∞âÒêî"
    objText.WriteLine (strLine)


    'ÉtÉ@ÉCÉãÇèoóÕ
    For icnt = 0 To lstSelected.ListCount - 1

        sJukenNo = Left(lstSelected.List(icnt), 4)

        sJukenNm = Mid(lstSelected.List(icnt), 7, 8)
        sJukenNm = Trim(sJukenNm)
        sKurikaisu = Trim(Mid(lstSelected.List(icnt), 26, 2))

        strLine = icnt + 1 & "," & g_int_CurrentNendo & "," & sJukenNo & "," & sJukenNm & "," & sKurikaisu
        objText.WriteLine (strLine)

    Next icnt

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
    MsgBox Err.Description, vbInformation, "ï‚åáçáäié“àÍóóï\"



End Sub

