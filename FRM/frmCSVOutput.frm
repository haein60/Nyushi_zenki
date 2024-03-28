VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCSVOutput 
   ClientHeight    =   10095
   ClientLeft      =   165
   ClientTop       =   -1620
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmCSVOutput.frx":0000
   ScaleHeight     =   10095
   ScaleWidth      =   12300
   WindowState     =   2  '最大化
   Begin VB.CheckBox chkHeader 
      Caption         =   "Check1"
      Height          =   225
      Left            =   9000
      TabIndex        =   17
      Top             =   1155
      Value           =   1  'ﾁｪｯｸ
      Width           =   195
   End
   Begin VB.CommandButton cmdShowDefset 
      Caption         =   "パターン指定 表示"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3615
      MaskColor       =   &H80000004&
      TabIndex        =   16
      Top             =   9120
      Width           =   2325
   End
   Begin VB.TextBox txtDefSet 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   9165
      Width           =   3315
   End
   Begin VB.CommandButton cmdSaveDefset 
      Caption         =   "パターン名 保存"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9930
      TabIndex        =   14
      Top             =   9165
      Width           =   2310
   End
   Begin VB.ComboBox cboDefSet 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   255
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   13
      Top             =   9165
      Width           =   3345
   End
   Begin VB.TextBox txtNendo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      Left            =   2500
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   11
      Tag             =   "[iZipCodeId]"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtFileName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2500
      TabIndex        =   8
      Top             =   1065
      Width           =   4950
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   6780
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   4215
      Width           =   1215
   End
   Begin VB.TextBox txtSerial 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   7920
      TabIndex        =   4
      Top             =   3645
      Width           =   1215
   End
   Begin VB.ComboBox cboTarget 
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   570
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7485
      TabIndex        =   1
      Top             =   1050
      Width           =   1215
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfOutputCol 
      Height          =   6580
      Left            =   240
      TabIndex        =   0
      Top             =   2265
      Width           =   7215
      _cx             =   12726
      _cy             =   11606
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   2
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
      Editable        =   1
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
   Begin VSFlex7LCtl.VSFlexGrid vsfColList 
      Height          =   6580
      Left            =   9720
      TabIndex        =   3
      Top             =   2280
      Width           =   3300
      _cx             =   5821
      _cy             =   11606
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
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
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   4650
      TabIndex        =   10
      Top             =   1560
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label lblTit1 
      BackStyle       =   0  '透明
      Caption         =   "[ 出力パターンの項目設定 ]"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   255
      TabIndex        =   22
      Top             =   2040
      Width           =   3315
   End
   Begin VB.Label lbl02 
      BackStyle       =   0  '透明
      Caption         =   "出力パターン名を指定して保存してください。"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   6615
      TabIndex        =   21
      Top             =   8925
      Width           =   4725
   End
   Begin VB.Label lbl01 
      BackStyle       =   0  '透明
      Caption         =   "登録されている出力パターン名を指定して表示してください。"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   255
      TabIndex        =   20
      Top             =   8925
      Width           =   5805
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "左表の行を選択後"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   7470
      TabIndex        =   19
      Top             =   6495
      Width           =   2115
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列名出力"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   9270
      TabIndex        =   18
      Top             =   1140
      Width           =   960
   End
   Begin VB.Label lblNendo 
      Alignment       =   1  '右揃え
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "年度"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1470
      TabIndex        =   12
      Top             =   1620
      Width           =   795
   End
   Begin VB.Label lblFileName 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "出力ファイル名"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   390
      TabIndex        =   9
      Top             =   1140
      Width           =   1905
   End
   Begin VB.Label lblSerial 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "No.指定挿入"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   7815
      TabIndex        =   5
      Top             =   3405
      Width           =   1380
   End
End
Attribute VB_Name = "frmCSVOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type uParam_Type
    sParaName   As String
    sValue      As String
End Type

Private prvlRowNoCol As Long
Private prvlIdCol As Long
Private prvlNameCol As Long
Private prvlDBFlieldCol As Long
Private prvlOutputCol As Long
Private prvlSortCol As Long
Private prvlSortNoCol As Long
Private prvlWhereTypeCol As Long
Private prvlWhereCol As Long

Private prvlCL_IdCol As Long
Private prvlCL_NameCol As Long
Private prvlCL_OutputCol As Long

Private prvsWhereType() As String

Private Const prvclNull As Integer = 0
Private Const prvclEq As Integer = 1
Private Const prvclDai As Integer = 2
Private Const prvclDaiEq As Integer = 3
Private Const prvclSho As Integer = 4
Private Const prvclShoEq As Integer = 5
Private Const prvclDaiSho As Integer = 6
Private Const prvclFromTo As Integer = 7
Private Const prvclOr As Integer = 8

Private Const prvcsOutputOn As String = "○"
Private Const prvcsOutputOff As String = ""

Private Const prvcsSortASC As String = "昇順"
Private Const prvcsSortDESC As String = "降順"
Private Const prvcsSortNULL As String = ""

Private Type uSortType_Type
    iID         As Integer
    sSortDisp   As String
    sSortSet    As String
End Type

Private prvuSortType_() As uSortType_Type

'*******************************************************************************
'* Form_Load                                                                   *
'*******************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Dim sSQL As String
    Dim oRs  As ADODB.Recordset


''''LoadResStrings Me                                '2021.12.23 del jhi
''''Call g_void_SetFontProperties(Me)                'set the font properties 2021.12.23 del jhi

    Me.Caption = "frmCSVOutput : 受験生＋素点情報"   'LoadResString(1096) '受験生情報

    lblTit1.FontSize = 11
    lblTit1.ForeColor = &HFF0000

    lbl01.FontSize = 10
    lbl01.ForeColor = &HFF0000

    lbl02.FontSize = 10
    lbl02.ForeColor = &HFF0000

    cmdShowDefset.FontSize = 11
    cmdSaveDefset.FontSize = 11


''''2021.12.03 update jhi 西暦対応
''''txtNendo.Text = Format(DateValue(Trim(str(Trim(g_int_CurrentNendo)) & "/01/01")), "gggee年")
    txtNendo.Text = Format(DateValue(Trim(str(Trim(g_int_CurrentNendo)) & "/01/01")), "yyyy年")

    txtFileName.Text = "C:\" & Format(Now, "YYYYMMDD") & ".csv"

    Call lsSetWhereType
    Call lsSetSortType

    sSQL = "select iID , vName from tbSTECSVOutputObject "
    sSQL = sSQL & " where siViewFlag = 1 "
    sSQL = sSQL & " order by iDispOrder "
    Set oRs = g_obj_Conn.Execute(sSQL)

    Do Until oRs.EOF
        cboTarget.AddItem oRs.Fields("vName")
        cboTarget.ItemData(cboTarget.NewIndex) = oRs.Fields("iID")
        oRs.MoveNext
    Loop

    oRs.Close
    Set oRs = Nothing

    Call lsInitGrid_ColList
    Call lsInitGrid_OutputCol

    cboTarget.ListIndex = 0

    Call lsSetDefSet

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー" ''''LoadResString(1729)

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim i As Long


    fMainForm.mnuTools.Enabled = False  ' disable tools menu
    For i = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next i

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, "エラー"    ''''LoadResString(1729)

End Sub


Private Sub lsShowDefGroup()

    On Error GoTo ErrProc

    Dim sSQL As String
    Dim oRs As ADODB.Recordset

    Dim lRow As Long
    Dim lLoopCnt As Long


    If cboDefSet.ListIndex = -1 Then Exit Sub

    sSQL = "SELECT "
    sSQL = sSQL & "  dg.iCSVOutputColId "
    sSQL = sSQL & ", cc.vColumnName "
    sSQL = sSQL & ", cc.vColumnValue "
    sSQL = sSQL & ", dg.siOutputFlag "
    sSQL = sSQL & ", dg.siDefSort "
    sSQL = sSQL & ", isnull ( convert( varchar , dg.iDefSortNo ) , '' ) as iDefSortNo "
    sSQL = sSQL & ", dg.siDefWhereType "
    sSQL = sSQL & ", dg.vWhere "
    sSQL = sSQL & " from tbSTECSVOutputDefGroupCols as dg "
    sSQL = sSQL & " inner join tbSTECSVOutputCols as cc on cc.iCSVOutputColID = dg.iCSVOutputColID and cc.siViewFlag = 1"
    sSQL = sSQL & " where dg.iCSVOutputDefGroupID = " & Trim(str(cboDefSet.ItemData(cboDefSet.ListIndex)))
    sSQL = sSQL & " and cc.siUserLevel <= " & Trim(str(glUserLevel))
    sSQL = sSQL & " order by dg.iDispOrder "

    Set oRs = g_obj_Conn.Execute(sSQL)

    lRow = 1
    vsfOutputCol.Rows = 1

    Do Until oRs.EOF
        vsfOutputCol.Rows = lRow + 1
        vsfOutputCol.TextMatrix(lRow, prvlRowNoCol) = Trim(str(lRow))
        vsfOutputCol.TextMatrix(lRow, prvlIdCol) = oRs.Fields("iCSVOutputColId")
        vsfOutputCol.TextMatrix(lRow, prvlNameCol) = oRs.Fields("vColumnName")
        vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) = oRs.Fields("vColumnValue")
        vsfOutputCol.TextMatrix(lRow, prvlOutputCol) = IIf(oRs.Fields("siOutputFlag") = 1, prvcsOutputOn, prvcsOutputOff)

        For lLoopCnt = LBound(prvuSortType_) To UBound(prvuSortType_)
            If prvuSortType_(lLoopCnt).iID = oRs.Fields("siDefSort") Then
                vsfOutputCol.TextMatrix(lRow, prvlSortCol) = prvuSortType_(lLoopCnt).sSortDisp
                Exit For
            End If
        Next

        vsfOutputCol.TextMatrix(lRow, prvlSortNoCol) = oRs.Fields("iDefSortNo")
        If LBound(prvsWhereType) <= oRs.Fields("siDefWhereType") And UBound(prvsWhereType) >= oRs.Fields("siDefWhereType") Then
            vsfOutputCol.TextMatrix(lRow, prvlWhereTypeCol) = prvsWhereType(oRs.Fields("siDefWhereType"))
        End If

        vsfOutputCol.TextMatrix(lRow, prvlWhereCol) = oRs.Fields("vWhere")
        oRs.MoveNext
        lRow = lRow + 1
    Loop

    oRs.Close
    Set oRs = Nothing

ErrProc:

On Error Resume Next
    oRs.Close
    Set oRs = Nothing

End Sub

Private Sub lsSetSortType()

    ReDim prvuSortType_(2)

    prvuSortType_(0).iID = 0
    prvuSortType_(0).sSortDisp = ""
    prvuSortType_(0).sSortSet = ""
    prvuSortType_(1).iID = 1
    prvuSortType_(1).sSortDisp = prvcsSortASC
    prvuSortType_(1).sSortSet = " ASC "
    prvuSortType_(2).iID = 2
    prvuSortType_(2).sSortDisp = prvcsSortDESC
    prvuSortType_(2).sSortSet = " DESC "

End Sub

Private Sub lsSetWhereType()

    ReDim prvsWhereType(prvclOr)

    prvsWhereType(prvclNull) = ""
    prvsWhereType(prvclEq) = "="
    prvsWhereType(prvclDai) = ">"
    prvsWhereType(prvclDaiEq) = ">="
    prvsWhereType(prvclSho) = "<"
    prvsWhereType(prvclShoEq) = "<="
    prvsWhereType(prvclDaiSho) = "<>"
    prvsWhereType(prvclFromTo) = "～"
    prvsWhereType(prvclOr) = "OR"

End Sub

Private Function lsGetWhereType(psText As String)

    Dim iLoopCnt As Integer

    lsGetWhereType = LBound(prvsWhereType) - 1

    For iLoopCnt = LBound(prvsWhereType) To UBound(prvsWhereType)
        If prvsWhereType(iLoopCnt) = psText Then
            lsGetWhereType = iLoopCnt
            Exit For
        End If
    Next

End Function

Private Sub lsGetColData()

    Dim sSQL  As String
    Dim oRs   As ADODB.Recordset
    Dim sWk   As String

    On Error GoTo ErrProc

    vsfColList.Rows = 1

    sSQL = "select iCSVOutputColId , vColumnName , vColumnValue from tbSTECSVOutputCols "
    sSQL = sSQL & " where iID = " & cboTarget.ItemData(cboTarget.ListIndex)
    sSQL = sSQL & " and siViewFlag = 1"
    sSQL = sSQL & " and siUserLevel <= " & Trim(str(glUserLevel))
    sSQL = sSQL & " order by iDispOrder "
    Set oRs = g_obj_Conn.Execute(sSQL)

    Do Until oRs.EOF
        sWk = oRs.Fields("iCSVOutputColId") & vbTab & oRs.Fields("vColumnName") & vbTab & oRs.Fields("vColumnValue")
        vsfColList.AddItem sWk
        oRs.MoveNext
    Loop

    oRs.Close
    Set oRs = Nothing

    Exit Sub

ErrProc:

End Sub

'*******************************************************************************
'* 左側の項目を表示する Grid -- No. 列名 出力 並順 並順No 式 条件              *
'*******************************************************************************
Private Sub lsInitGrid_OutputCol()

    With vsfOutputCol

        .Redraw = flexRDNone

        .cols = 9
        .Rows = 1

        .FixedCols = 3
        .FixedRows = 1

        .Row = 0

        .Col = 0                       '0
        prvlRowNoCol = .Col
        .Text = "   No"
        .ColWidth(.Col) = 650

        .Col = .Col + 1                '1 xxx
        prvlIdCol = .Col
        .Text = "ID"
        .ColWidth(.Col) = 0            ''''for debug:300

        .Col = .Col + 1                '2
        prvlNameCol = .Col
        .Text = "　　　　　  　列　名"
        .ColWidth(.Col) = 2300

        .Col = .Col + 1                '3 xxx
        prvlDBFlieldCol = .Col
        .Text = "出力時指定"
        .ColWidth(.Col) = 0            '''''for debug:600

        .Col = .Col + 1                '4
        prvlOutputCol = .Col
        .Text = "　出力"
        .ColWidth(.Col) = 600
        .ColAlignment(.Col) = flexAlignCenterBottom

        .Col = .Col + 1                '5
        prvlSortCol = .Col
        .Text = "　並順"
        .ColWidth(.Col) = 800
        .ColAlignment(.Col) = flexAlignCenterBottom

        .Col = .Col + 1                '6
        prvlSortNoCol = .Col
        .Text = "　並順No"
        .ColWidth(.Col) = 800
        .ColAlignment(.Col) = flexAlignCenterBottom

        .Col = .Col + 1                '7
        prvlWhereTypeCol = .Col
        .Text = "式"
        .ColWidth(.Col) = 400
        .ColAlignment(.Col) = flexAlignCenterBottom

        .Col = .Col + 1                '8
        prvlWhereCol = .Col
        .Text = "　　　　条件"
        .ColWidth(.Col) = 1600
        .ColAlignment(.Col) = flexAlignLeftBottom

        .Redraw = flexRDDirect

    End With

End Sub

'*******************************************************************************
'* 右側の項目を表示する Grid -- 列名                                           *
'*******************************************************************************
Private Sub lsInitGrid_ColList()

    With vsfColList
        .Redraw = flexRDNone

        .cols = 3
        .Rows = 1
        .FixedCols = 0
        .FixedRows = 1

        .Row = 0

        .Col = 0
        prvlCL_IdCol = .Col
        .Text = "No"
        .ColWidth(.Col) = 0

        .Col = 1
        prvlCL_NameCol = .Col
        .Text = "        　　　　   列　　 名"
        .ColWidth(.Col) = 3000

        .Col = 2 'xx
        prvlCL_OutputCol = .Col
        .Text = "出力時指定"
        .ColWidth(.Col) = 0        ''''for debug=6000

        .Redraw = flexRDDirect
    End With

End Sub

Private Sub cboDefSet_Click()
'    Call lsShowDefGroup
End Sub

'*******************************************************************************
'* 非表示のcombo --- 受験生情報 1件が表示される                                *
'*******************************************************************************
Private Sub cboTarget_Click()

    Call lsGetColData

End Sub

Private Sub cmdDeselect_Click()

    Dim lRow As Long
    Dim sWk  As String

    If vsfOutputCol.Row <= 0 Then
        MsgBox "項目が選択されてません。", vbOKOnly, "選択"
        Exit Sub
    End If

    lRow = vsfOutputCol.Row
    If vbNo = MsgBox("№：" & Trim(str(lRow)) & "をリストから削除します。よろしいですか？", vbYesNo, "削除確認") Then
        Exit Sub
    End If

    vsfOutputCol.RemoveItem lRow

    If lRow < vsfOutputCol.Rows Then
        Call lsResetRowNumber(lRow)
    End If

End Sub

'*******************************************************************************
'* uspに付ける parameterを 作成する                                            *
'*******************************************************************************
Private Sub lsSetParam(psStr, puParam_() As uParam_Type)

    Dim i    As Integer
    Dim iPos As Integer


    For i = LBound(puParam_) To UBound(puParam_)
        iPos = InStr(1, psStr, "%" & puParam_(i).sParaName & "%")
        If iPos > 0 Then
            psStr = Left(psStr, iPos - 1) & puParam_(i).sValue & Mid(psStr, iPos + Len("%" & puParam_(i).sParaName & "%"))
        End If
    Next

End Sub

'*******************************************************************************
'* CSVファイル出力　main処理                                                   *
'*-----------------------------------------------------------------------------*
'* 2021.12.24 jhi 並順 不具合修正完成                                          *
'*******************************************************************************
Private Sub lsCSVFileOutput()

    On Error GoTo ErrProc

    Dim oRs            As ADODB.Recordset
    Dim lRow           As Long

    Dim sTarget        As String
    Dim sConstractor   As String
    Dim sDeconstractor As String

    Dim sWk            As String
    Dim iFileNo        As Integer

    Dim sWhere         As String
    Dim sWhere2        As String

    Dim iLoopCnt       As Integer
    Dim sSort          As String
    Dim sHeader        As String

    Dim iErrPos        As Integer
    Dim sErrMsg        As String

    Dim oFld()         As ADODB.Field
    Dim sDim()         As String

    Dim uParam_()      As uParam_Type

    Dim lSortRows()    As Long        ''''sort情報格納 ReDim Array
    Dim lSortNo        As Long

    Dim sFileName      As String
    Dim sNendo         As String
    Dim sSQL           As String


    sHeader = ""

    '***************************************************************************
    '* 指定項目の整合性check
    '***************************************************************************
    sFileName = Trim(txtFileName.Text)
    If sFileName = "" Then
        MsgBox "出力ファイル名が指定されていません。", vbOKOnly, "入力エラー"
        Exit Sub
    End If

    sNendo = Trim(txtNendo.Text)
    If sNendo = "" Then
        MsgBox "年度が指定されていません。", vbOKOnly, "入力エラー"
        Exit Sub
    End If


    If (vsfOutputCol.Rows - 1) = 0 Then
        MsgBox "出力パターン項目が選択されていません。", vbOKOnly, "入力エラー"
        Exit Sub
    End If


    cmdOutput.Enabled = False



    sSQL = "SELECT "
    sSQL = sSQL & " vTargetObject , isnull( vConstractor , '' ) , isnull( vDeconstractor , '' ) "
    sSQL = sSQL & " FROM tbSTECSVOutputObject "
    sSQL = sSQL & " WHERE iID=" & Trim(str(cboTarget.ItemData(cboTarget.ListIndex)))

'2021.12.24 cyosa
'SELECT
'--    *,
'    vTargetObject                  --tbSTEExamineeCSVData
'   ,isnull( vConstractor , '' )    --exec uspSTEMakeExamineeCSVData %iNendo%
'   ,isnull( vDeconstractor , '' )  --null
'From
'    tbSTECSVOutputObject
'Where
'    1=1 --iID=1

    Set oRs = g_obj_Conn.Execute(sSQL)

    sTarget = oRs.Fields(0)          'tbSTEExamineeCSVData
    sConstractor = oRs.Fields(1)     'exec uspSTEMakeExamineeCSVData %iNendo%
    sDeconstractor = oRs.Fields(2)   'null
    oRs.Close
    Set oRs = Nothing

    '-----------------------------------------------------------------------
    ' 指定年度のcsv data を tbSTEExamineeCSVData Tableを作成
    ' exec uspSTEMakeExamineeCSVData %iNendo%
    '-----------------------------------------------------------------------
    If sConstractor <> "" Then

        ReDim uParam_(0)

        uParam_(0).sParaName = "iNendo"
        uParam_(0).sValue = Year(CDate(txtNendo.Text & "1月1日"))
 
        '-----------------------------------------------------------------------
        ' "exec uspSTEMakeExamineeCSVData 2020"
        ' を実行して tbSTEExamineeCSVData Tableに csvデータ出力で使用するデータを作成する
        '-----------------------------------------------------------------------
        sSQL = sConstractor
        Call lsSetParam(sSQL, uParam_)    'exec uspSTEMakeExamineeCSVData %iNendo%

        g_obj_Conn.CommandTimeout = 360 '2022.02.03 add jhi 6分
        g_obj_Conn.Execute sSQL

    End If


    sSQL = "SELECT "

'出力パターン項目で、出力に○が1個があったら

    For lRow = 1 To vsfOutputCol.Rows - 1  'vsfOutputCol.Rows=追加した項目の行数
        If vsfOutputCol.TextMatrix(lRow, prvlOutputCol) = prvcsOutputOn Then
            sSQL = sSQL & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " "
            sHeader = vsfOutputCol.TextMatrix(lRow, prvlNameCol)
            Exit For
        End If
    Next

    For lRow = lRow + 1 To vsfOutputCol.Rows - 1
        If vsfOutputCol.TextMatrix(lRow, prvlOutputCol) = prvcsOutputOn Then
            sSQL = sSQL & "," & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " "
            sHeader = sHeader & "," & vsfOutputCol.TextMatrix(lRow, prvlNameCol)
        End If
    Next

    '---------------------------------------------------------------------------
    ' Select ... from tbSTEExamineeCSVDataまで作成
    '---------------------------------------------------------------------------
    sSQL = sSQL & " FROM " & sTarget

    '---------------------------------------------------------------------------
    ' Select ... from tbSTEExamineeCSVDataまで作成
    '---------------------------------------------------------------------------
    sWhere = ""
    For lRow = 1 To vsfOutputCol.Rows - 1
        If vsfOutputCol.TextMatrix(lRow, prvlWhereTypeCol) <> "" Then '式colに何が入っていると
        ''''If vsfOutputCol.TextMatrix(lRow, 8) <> "" Then
            If sWhere = "" Then
                sWhere2 = " WHERE "
            Else
                sWhere2 = " AND "
            End If

            Select Case lsGetWhereType(Trim(vsfOutputCol.TextMatrix(lRow, prvlWhereTypeCol)))
            Case prvclEq
                sWhere2 = sWhere2 & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " = '" & vsfOutputCol.TextMatrix(lRow, prvlWhereCol) & "'"
            Case prvclDai
                sWhere2 = sWhere2 & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " > '" & vsfOutputCol.TextMatrix(lRow, prvlWhereCol) & "'"
            Case prvclDaiEq
                sWhere2 = sWhere2 & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " >= '" & vsfOutputCol.TextMatrix(lRow, prvlWhereCol) & "'"
            Case prvclSho
                sWhere2 = sWhere2 & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " < '" & vsfOutputCol.TextMatrix(lRow, prvlWhereCol) & "'"
            Case prvclShoEq
                sWhere2 = sWhere2 & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " <= '" & vsfOutputCol.TextMatrix(lRow, prvlWhereCol) & "'"
            Case prvclDaiSho
                sWhere2 = sWhere2 & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " <> '" & vsfOutputCol.TextMatrix(lRow, prvlWhereCol) & "'"
            Case prvclFromTo
                sDim = Split(vsfOutputCol.TextMatrix(lRow, prvlWhereCol), ",")
                If LBound(sDim) = 0 And UBound(sDim) >= 1 Then
                    sWhere2 = sWhere2 & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " BETWEEN '" & sDim(0) & "' AND '" & sDim(1) & "' "
                Else
                    sWhere2 = ""
                End If
            Case prvclOr
                sDim = Split(vsfOutputCol.TextMatrix(lRow, prvlWhereCol), ",")
                sWhere2 = sWhere2 & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " IN ( '" & sDim(0) & "' "
                For iLoopCnt = LBound(sDim) + 1 To UBound(sDim)
                    sWhere2 = sWhere2 & " , '" & sDim(iLoopCnt) & "' "
                Next
                sWhere2 = sWhere2 & " ) "
            End Select
            If sWhere2 <> "" Then
                sWhere = sWhere & sWhere2
            End If
        End If
    Next


    '---------------------------------------------------------------------------
    ' sort flag処理
    '---------------------------------------------------------------------------
    sSort = ""
    lSortNo = 0
''''これににするとReDim Preserve lSortRows(1, lSortNo)でインデックスが有効範囲にありません
''''多次元配列の要素数をReDim Preserveで変更(既存データを保持)する場合は最終次元しか変更できません。
''''ReDimのみ(既存データを破棄)は可能です。

''''ReDim lSortRows(0, 0)


 ''''2021.12.24 add これがないのでエラーななってしまった。--->インデックスが有効範囲にありません
    ReDim lSortRows(1, 0)

    For lRow = 1 To vsfOutputCol.Rows - 1
        
        If vsfOutputCol.TextMatrix(lRow, prvlSortCol) <> "" And vsfOutputCol.TextMatrix(lRow, prvlSortNoCol) <> "" Then

            ReDim Preserve lSortRows(1, lSortNo)

            lSortRows(0, lSortNo) = lRow
            lSortRows(1, lSortNo) = vsfOutputCol.TextMatrix(lRow, prvlSortNoCol)
            lSortNo = lSortNo + 1

        End If
    Next


reSort:
    For lRow = LBound(lSortRows, 2) To UBound(lSortRows, 2) - 1
        If lSortRows(1, lRow) > lSortRows(1, lRow + 1) Then
            lSortNo = lSortRows(0, lRow)
            lSortRows(0, lRow) = lSortRows(0, lRow + 1)
            lSortRows(0, lRow + 1) = lSortNo
            lSortNo = lSortRows(1, lRow)
            lSortRows(1, lRow) = lSortRows(1, lRow + 1)
            lSortRows(1, lRow + 1) = lSortNo
            GoTo reSort
        End If
    Next

    For lSortNo = LBound(lSortRows, 2) To UBound(lSortRows, 2)
        lRow = lSortRows(0, lSortNo)
        Select Case vsfOutputCol.TextMatrix(lRow, prvlSortCol)
        Case prvcsSortASC
            If sSort = "" Then
                sSort = " order by " & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " ASC "
            Else
                sSort = sSort & " , " & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " ASC "
            End If
        Case prvcsSortDESC
            If sSort = "" Then
                sSort = " order by " & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " DESC "
            Else
                sSort = sSort & " , " & vsfOutputCol.TextMatrix(lRow, prvlDBFlieldCol) & " DESC "
            End If
        End Select
    Next

    Set oRs = g_obj_Conn.Execute(sSQL & sWhere & sSort)

iErrPos = 1

    If Not oRs.EOF Then

        iFileNo = FreeFile

        Open sFileName For Output As #iFileNo

iErrPos = 2

        If chkHeader.Value Then
'列名をヘッダとして出力
            Print #iFileNo, sHeader
        End If

        ReDim oFld(oRs.Fields.Count - 1)

        For lRow = 0 To UBound(oFld)
            Set oFld(lRow) = oRs.Fields(lRow)
        Next

        Do Until oRs.EOF
            sWk = gfNullChkStr(oFld(0))
            For lRow = 1 To UBound(oFld)
                sWk = sWk & "," & gfNullChkStr(oFld(lRow))
            Next

            Print #iFileNo, sWk

            oRs.MoveNext
        Loop

        For lRow = 0 To UBound(oFld)
            Set oFld(lRow) = Nothing
        Next

        Close #iFileNo

iErrPos = 1

    End If

    oRs.Close
    Set oRs = Nothing

iErrPos = 0


    MsgBox "指定項目のcsvデータが正常に作成されました。指定ファイルをご確認ください。", vbOKOnly, "確認"

    cmdOutput.Enabled = True
    Exit Sub

ErrProc:

    sErrMsg = Trim(str(Err.Number)) & ":" & Err.Description

On Error GoTo ErrProc2

    If iErrPos = 2 Then
        Close #iFileNo
    End If

ErrProc2:

On Error GoTo 0
On Error GoTo ErrProc3

    If iErrPos = 1 Then
        oRs.Close
    End If

ErrProc3:
    Set oRs = Nothing
    MsgBox sErrMsg, vbOKOnly, "出力エラー"

End Sub

Private Sub cmdOutput_Click()

    cmdOutput.Enabled = False

    Call lsCSVFileOutput

    cmdOutput.Enabled = True

End Sub

Private Sub cmdSaveDefset_Click()

    cmdSaveDefset.Enabled = False
    Call lsSaveDefset
    cmdSaveDefset.Enabled = True

End Sub

Private Sub lsSaveDefset()

    On Error GoTo ErrProc

    Dim oRs As ADODB.Recordset

    Dim sName    As String
    Dim lRow     As Long
    Dim lID      As Long
    Dim i        As Long

    Dim sSQL     As String


    sName = Trim(txtDefSet.Text)
    If sName <> "" Then
        If vbOK = MsgBox("指定内容を「" & sName & "」として新規保存します。" & vbCrLf & "よろしいですか？", vbOKCancel, "新規保存確認") Then

            sSQL = "insert into tbSTECSVOutputDefGroup values ( "
            sSQL = sSQL & Trim(str(cboTarget.ItemData(cboTarget.ListIndex)))
            sSQL = sSQL & " , '" & sName & "' "
            sSQL = sSQL & " , 1 "
            sSQL = sSQL & " , " & Trim(str((cboDefSet.ListCount + 1) * 10))
            sSQL = sSQL & " ) "

            Call g_obj_Conn.Execute(sSQL)

            sSQL = "select max(iCSVOutputDefGroupID) from tbSTECSVOutputDefGroup where vCSVOutputDefGroupName = '" & sName & "' "
            Set oRs = g_obj_Conn.Execute(sSQL)

            lID = oRs.Fields(0)
            oRs.Close
            Set oRs = Nothing
        Else
            MsgBox "上書き保存する場合は、右の入力をクリアして再度保存を押してください。", vbOKOnly
            Exit Sub
        End If
    Else
        sName = Trim(cboDefSet.Text)
        If vbCancel = MsgBox("指定内容を「" & sName & "」に上書き保存します。" & vbCrLf & "よろしいですか？", vbOKCancel, "新規保存確認") Then
            MsgBox "新規保存する場合は、右の入力に名称を入力して再度保存を押してください。", vbOKOnly
            Exit Sub
        End If
        lID = cboDefSet.ItemData(cboDefSet.ListIndex)
    End If

    sSQL = "delete from tbSTECSVOutputDefGroupCols where iCSVOutputDefGroupID = " & Trim(str(lID))
    Call g_obj_Conn.Execute(sSQL)

    For lRow = 1 To vsfOutputCol.Rows - 1

        sSQL = "insert into tbSTECSVOutputDefGroupCols values ( "
        sSQL = sSQL & Trim(str(lID))
        sSQL = sSQL & " , " & vsfOutputCol.TextMatrix(lRow, prvlIdCol)
        sSQL = sSQL & " , " & IIf(vsfOutputCol.TextMatrix(lRow, prvlOutputCol) = prvcsOutputOn, 1, 0)

        For i = LBound(prvuSortType_) To UBound(prvuSortType_)
            If vsfOutputCol.TextMatrix(lRow, prvlSortCol) = prvuSortType_(i).sSortDisp Then
                sSQL = sSQL & " , " & prvuSortType_(i).iID
                Exit For
            End If
        Next i

        sSQL = sSQL & " , " & IIf(vsfOutputCol.TextMatrix(lRow, prvlSortNoCol) = "", "NULL", vsfOutputCol.TextMatrix(lRow, prvlSortNoCol))

        For i = LBound(prvsWhereType) To UBound(prvsWhereType)
            If vsfOutputCol.TextMatrix(lRow, prvlWhereTypeCol) = prvsWhereType(i) Then
                sSQL = sSQL & " , " & Trim(str(i))
                Exit For
            End If
        Next i

        sSQL = sSQL & " , '" & vsfOutputCol.TextMatrix(lRow, prvlWhereCol) & "' "
        sSQL = sSQL & " , '" & Trim(str(lRow * 10)) & "' "
        sSQL = sSQL & " ) "

        Call g_obj_Conn.Execute(sSQL)

    Next

    Call lsSetDefSet

    For i = 0 To cboDefSet.ListCount - 1
        If cboDefSet.ItemData(i) = lID Then
            cboDefSet.ListIndex = i
            Exit For
        End If
    Next

    Exit Sub

ErrProc:
    MsgBox "lsSaveDefset関数でエラーが発生しました。", vbOKOnly, "エラー"

End Sub

Private Sub cmdSelect_Click()

    On Error GoTo ErrProc

    Dim lRow    As Long
    Dim sWk     As String
    Dim sSerial As String


    If vsfColList.Row <= 0 Then
        MsgBox "選択されていません。", vbOKOnly, "選択エラー"
        Exit Sub
    End If

    sSerial = Trim(txtSerial.Text)
    If sSerial <> "" Then
        If Not gf_LongCheck(sSerial) Then
            MsgBox "選択されていません。", vbOKOnly, "選択エラー"
            Exit Sub
        End If
    End If

    sWk = ""
    sWk = sWk & Trim(str(vsfOutputCol.Rows))
    sWk = sWk & vbTab & vsfColList.TextMatrix(vsfColList.Row, prvlCL_IdCol)
    sWk = sWk & vbTab & vsfColList.TextMatrix(vsfColList.Row, prvlCL_NameCol)
    sWk = sWk & vbTab & vsfColList.TextMatrix(vsfColList.Row, prvlCL_OutputCol)
    sWk = sWk & vbTab & prvcsOutputOn

    If sSerial = "" Then
        lRow = vsfOutputCol.Rows
    Else
        If CLng(sSerial) >= vsfOutputCol.Rows Then
            lRow = vsfOutputCol.Rows
        Else
            lRow = CLng(sSerial)
        End If
    End If

    vsfOutputCol.AddItem sWk, lRow

    If lRow < vsfOutputCol.Rows Then
        Call lsResetRowNumber(lRow)
    End If

    Exit Sub

ErrProc:
    MsgBox "cmdSelect_Click関数で、エラーが発生しました。", vbOKOnly, "エラー"



End Sub

Private Sub lsResetRowNumber(plRow As Long)

    Dim lRow As Long

    For lRow = CLng(plRow) To vsfOutputCol.Rows - 1
        vsfOutputCol.TextMatrix(lRow, prvlRowNoCol) = Trim(str(lRow))
    Next

End Sub


Private Sub lsSetDefSet()

    On Error GoTo ErrProc

    Dim sSQL  As String
    Dim oRs   As ADODB.Recordset


    cboDefSet.Clear

    sSQL = "select iCSVOutputDefGroupID , vCSVOutputDefGroupName from tbSTECSVOutputDefGroup"
    sSQL = sSQL & " where iCSVOutputObjectID = " & Trim(str(cboTarget.ItemData(cboTarget.ListIndex)))
    sSQL = sSQL & " and siViewFlag = 1 "
    sSQL = sSQL & " order by iDispOrder "
 
   Set oRs = g_obj_Conn.Execute(sSQL)

    Do Until oRs.EOF
        cboDefSet.AddItem oRs.Fields("vCSVOutputDefGroupName")
        cboDefSet.ItemData(cboDefSet.NewIndex) = oRs.Fields("iCSVOutputDefGroupID")
        oRs.MoveNext
    Loop

    oRs.Close
    Set oRs = Nothing

    Exit Sub

ErrProc:
On Error Resume Next
    oRs.Close
    Set oRs = Nothing

End Sub

Private Sub vsfColList_DblClick()

    Call cmdSelect_Click

End Sub

Private Sub vsfOutputCol_Click()

    Dim iLoopCnt As Integer

    If vsfOutputCol.Row <= 0 Then Exit Sub

    Select Case vsfOutputCol.Col
    Case prvlWhereTypeCol
'        If Trim(vsfOutputCol.Text) = "" Then
'            vsfOutputCol.Text = prvsWhereType(LBound(prvsWhereType))
'            Exit Sub
'        End If
        For iLoopCnt = LBound(prvsWhereType) To UBound(prvsWhereType)
            If prvsWhereType(iLoopCnt) = Trim(vsfOutputCol.Text) Then
                If iLoopCnt = UBound(prvsWhereType) Then
'                    vsfOutputCol.Text = ""
                    vsfOutputCol.Text = prvsWhereType(LBound(prvsWhereType))
                Else
                    vsfOutputCol.Text = prvsWhereType(iLoopCnt + 1)
                End If
                Exit For
            End If
        Next
    Case prvlOutputCol
        If vsfOutputCol.Text = prvcsOutputOn Then
            vsfOutputCol.Text = prvcsOutputOff
        Else
            vsfOutputCol.Text = prvcsOutputOn
        End If
    Case prvlSortCol
        If vsfOutputCol.Text = prvcsSortASC Then
            vsfOutputCol.Text = prvcsSortDESC
        ElseIf vsfOutputCol.Text = prvcsSortDESC Then
            vsfOutputCol.Text = prvcsSortNULL
        Else
            vsfOutputCol.Text = prvcsSortASC
        End If
    Case prvlWhereCol, prvlSortNoCol
        vsfOutputCol.EditCell
    End Select

End Sub

Private Sub vsfOutputCol_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Select Case Col
    Case prvlWhereCol, prvlSortNoCol
    Case Else
        Cancel = True
    End Select

End Sub

'*******************************************************************************
'* パターン名 表示                                                             *
'* 2021.12.23 cyosa jhi                                                        *
'*******************************************************************************
Private Sub cmdShowDefset_Click()

    Call lsShowDefGroup

End Sub

Private Sub UpDown1_DownClick()

''''txtNendo.Text = Format(DateValue(Trim(str(Trim(Year(CDate(txtNendo.Text & "1月1日")) - 1)) & "/01/01")), "gggee年")
    txtNendo.Text = Format(DateValue(Trim(str(Trim(Year(CDate(txtNendo.Text & "1月1日")) - 1)) & "/01/01")), "yyyy年")

End Sub

Private Sub UpDown1_UpClick()
''''txtNendo.Text = Format(DateValue(Trim(str(Trim(Year(CDate(txtNendo.Text & "1月1日")) + 1)) & "/01/01")), "gggee年")
    txtNendo.Text = Format(DateValue(Trim(str(Trim(Year(CDate(txtNendo.Text & "1月1日")) + 1)) & "/01/01")), "yyyy年")
End Sub
