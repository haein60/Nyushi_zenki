VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form dlgSeisekiChohyoIchiran 
   BorderStyle     =   3  'ŒÅ’èÀÞ²±Û¸Þ
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
   Begin VSFlex7LCtl.VSFlexGrid vsfData 
      Height          =   4575
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   5775
      _cx             =   10186
      _cy             =   8070
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
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
      FormatString    =   "<Number|<Name"
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
      ShowComboButton =   1
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
   Begin VB.CommandButton cmdCnf 
      Caption         =   "CANCEL"
      Height          =   495
      Index           =   1
      Left            =   4920
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCnf 
      Caption         =   "OK"
      Height          =   495
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "dlgSeisekiChohyoIchiran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private prvsID As String

Private Sub cmdCnf_Click(Index As Integer)

    If Index = 0 Then
        If vsfData.Row <> 0 Then
            prvsID = vsfData.TextMatrix(vsfData.Row, 0)
        End If
    Else
        prvsID = ""
    End If

    Unload Me

End Sub

Private Sub Form_Load()

Dim sSQL As String
Dim oRs As ADODB.Recordset

On Error GoTo ErrProc

    vsfData.ColWidth(0) = 1200
    vsfData.ColWidth(1) = 3000

    sSQL = "SELECT "
    sSQL = sSQL & "  isnull( vPrintCommandName , '' ) "
    sSQL = sSQL & ", isnull( vPrintControlFileName , '' ) "
    sSQL = sSQL & " FROM tbSTRPrintCommandProfile "
    sSQL = sSQL & " ORDER BY vPrintControlFileName "

    Set oRs = g_obj_Conn.Execute(sSQL)

    With vsfData

        .Rows = 1

        Do Until oRs.EOF
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
            .Text = oRs.Fields(1)
            .Col = 1
            .Text = oRs.Fields(0)
            oRs.MoveNext
        Loop

    End With

    oRs.Close
    Set oRs = Nothing

Exit Sub
ErrProc:

End Sub

Public Sub getPrintCommandId(psID As String)

    Me.Show vbModal

    psID = prvsID

End Sub
