VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSystemData 
   Caption         =   "frmSystemData : 試験年度設定"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmSystemData.frx":0000
   ScaleHeight     =   10485
   ScaleWidth      =   14085
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdClear 
      Caption         =   "ｸﾘｱ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6735
      TabIndex        =   17
      Top             =   3210
      Width           =   615
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ｸﾘｱ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6735
      TabIndex        =   16
      Top             =   2610
      Width           =   615
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ｸﾘｱ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6735
      TabIndex        =   15
      Top             =   2010
      Width           =   615
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   6735
      TabIndex        =   14
      Top             =   1425
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5295
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtDay 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00400000&
      Height          =   440
      Index           =   2
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   9
      Tag             =   "[iZipCodeId]"
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtDay 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00400000&
      Height          =   440
      Index           =   1
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   6
      Tag             =   "[iZipCodeId]"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtDay 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00400000&
      Height          =   440
      Index           =   0
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   3
      Tag             =   "[iZipCodeId]"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtNendo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   405
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "[iZipCodeId]"
      Top             =   1440
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtSetDay 
      Height          =   435
      Left            =   7215
      TabIndex        =   13
      Top             =   1455
      Visible         =   0   'False
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   767
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月DD日"
      Format          =   16711683
      CurrentDate     =   43493
   End
   Begin VB.Label lblMsg2 
      BackStyle       =   0  '透明
      Caption         =   "lblMsg"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1035
      TabIndex        =   23
      Top             =   5535
      Width           =   11925
   End
   Begin VB.Label lblMikameBK 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "２次試験３日目日付"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   7710
      TabIndex        =   22
      Top             =   3270
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblFirstDayBK 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "２次試験１日目日付"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   7710
      TabIndex        =   21
      Top             =   2070
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblHutukameBK 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "２次試験１日目日付"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   7695
      TabIndex        =   20
      Top             =   2700
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  '透明
      Caption         =   "lblMsg"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1050
      TabIndex        =   19
      Top             =   5145
      Width           =   9000
   End
   Begin VB.Label lblNendo_org 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "年度（和暦）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   9390
      TabIndex        =   18
      Top             =   1545
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label lblMikame 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "２次試験日 ３日目"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   1920
      TabIndex        =   11
      Top             =   3330
      Width           =   2355
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  '透明
      Caption         =   "*"
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
      Height          =   255
      Index           =   3
      Left            =   4305
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHutukame 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "２次試験日 ２日目"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   1920
      TabIndex        =   8
      Top             =   2730
      Width           =   2325
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  '透明
      Caption         =   "*"
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
      Height          =   255
      Index           =   2
      Left            =   4305
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblFirstDay 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "２次試験日 １日目"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   1920
      TabIndex        =   5
      Top             =   2130
      Width           =   2325
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  '透明
      Caption         =   "*"
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
      Height          =   255
      Index           =   0
      Left            =   4305
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblNendo 
      BackColor       =   &H00F4DBC4&
      BackStyle       =   0  '透明
      Caption         =   "入試年度"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   1920
      TabIndex        =   2
      Top             =   1545
      Width           =   1545
   End
   Begin VB.Label lblErrIndicator 
      BackStyle       =   0  '透明
      Caption         =   "*"
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
      Height          =   255
      Index           =   1
      Left            =   4305
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmSystemData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private prvoSetObj          As Object
Private prvlSystemProfileId As Long

Private Sub Form_Load()

    On Error GoTo ErrHandler

    Dim oRs    As ADODB.Recordset
    Dim lIndex As Long
    Dim sSQL   As String

    LoadResStrings Me

    Call g_void_SetFontProperties(Me)     ' set the font properties

    Me.Caption = "frmSystemData : 試験年度設定"  ''''LoadResString(2600)


    lblMsg.Caption = "入試年度更新により、指定入試年度のデータを閲覧することが出来ます。"


    lblMsg2.ForeColor = &HFF
    lblMsg2.Caption = "※過去入試年度を指定、更新した場合、システムを再起動してから指定年度の情報をご閲覧ください。"



    sSQL = "SELECT Top 1 iSystemProfileId , isnull( iNendo , year(getdate()) ) as iNendo From tbSTESystemProfile where iActiveFlag = 1 Order by iSystemProfileID desc "

    Set oRs = g_obj_Conn.Execute(sSQL)

    If oRs.EOF Then
        MsgBox "システムパラメータが見つかりませんでした。" & vbCrLf & "時間を置いて、再度実行してください。", vbOKOnly, "エラー"
        Set oRs = Nothing
        Exit Sub
    Else
        prvlSystemProfileId = oRs.Fields(0)

''''    2020.01.27 del jhi
''''    txtNendo.Text = Format(DateValue(Trim(str(Trim(oRs.Fields(1))) & "/01/01")), "gggee年")

''''    2020.01.27 add jhi
        txtNendo.Text = Format(DateValue(Trim(str(Trim(oRs.Fields(1))) & "/01/01")), "yyyy年")
    End If

    oRs.Close
    Set oRs = Nothing

    sSQL = "SELECT "
    sSQL = sSQL & "  isnull( convert( varchar , dtSecondExamDay1 , 111 ) , '' ) as dtSecondExamDay1 "
    sSQL = sSQL & ", isnull( convert( varchar , dtSecondExamDay2 , 111 ) , '' ) as dtSecondExamDay2 "
    sSQL = sSQL & ", isnull( convert( varchar , dtSecondExamDay3 , 111 ) , '' ) as dtSecondExamDay3 "
    sSQL = sSQL & " From tbSTESecondExamProfile where iSystemProfileId = " & prvlSystemProfileId

    Set oRs = g_obj_Conn.Execute(sSQL)

    If oRs.EOF Then
        MsgBox "システムパラメータが見つかりませんでした。" & vbCrLf & "時間を置いて、再度実行してください。", vbOKOnly, "エラー"
        Set oRs = Nothing
        Exit Sub
    Else
        For lIndex = 0 To 2
            txtDay(lIndex).Text = Trim(oRs.Fields(lIndex))
        Next
        For lIndex = 0 To 2
            If txtDay(lIndex).Text <> "" Then
'                txtDay(lIndex).Text = "2004/02/03"
                txtDay(lIndex).Text = g_dt_ConvertDate(txtDay(lIndex).Text)
            End If
        Next
    End If

    oRs.Close
    Set oRs = Nothing

    Exit Sub

ErrHandler:
    MsgBox Err.Description

End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler
    Dim i As Integer


    fMainForm.mnuTools.Enabled = False  ' disable tools menu

    For i = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(i).Enabled = False
    Next i

    Exit Sub

ErrorHandler:

''''2020.01.27 del jhi
''''MsgBox Err.Description, vbInformation, LoadResString(1729)

''''2020.01.27 add jhi
    MsgBox Err.Description, vbInformation, "エラー" 'LoadResString(1729)は、"エラー"と定義されている

End Sub


'*******************************************************************************
'* 【更新】ボタンの処理                                                        *
'*******************************************************************************
Private Sub cmdUpdate_Click()

    On Error GoTo ErrHandler

    Dim sSQL       As String
    Dim rinf       As Long
    Dim strYYYY    As String


    ''''2021.1.09 add jhi
    rinf = myMsgBox("入試年度を更新します。よろしいですか?", "確認")
    If (rinf = vbCancel) Then
       Exit Sub
    End If


    strYYYY = Year(CDate(txtNendo.Text & "1月1日"))

''''2023.01.23 del jhi
''''sSQL = "Update tbSTESystemProfile SET iNendo = " & Year(CDate(txtNendo.Text & "1月1日")) & " where iActiveFlag = " & prvlSystemProfileId

''''年度セットをすると無条件iCurrentPhaseを1フェーズにする 2023.01.23 add jhi
    sSQL = "Update tbSTESystemProfile SET iNendo = " & Year(CDate(txtNendo.Text & "1月1日")) & ",iCurrentPhase=0 where iActiveFlag = " & prvlSystemProfileId

    g_obj_Conn.Execute sSQL

    sSQL = "Update tbSTESecondExamProfile SET "
    If txtDay(0).Text = "" Then
        sSQL = sSQL & "  dtSecondExamDay1 = NULL "
    Else
        sSQL = sSQL & "  dtSecondExamDay1 = '" & Format(CDate(txtDay(0).Text), "YYYY/MM/DD") & "' "
    End If

    If txtDay(1).Text = "" Then
        sSQL = sSQL & ", dtSecondExamDay2 = NULL "
    Else
        sSQL = sSQL & ", dtSecondExamDay2 = '" & Format(CDate(txtDay(1).Text), "YYYY/MM/DD") & "' "
    End If

    If txtDay(2).Text = "" Then
        sSQL = sSQL & ", dtSecondExamDay3 = NULL "
    Else
        sSQL = sSQL & ", dtSecondExamDay3 = '" & Format(CDate(txtDay(2).Text), "YYYY/MM/DD") & "' "
    End If

    sSQL = sSQL & " where iSystemProfileId = " & prvlSystemProfileId

    g_obj_Conn.Execute sSQL


    ''''設定年度を global変数に設定 2021.12.22 add jhi
    g_int_CurrentNendo = CInt(strYYYY)

    MsgBox "入試年度を更新しました。 指定入試年度のデータの閲覧が出来ます。", vbOKOnly, "更新完了"

    Exit Sub

ErrHandler:
    MsgBox Err.Description

End Sub

Private Sub cmdClear_Click(Index As Integer)

    txtDay(Index).Text = ""

End Sub


Private Sub dtSetDay_Change()

    dtSetDay.Visible = False
    prvoSetObj.Text = g_dt_ConvertDate(dtSetDay.Value)
    prvoSetObj.ZOrder 0

End Sub

Private Sub dtSetDay_LostFocus()

    dtSetDay.Visible = False
    prvoSetObj.Text = g_dt_ConvertDate(dtSetDay.Value)
    prvoSetObj.ZOrder 0

End Sub

Private Sub txtDay_GotFocus(Index As Integer)

    ' position date picker control over cell
    With txtDay(Index)

        dtSetDay.Move .Left, .Top, .Width, .Height

        If Trim(.Text) <> "" Then
            ' initialize value, save original in tag in case user hits escape
            'dtBirthDay.Value = .Text
            'dtBirthDay.Tag = .Text
            
            ' changed the above two lines in Comdesign , arka 9Apr 2002
            
            dtSetDay.Value = g_dt_ConvertDate(.Text)
            dtSetDay.Tag = g_dt_ConvertDate(.Text)
        Else
            dtSetDay.Tag = #1/1/2022#
        End If

        ' show and activate date picker control
        dtSetDay.Visible = True
        dtSetDay.SetFocus

        Set prvoSetObj = txtDay(Index)

    End With

    ' make it drop down the calendar
    SendKeys "{f4}"

End Sub

Private Sub UpDown1_DownClick()

''''2020.01.27 del jhi
''''txtNendo.Text = Format(DateValue(Trim(str(Trim(Year(CDate(txtNendo.Text & "1月1日")) - 1)) & "/01/01")), "gggee年")

''''2020.01.27 add jhi
    txtNendo.Text = Format(DateValue(Trim(str(Trim(Year(CDate(txtNendo.Text & "1月1日")) - 1)) & "/01/01")), "yyyy年")

End Sub

Private Sub UpDown1_UpClick()

''''2020.01.27 del jhi
''''txtNendo.Text = Format(DateValue(Trim(str(Trim(Year(CDate(txtNendo.Text & "1月1日")) + 1)) & "/01/01")), "gggee年")

''''2020.01.27 add jhi
    txtNendo.Text = Format(DateValue(Trim(str(Trim(Year(CDate(txtNendo.Text & "1月1日")) + 1)) & "/01/01")), "yyyy年")

End Sub
