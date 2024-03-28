VERSION 5.00
Begin VB.Form dlgExcelReportInput 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ﾀﾞｲｱﾛｸﾞ ｷｬﾌﾟｼｮﾝ"
   ClientHeight    =   705
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "dlgExcelReportInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private g_nTop As Long

Private prvsMandatoryCtls() As String

Private Function chkMandatory() As Long

Dim lCnt As Long
Dim luCnt As Long
Dim llCnt As Long
Dim oCtl As Object

chkMandatory = -1

    llCnt = LBound(prvsMandatoryCtls)
    luCnt = UBound(prvsMandatoryCtls)

    If luCnt = 0 Then Exit Function

    For lCnt = llCnt + 1 To luCnt
        For Each oCtl In Me.Controls
            If oCtl.Name = prvsMandatoryCtls(lCnt) Then
                Select Case TypeName(oCtl)
                Case "TextBox"
                    If Trim(oCtl.Text) = "" Then
                        chkMandatory = lCnt
                        If oCtl.Enabled Then
                            oCtl.SetFocus
                        End If
                        Exit Function
                    End If
                Case "ComboBox"
                    If oCtl.ListIndex < 0 Then
                        chkMandatory = lCnt
                        If oCtl.Enabled Then
                            oCtl.SetFocus
                        End If
                        Exit Function
                    End If
                Case "DTPicker"
                End Select
            End If
        Next
    Next

End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// Form_Load
Private Sub Form_Load()
    ' フォームを中央に配置
'    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    g_nTop = 100
    g_bInput = False
    ReDim prvsMandatoryCtls(0)
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// btnOK_Click
Private Sub btnOK_Click()
Dim i As Long

    i = chkMandatory

    If i >= 0 Then
        MsgBox "必須項目が入力されていません。", vbOKOnly, "入力エラー"
        Exit Sub
    End If

    For i = 0 To Me.Controls.Count - 1
        Select Case TypeName(Me.Controls(i))
        Case "TextBox"
            SetParam Me.Controls(i).Tag, Me.Controls(i).Text
        Case "ComboBox"
            If Me.Controls(i).ListIndex >= 0 Then
                If Me.Controls(i).ItemData(Me.Controls(i).ListIndex) > 0 Then
                    SetParam Me.Controls(i).Tag, Me.Controls(i).ItemData(Me.Controls(i).ListIndex)
                Else
                    SetParam Me.Controls(i).Tag, Me.Controls(i).Text
                End If
            Else
                SetParam Me.Controls(i).Tag, Me.Controls(i).Text
            End If
        Case "DTPicker"
            SetParam Me.Controls(i).Tag, Me.Controls(i).Value
        End Select
    Next
    g_bInput = True
    Unload Me
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// AdjustSize
Public Sub AdjustSize()
    g_nTop = g_nTop + 200
    btnOK.Top = g_nTop
    g_nTop = g_nTop + 900
    Me.Height = g_nTop
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// SetTitle
Public Sub SetTitle(sTitle As String)
    Me.Caption = sTitle
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// AddLabel
Public Sub AddLabel(sTitle As String, bMandatory As Boolean)
    Dim sName As String
    Dim objControl As Object
    ' Add Control
    sName = "Ctrl" & CStr(Me.Controls.Count)
    Set objControl = Me.Controls.Add("VB.Label", sName)
    ' Set Prop
    objControl.Top = g_nTop
    g_nTop = g_nTop + 230
    objControl.Left = 200
    objControl.Height = 200
    objControl.Width = 4500
    objControl.Visible = True
    objControl.Caption = sTitle
    If bMandatory Then
        objControl.Caption = objControl.Caption & " (*)"
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// AddText
Public Sub AddText(sVarName As String, sInit As String, bMandatory As Boolean)
    Dim sName As String
    Dim objControl As Object
    ' Add Control
    sName = "Ctrl" & CStr(Me.Controls.Count)
    Set objControl = Me.Controls.Add("VB.TextBox", sName)
    ' Set Prop
    objControl.Top = g_nTop
    g_nTop = g_nTop + 400
    objControl.Left = 200
    objControl.Height = 300
    objControl.Width = 4500
    objControl.Visible = True
    objControl.Tag = sVarName
    objControl.Text = sInit
    If bMandatory Then
        ReDim Preserve prvsMandatoryCtls(UBound(prvsMandatoryCtls) + 1)
        prvsMandatoryCtls(UBound(prvsMandatoryCtls)) = sName
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// AddTextPassword
Public Sub AddTextPassword(sVarName As String, sInit As String, bMandatory As Boolean)
    Dim sName As String
    Dim objControl As Object
    ' Add Control
    sName = "Ctrl" & CStr(Me.Controls.Count)
    Set objControl = Me.Controls.Add("VB.TextBox", sName)
    ' Set Prop
    objControl.Top = g_nTop
    g_nTop = g_nTop + 400
    objControl.Left = 200
    objControl.Height = 300
    objControl.Width = 4500
    objControl.Visible = True
    objControl.Tag = sVarName
    objControl.PasswordChar = "*"
    objControl.Text = sInit
    If bMandatory Then
        ReDim Preserve prvsMandatoryCtls(UBound(prvsMandatoryCtls) + 1)
        prvsMandatoryCtls(UBound(prvsMandatoryCtls)) = sName
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// AddComboBox
Public Sub AddComboBox(sVarName As String, aComboItem() As String, aComboItemId() As Long, bMandatory As Boolean)
    Dim sName As String
    Dim objControl As Object
    ' Add Control
    sName = "Ctrl" & CStr(Me.Controls.Count)
    Set objControl = Me.Controls.Add("VB.ComboBox", sName)
    ' Set Prop
    objControl.Top = g_nTop
    g_nTop = g_nTop + 400
    objControl.Left = 200
    objControl.Width = 4500
    objControl.Visible = True
    objControl.Tag = sVarName
    ' AddItem
    Dim i As Long
    For i = 0 To UBound(aComboItem)
        objControl.AddItem aComboItem(i)
        objControl.ItemData(objControl.NewIndex) = aComboItemId(i)
    Next
    If objControl.ListCount > 0 Then
        objControl.ListIndex = 0
    Else
        objControl.Text = aComboItem(0)
    End If
    If bMandatory Then
        ReDim Preserve prvsMandatoryCtls(UBound(prvsMandatoryCtls) + 1)
        prvsMandatoryCtls(UBound(prvsMandatoryCtls)) = sName
    End If
'    objControl.Text = aComboItem(0)
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'// AddDate
Public Sub AddDate(sVarName As String, sInit As String, bMandatory As Boolean)
    Dim sName As String
    Dim objControl As Object
    ' Add Control
    sName = "Ctrl" & CStr(Me.Controls.Count)
    Set objControl = Me.Controls.Add("MSComCtl2.DTPicker", sName)
    ' Set Prop
    objControl.Top = g_nTop
    g_nTop = g_nTop + 400
    objControl.Left = 200
    objControl.Height = 300
    objControl.Width = 4500
    objControl.Visible = True
    objControl.Tag = sVarName
    objControl.Value = sInit
    If bMandatory Then
        ReDim Preserve prvsMandatoryCtls(UBound(prvsMandatoryCtls) + 1)
        prvsMandatoryCtls(UBound(prvsMandatoryCtls)) = sName
    End If
End Sub

