VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form dlgChangeZipForCheck 
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "2438"
   ClientHeight    =   7515
   ClientLeft      =   3480
   ClientTop       =   2040
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "dlgChangeZipForCheck.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCity 
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
      IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtPrefecture 
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
      IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "1062"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4785
      TabIndex        =   6
      Top             =   1635
      Width           =   1350
   End
   Begin MSFlexGridLib.MSFlexGrid grdArea 
      Height          =   3885
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6853
      _Version        =   393216
      FixedCols       =   0
      ForeColor       =   8388608
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtZipCode 
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "1061"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1890
      TabIndex        =   9
      Top             =   6780
      Width           =   1350
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "1060"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      TabIndex        =   8
      Top             =   6780
      Width           =   1350
   End
   Begin VB.Label lblErrorDetails 
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
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   6855
   End
   Begin VB.Label lblPrefectureName 
      Caption         =   "1305"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblZipCodeId 
      Caption         =   "1304"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblZipCode 
      Caption         =   "1303"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "dlgChangeZipForCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   frmBrowse
'Author         :   Vishal Kamath
'Created On     :
'Description    :   This form makes a provision for inserting data in the Examinee Profile Table.
'Reference      :   Functional Specs Of Maintain Examinee Data Ver 1.0
'**************************************************************************************************
Private f_int_CheckRadio As Integer                         'Checks the status of the radio button clicked
Private f_str_sql As String                                 'Sql
'Private Const g_str_GeneralHeader As String = "^ ZipCode |^Prefecture |^City |^ Address1 |^ Address2 "
Private f_obj_rsZip As New ADODB.Recordset                  'ADO Recordset
Private f_int_Row As Integer                                'Gets the row number in the grid
Private f_str_Address  As String
Dim f_str_GeneralHeader As String

Public goParentForm As Form

Public Sub f_void_FormatGridGeneral(ByRef g_obj_Grid As Object, ByVal g_int_cols As Integer, ByVal g_str_Header As String)
    
    Dim l_str_Header As String
    Dim l_int_Cnt As Integer
    On Error GoTo ErrorHandler
    With g_obj_Grid
        .Rows = 1
        .Cols = g_int_cols
        .FormatString = ">iZipCodeId|" & g_str_Header
        .ColWidth(0) = 0
        For l_int_Cnt = 1 To g_int_cols - 1
            .ColWidth(l_int_Cnt) = 2000
        Next
    End With
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub CancelButton_Click()
    Unload Me
    goParentForm.Show
End Sub

Private Sub cmdSearch_Click()
    Dim l_str_SearchString As String
    On Error GoTo ErrorHandler
    
    f_str_sql = ""
    l_str_SearchString = ""

    If Trim(txtZipCode.Text = "") _
        And Trim(txtPrefecture.Text = "") _
        And Trim(txtCity.Text = "") Then
        MsgBox "åüçıçiçûÇ›ÇÃÇΩÇﬂÅAÇPçÄñ⁄à»è„éwíËÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", vbOKOnly, "ì¸óÕ"
        If txtZipCode.Enabled Then txtZipCode.SetFocus
        Exit Sub
    End If

    If txtZipCode.Text <> "" Then
        If l_str_SearchString <> "" Then
'            l_str_SearchString = l_str_SearchString & " AND  vZipCodeName LIKE '%" & Trim(txtZipCode.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " AND  vZipCodeName LIKE '" & Trim(txtZipCode.Text) & "%'"
        Else
'            l_str_SearchString = l_str_SearchString & " vZipCodeName LIKE '%" & Trim(txtZipCode.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " vZipCodeName LIKE '" & Trim(txtZipCode.Text) & "%'"
        End If
    End If
    
    If txtPrefecture.Text <> "" Then
        If l_str_SearchString <> "" Then
'            l_str_SearchString = l_str_SearchString & " AND  vPrefectureName LIKE '%" & Trim(txtPrefecture.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " AND  vPrefectureName LIKE '" & Trim(txtPrefecture.Text) & "%'"
        Else
'            l_str_SearchString = l_str_SearchString & " vPrefectureName LIKE '%" & Trim(txtPrefecture.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " vPrefectureName LIKE '" & Trim(txtPrefecture.Text) & "%'"
        End If
    End If
    
    If txtCity.Text <> "" Then
        If l_str_SearchString <> "" Then
'            l_str_SearchString = l_str_SearchString & " AND  vCityName LIKE '%" & Trim(txtCity.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " AND  vAddress1 LIKE '" & Trim(txtCity.Text) & "%'"
        Else
'            l_str_SearchString = l_str_SearchString & " vCityName LIKE '%" & Trim(txtCity.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " vAddress1 LIKE '" & Trim(txtCity.Text) & "%'"
        End If
    End If
       
    f_str_sql = "Select * from tbSTEZipCodeMaster"
    If l_str_SearchString <> "" Then
        f_str_sql = f_str_sql & " WHERE " & l_str_SearchString
    End If
    
    If grdArea.Rows > 1 Then
        Call f_void_FormatGridGeneral(grdArea, 5, f_str_GeneralHeader)
    End If
    Call fp_Void_PopulateGrid(grdArea, f_str_sql)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    LoadResStrings Me
    Me.Caption = LoadResString(2438)
    Call g_void_SetFontProperties(Me)     ' set the font properties
    g_void_OpenConnection
    f_str_GeneralHeader = LoadResString(2509)
    txtZipCode.Text = goParentForm.txtZipCode.Text
    Call f_void_FormatGridGeneral(grdArea, 5, f_str_GeneralHeader)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub grdArea_Click()
    On Error GoTo ErrorHandler
    With grdArea
        f_int_Row = .Row
        .Col = 0
        txtZipCode.Tag = .Text
        .Col = .Col + 1
        txtZipCode.Text = .Text
        .Col = .Col + 1
        txtPrefecture.Text = .Text
        .Col = .Col + 1
        txtCity.Text = .Text
    End With
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub OKButton_Click()
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    Dim l_int_RetVal As Integer
    On Error GoTo ErrorHandler
    
    If txtZipCode.Text = "" Then
        lblErrorDetails.Caption = LoadResString(2444)
        Exit Sub
    End If
    lblErrorDetails.Caption = ""
    l_int_RetVal = MsgBox(LoadResString(2439) & txtZipCode.Text & LoadResString(2440), vbYesNo + vbQuestion)
    If l_int_RetVal = vbYes Then
        If Len(Trim(txtZipCode.Text)) <> 0 Then
            l_str_Sql = "SELECT vPrefectureName, vCityName, vAddress1, vAddress2 FROM tbSTEZipCodeMaster"
            l_str_Sql = l_str_Sql & " WHERE vZipCodeName='" & Trim(txtZipCode.Text) & "'"
            l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
            If Not l_obj_Rst.EOF Then
                goParentForm.txtZipAddress.Text = l_obj_Rst("vPrefectureName") & "," & l_obj_Rst("vCityName") & "," & _
                                                l_obj_Rst("vAddress1") & "," & l_obj_Rst("vAddress2")
                goParentForm.txtZipPref.Text = l_obj_Rst("vPrefectureName")
            End If
            l_obj_Rst.Close
            Set l_obj_Rst = Nothing
        End If
        
        goParentForm.txtZipCodeId.Text = ""
        If txtZipCode.Text <> "" Then
            goParentForm.txtZipCodeId.Text = txtZipCode.Tag
            goParentForm.txtZipCode.Text = txtZipCode.Text
        End If
        Unload Me
        goParentForm.Show
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Public Sub fp_Void_PopulateGrid(g_obj_Grid As Object, g_str_Sql As String)
   
     'Populates Grid
    Dim l_obj_Recordset As ADODB.Recordset
    Dim l_str_Department As String                                 'sql statement
    Dim l_int_Cnt As Integer                                       'Counter for the recordset
    Dim l_int_Cols As Integer
    
    On Error GoTo ErrorHandler
    
    Set l_obj_Recordset = New ADODB.Recordset                  'Recordset Variable

    l_obj_Recordset.Open g_str_Sql, g_obj_Conn, adOpenForwardOnly, adLockReadOnly

   If l_obj_Recordset.EOF Then
        lblErrorDetails.Caption = LoadResString(1964)
   Else
    lblErrorDetails.Caption = ""
    OKButton.Enabled = True
    Do While Not l_obj_Recordset.EOF
        g_obj_Grid.Rows = g_obj_Grid.Rows + 1
        g_obj_Grid.Row = g_obj_Grid.Rows - 1
        
        With g_obj_Grid
            For l_int_Cols = 0 To .Cols - 1
                .Col = l_int_Cols
                If Trim(l_obj_Recordset(l_int_Cols)) <> "" Then
                    .Text = Trim(l_obj_Recordset(l_int_Cols))
                Else
                    .Text = ""
                End If
            Next
        End With
        
        l_obj_Recordset.MoveNext
    Loop
    l_obj_Recordset.Close
    Set l_obj_Recordset = Nothing
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub txtZipCode_KeyPress(KeyAscii As Integer)
    Call NumericOnly(Me, KeyAscii)
End Sub
