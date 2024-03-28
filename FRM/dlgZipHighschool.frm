VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form dlgZipHighschool 
   Caption         =   "2438"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   Picture         =   "dlgZipHighschool.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
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
      Left            =   270
      TabIndex        =   6
      Top             =   6780
      Width           =   1350
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
      Left            =   1965
      TabIndex        =   5
      Top             =   6780
      Width           =   1350
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
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   1455
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
      Left            =   4665
      TabIndex        =   2
      Top             =   1635
      Width           =   1350
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
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
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
      Left            =   2280
      TabIndex        =   0
      Top             =   1680
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid grdArea 
      Height          =   3885
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6853
      _Version        =   393216
      FixedCols       =   0
      ForeColor       =   8388608
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
      Left            =   240
      TabIndex        =   10
      Top             =   720
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
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
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
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
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
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   6855
   End
End
Attribute VB_Name = "dlgZipHighschool"
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

Public Sub f_void_FormatGridGeneral(ByRef g_obj_Grid As Object, ByVal g_int_cols As Integer, ByVal g_str_Header As String)
    
    Dim l_str_Header As String
    Dim l_int_Cnt As Integer
    On Error GoTo ErrorHandler
    With g_obj_Grid
        .Rows = 1
        .Cols = g_int_cols
        .FormatString = g_str_Header
        For l_int_Cnt = 0 To g_int_cols - 1
            .ColWidth(l_int_Cnt) = 2000
        Next
    End With
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub CancelButton_Click()
    Unload Me
    frmHighSchoolType.Show
End Sub

Private Sub cmdSearch_Click()
    Dim l_str_SearchString As String
    On Error GoTo ErrorHandler
    
    f_str_sql = ""
    l_str_SearchString = ""
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
            l_str_SearchString = l_str_SearchString & " AND  vCityName LIKE '" & Trim(txtCity.Text) & "%'"
        Else
'            l_str_SearchString = l_str_SearchString & " vCityName LIKE '%" & Trim(txtCity.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " vCityName LIKE '" & Trim(txtCity.Text) & "%'"
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
    txtZipCode.Text = frmHighSchoolType.txtZipCodeId.Text
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
            l_str_Sql = "SELECT vPrefectureName, vCityName, ISNULL(vAddress1,'') as vAddress1, ISNULL(vAddress2,'') as vAddress2 FROM tbSTEZipCodeMaster"
            l_str_Sql = l_str_Sql & " WHERE vZipCodeName='" & Trim(txtZipCode.Text) & "'"
            l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
            If Not l_obj_Rst.EOF Then
'                frmHighSchoolType.txtZipAddress.Text = l_obj_Rst("vPrefectureName") & "," & l_obj_Rst("vCityName") & "," & _
                l_obj_Rst("vAddress1") & "," & l_obj_Rst("vAddress2")
                frmHighSchoolType.txtZipCodeId.Text = Trim(txtZipCode.Text)
                frmHighSchoolType.txtAddress1.Text = l_obj_Rst("vAddress1")
                frmHighSchoolType.txtAddress2.Text = l_obj_Rst("vAddress2")
            End If
            l_obj_Rst.Close
            Set l_obj_Rst = Nothing
        End If
        
        frmHighSchoolType.txtZipCodeId.Text = ""
        If txtZipCode.Text <> "" Then
            frmHighSchoolType.txtZipCodeId.Text = txtZipCode.Text
        End If
        Unload Me
        frmHighSchoolType.Show
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
                If Trim(l_obj_Recordset(l_int_Cols + 1)) <> "" Then
                    .Text = Trim(l_obj_Recordset(l_int_Cols + 1))
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



