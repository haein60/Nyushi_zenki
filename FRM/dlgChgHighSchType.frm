VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form dlgChgHighSchType 
   BorderStyle     =   3  'ŒÅ’èÀÞ²±Û¸Þ
   Caption         =   "2441"
   ClientHeight    =   7110
   ClientLeft      =   4920
   ClientTop       =   2040
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "dlgChgHighSchType.frx":0000
   ScaleHeight     =   7110
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtHighSchoolCode 
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
      Height          =   420
      Left            =   2730
      TabIndex        =   1
      Top             =   720
      Width           =   1335
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
      Left            =   4800
      TabIndex        =   4
      Top             =   1200
      Width           =   1350
   End
   Begin VB.TextBox txtHighSchoolName 
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
      Height          =   420
      Left            =   2730
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
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
      Left            =   1830
      TabIndex        =   7
      Top             =   6330
      Width           =   1350
   End
   Begin VB.CommandButton cmdOk 
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
      Left            =   360
      TabIndex        =   6
      Top             =   6330
      Width           =   1350
   End
   Begin MSFlexGridLib.MSFlexGrid grdArea 
      Height          =   3885
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   5745
      _ExtentX        =   10134
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
      TabIndex        =   8
      Top             =   1800
      Width           =   5655
   End
   Begin VB.Label lblHighSchoolCode1 
      Caption         =   "1103"
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
      Left            =   270
      TabIndex        =   0
      Top             =   735
      Width           =   2250
   End
   Begin VB.Label lblHighSchoolName 
      Caption         =   "1104"
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
      Left            =   285
      TabIndex        =   2
      Top             =   1200
      Width           =   2250
   End
End
Attribute VB_Name = "dlgChgHighSchType"
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
'Private Const g_str_GeneralHeader As String = "^ HighSchoolCode |^ HighSchoolName "
Private f_obj_rsZip As New ADODB.Recordset                  'ADO Recordset
Private f_int_Row As Integer                                'Gets the row number in the grid
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
            .ColWidth(l_int_Cnt) = 3000
        Next
    End With
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub CancelButton_Click()
    Unload Me
    frmExamineeProfile.Show
End Sub

Private Sub cmdSearch_Click()
    Dim l_str_SearchString As String
    On Error GoTo ErrorHandler
    
    f_str_sql = ""
    l_str_SearchString = ""
    If txtHighSchoolCode.Text <> "" Then
        If l_str_SearchString <> "" Then
'            l_str_SearchString = l_str_SearchString & " AND  vHighSchoolCode LIKE '%" & Trim(txtHighSchoolCode.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " AND  vHighSchoolCode LIKE '" & Trim(txtHighSchoolCode.Text) & "%'"
        Else
'            l_str_SearchString = l_str_SearchString & " vHighSchoolCode LIKE '%" & Trim(txtHighSchoolCode.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " vHighSchoolCode LIKE '" & Trim(txtHighSchoolCode.Text) & "%'"
        End If
    End If

    If txtHighSchoolName.Text <> "" Then
        If l_str_SearchString <> "" Then
'            l_str_SearchString = l_str_SearchString & " AND  vHighSchoolName LIKE '%" & Trim(txtHighSchoolName.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " AND  vHighSchoolName LIKE '" & Trim(txtHighSchoolName.Text) & "%'"
        Else
'            l_str_SearchString = l_str_SearchString & " vHighSchoolName LIKE '%" & Trim(txtHighSchoolName.Text) & "%'"
            l_str_SearchString = l_str_SearchString & " vHighSchoolName LIKE '" & Trim(txtHighSchoolName.Text) & "%'"
        End If
    End If

    f_str_sql = "Select vHighSchoolCode,vHighSchoolName FROM tbSTEHighSchoolType"
    
    If l_str_SearchString <> "" Then
        f_str_sql = f_str_sql & " WHERE " & l_str_SearchString
    End If

    If grdArea.Rows > 1 Then
        Call f_void_FormatGridGeneral(grdArea, 2, f_str_GeneralHeader)
    End If
    Call fp_Void_PopulateGrid(grdArea, f_str_sql)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    LoadResStrings Me
    Me.Caption = LoadResString(2441)
    Call g_void_SetFontProperties(Me)     ' set the font properties
    f_str_GeneralHeader = LoadResString(2510)
    txtHighSchoolCode.Text = frmExamineeProfile.txtHighSchoolID.Text
    Call f_void_FormatGridGeneral(grdArea, 2, f_str_GeneralHeader)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub grdArea_Click()
    On Error GoTo ErrorHandler
    f_int_Row = grdArea.Row
    With grdArea
        .Col = 0
        txtHighSchoolCode.Text = .Text
        .Col = 1
        txtHighSchoolName.Text = .Text
    End With
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Private Sub cmdOK_Click()
    Dim l_int_RetVal As Integer
    On Error GoTo ErrorHandler
    
    If txtHighSchoolCode.Text = "" Then
        lblErrorDetails.Caption = LoadResString(2443)
        Exit Sub
    End If
    lblErrorDetails.Caption = ""
    l_int_RetVal = MsgBox(LoadResString(2442) & txtHighSchoolCode.Text & LoadResString(2440), vbYesNo + vbQuestion)
    If l_int_RetVal = vbYes Then
        frmExamineeProfile.txtHighSchoolID.Text = txtHighSchoolCode.Text
        Unload Me
        frmExamineeProfile.Show
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
    cmdOK.Enabled = True
    Do While Not l_obj_Recordset.EOF
        g_obj_Grid.Rows = g_obj_Grid.Rows + 1
        g_obj_Grid.Row = g_obj_Grid.Rows - 1
        
        With g_obj_Grid
            For l_int_Cols = 0 To .Cols - 1
                .Col = l_int_Cols
                If l_obj_Recordset.Fields(l_int_Cols).Value = Null Then
                    .Text = ""
                Else
                    .Text = l_obj_Recordset.Fields(l_int_Cols).Value
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


