Attribute VB_Name = "Nyushi"
Option Explicit
'*************************************************************************************************
'Form Name      :   ComboCollection
'Author         :   Dileep Cherian
'Created On     :   10/8/01
'Description    :   This form makes a provision for master maintenance of tbSTRZipCode Table.
'Reference      :   FunctionalSpecs OF MasterMaintanance Ver 1.0.doc
'***************************************************************************************************
'Ammemdments    -   NyushiImpactAnalysisNewChange.doc(ver 1.0)
'Modification History - 05/04/2002 - Mahesh Deshpande, Dileep Cherian
'Caption of master maintenance forms should display the mode in which they are at any time
'ie; Edit, Query or New Mode
'Date values in the grid should also be displayed in Japanese format
'While searching, if no records are found, then the grid should be cleard off
'**************************************************************************************************

Public glUserID                As Long               'ログインユーザID
Public glUserLevel             As Long               'ユーザーの権限レベル 大きいほど高い
Public gsUserPwd               As String             'Excelの読み取りパスワードに使用

Public g_obj_Conn              As ADODB.Connection   ' global connection object
Private l_str_ConnStr          As String             ' to store the connection parameters, retrieved from the registry
Public g_int_ExamType          As Long               ' to keep track of the exam type
Public g_bln_RunLogic          As Boolean            ' to check whether the Distribution Logic is already run or not

Dim m_obj_Rst                  As ADODB.Recordset    '
Public fMainForm               As frmMain            ' global instance of the MDI form

Public g_bln_InterviewHappened As Boolean            ' to check whether interview 1 has happend or not, for report to take place

Public gstrimgpath             As String             ' path of the image files
Public NEWICON                 As String
Public CLEARICON               As String
Public CANCELICON              As String
Public DELETEICON              As String
Public SAVEICON                As String
Public SEARCHICON              As String

Public gsExamCheckNendo        As String
Public gsExamIDFrom            As String
Public gsExamIDTo              As String

'作成後のユーザーとの仕様すり合わせによる改修時、共通化したフォーム上で科目ごとの処理が
'発生したため、ここで科目名を定義しておき、科目名が一致することを処理条件とする
Public Const gcsHyotei         As String = "評定値"  '列見出しがこの値のとき、隣の列が成績概評出力だと判断する
Public Const gcsSeisekiGaihyo  As String = "成績概評"
Public Const gcsKessekiNissuu  As String = "欠席日数"

Private Const prvcExamCheck As String = "frmExamineeCheck"

Public gbExamCheckNewShow As Boolean

Public Const gclExamineeStatus_Default As Long = 0
Public Const gclExamineeStatus_1stPass As Long = 1
Public Const gclExamineeStatus_2ndPass As Long = 2
Public Const gclExamineeStatus_2ndWait As Long = 3
Public Const gclExamineeStatus_2ndWaitPass As Long = 6
Public Const gclExamineeStatus_Refuse As Long = -1

Public Const gclPhase_WaitPass As Long = 3

Public Type puPrintCategoryType
    iID             As Integer
    sDispName       As String
    dDefStartScore  As Double
    dDefEndScore    As Double
    dDefScaleScore  As Double
End Type

'-------------------------------------------------------------------------------
'frmMainからこちに移動 2021.12.28 add jhi
'-------------------------------------------------------------------------------
Public g_int_CurrentNendo    As Integer            ' to store the current year
Public f_int_CurrentPhase    As Integer

Public Type prvuMenues_Type
    oMnuObj As Object
    sTVKey As String
    sIniKey As String
    sCaption As String
    lParent As Long
    bVisible As Boolean
End Type

''''2021.12.28 add jhi globalに宣言する
Public uMenues_() As prvuMenues_Type

''''----------------------------------------------------------------------------
''''条件付きコンパイル引数の設定 2022.02.01 add jhi
''''----------------------------------------------------------------------------
#If zengo_kubun = 1 Then
   Public Const gTit = "前期 - 入試システム"
#Else
   Public Const gTit = "後期 - 入試システム"
#End If

'******************************************************************************
'* 起動 Main Program                                                          *
'*----------------------------------------------------------------------------*
'* ログインチェック                                                           *
'******************************************************************************
Sub Main()

    On Error GoTo ErrorHandler

    Dim oPassCheck  As Object
    Dim lRtn        As Long
    Dim l_bln_Conn  As Boolean                      ' to check the status of database connection
    Dim sSQL        As String
    Dim oRs         As ADODB.Recordset
    Dim sPWD        As String
    Dim sUserId     As String


    'アプリの多重起動を禁止する
    If App.PrevInstance = True Then
        MsgBox "入試システムは既に起動しています。"
        End
    End If

    l_bln_Conn = g_void_OpenConnection()            ' open the database connection

    If Not l_bln_Conn Then
        ' there is an error in opening the database connection, exit the procedure
''''    MsgBox LoadResString(2102), vbCritical, LoadResString(1905)
        MsgBox "データベース接続エラーです。しばらくたってからもう一度お試してください。", vbCritical, gTit
        End
    End If

'    Set oPassCheck = New clsPassCheck
'    lRtn = oPassCheck.PasswordCheck
'    Set oPassCheck = Nothing
'    If lRtn <= 0 Then End

    '---------------------------------------------------------------------------
    'windows login userを取得する
    '---------------------------------------------------------------------------
    sUserId = gf_GetLoginUser

''''-------------------------------------------------------
''''tbCpfSystemUserでsiUserLevel列がないので不明
''''-------------------------------------------------------
''''sSQL = " SELECT TOP 1 iUserID , siUserLevel , vPassword"         ''''2021.10.12 del jhi
    sSQL = " SELECT TOP 1 iUserID , 1 as siUserLevel , vPassword"    ''''2021.10.12 add jhi

    sSQL = sSQL & " FROM tbCpfSystemUser "
    sSQL = sSQL & " WHERE vLoginID = '" & sUserId & "' "

    Set oRs = g_obj_Conn.Execute(sSQL)

    If oRs.EOF Then
        MsgBox "現在のログインユーザに対して使用許可がされていません。", vbExclamation Or vbOKOnly, "ユーザ認証エラー"
        oRs.Close
        Set oRs = Nothing
        g_obj_Conn.Close
        End
    End If

'    glUserID = lRtn
    glUserID = oRs.Fields(0)
    glUserLevel = oRs.Fields(1)
    sPWD = oRs.Fields(2)

    oRs.Close
    Set oRs = Nothing

    If sPWD <> "" Then

Dim oGao As Object
Dim sKey As String

        Set oGao = CreateObject("GaoEncode.GaoeAPI")

        sKey = GetSetting("Nyushi", "Settings", "UPWD", "UPWD")
        gsUserPwd = Replace(oGao.DecodeStr(sPWD, sKey, 0), vbCrLf, "")

        Set oGao = Nothing

    Else
        gsUserPwd = ""
    End If

    ' starting point of the application
    gstrimgpath = App.Path + "\images\"
    NEWICON = gstrimgpath + "New.ico"
    CLEARICON = gstrimgpath + "Refresh.ico"
    CANCELICON = gstrimgpath + "Cancel.ico"
    DELETEICON = gstrimgpath + "Delete.ico"
    SAVEICON = gstrimgpath + "Save.ico"

''''SEARCHICON = gstrimgpath + "Query.ico"   ''''2022.01.06 del jhi
    SEARCHICON = gstrimgpath + "Search.ico"  ''''2022.01.06 add jhi

    frmSplash.Show                  ' display the splash screen
    frmSplash.Refresh

    Set fMainForm = New frmMain     ' create instance of the MDI form
    Load fMainForm
    Unload frmSplash

    fMainForm.Show

    Exit Sub

ErrorHandler:
    MsgBox "起動時の初期化に失敗しました。" & vbCrLf & Err.Description, vbExclamation Or vbOKOnly, "起動失敗"
    End

End Sub


'*******************************************************************************
'* DBを接続処理 繋げる                                                         *
'*-----------------------------------------------------------------------------*
'* 2022.02.01 update jhi                                                       *
'*******************************************************************************
Public Function g_void_OpenConnection() As Boolean

    On Error Resume Next

    'function which opens the database connection
    Dim l_str_Machine     As String
    Dim l_str_Database    As String
    Dim l_str_User        As String
    Dim l_str_Pwd         As String


    Set g_obj_Conn = New ADODB.Connection

    'Registory [HKEY_CURRENT_USER]-[Software]-[VB and VBA Program]-[Nyushi]-[Setting]
    'から値を取得する


''''----------------------------------------------------------------------------
''''条件付きコンパイル引数の設定 2022.02.01 add jhi
''''----------------------------------------------------------------------------
#If zengo_kubun = 1 Then


    l_str_Machine = GetSetting("Nyushi_zenki", "Settings", "MachineName", "")

    '---------------------------------------------------------------------------
    ' 前期用DBに繋げる 2022.01.16 update jhi 再確認
    '---------------------------------------------------------------------------
''''l_str_Database = GetSetting("Nyushi_zenki", "Settings", "DatabaseName", "") ''''2022.01.02 del jhi
    l_str_Database = "STE0100"                                                  ''''2022.01.02 add jhi

    l_str_User = GetSetting("Nyushi_zenki", "Settings", "DatabaseUser", "")
    l_str_Pwd = GetSetting("Nyushi_zenki", "Settings", "DatabasePassword", "")

    ' check for all database connection parameters
    If Trim(l_str_User) = "" Or Trim(l_str_Database) = "" Or Trim(l_str_Machine) = "" Then
        g_void_OpenConnection = False
        Exit Function
    End If

'    l_str_ConnStr = "Provider =SQLOLEDB;Server=" & l_str_Machine & ";UID=" & l_str_User
'    l_str_ConnStr = l_str_ConnStr & ";PWD=" & l_str_Pwd & ";Database=" & l_str_Database
    l_str_ConnStr = ";DSN=" & l_str_Machine & ";UID=" & l_str_User & ";PWD=" & l_str_Pwd & ";Database=" & l_str_Database

#Else

    l_str_Machine = GetSetting("Nyushi_goki", "Settings", "MachineName", "")

    '---------------------------------------------------------------------------
    ' 前期用DBに繋げる 2022.01.16 update jhi 再確認
    '---------------------------------------------------------------------------
''''l_str_Database = GetSetting("Nyushi_goki", "Settings", "DatabaseName", "")      ''''2022.01.02 del jhi
    l_str_Database = "STE0100_goki"                                                 ''''2022.01.02 add jhi

    l_str_User = GetSetting("Nyushi_goki", "Settings", "DatabaseUser", "")
    l_str_Pwd = GetSetting("Nyushi_goki", "Settings", "DatabasePassword", "")

    ' check for all database connection parameters
    If Trim(l_str_User) = "" Or Trim(l_str_Database) = "" Or Trim(l_str_Machine) = "" Then
        g_void_OpenConnection = False
        Exit Function
    End If

'    l_str_ConnStr = "Provider =SQLOLEDB;Server=" & l_str_Machine & ";UID=" & l_str_User
'    l_str_ConnStr = l_str_ConnStr & ";PWD=" & l_str_Pwd & ";Database=" & l_str_Database
    l_str_ConnStr = ";DSN=" & l_str_Machine & ";UID=" & l_str_User & ";PWD=" & l_str_Pwd & ";Database=" & l_str_Database


#End If


    With g_obj_Conn
        .CursorLocation = adUseClient
        .Open l_str_ConnStr
    End With

    If Err.Number <> 0 Then
        g_void_OpenConnection = False
    Else
        g_void_OpenConnection = True
    End If
   

End Function

Public Function RetrieveRecords(ByVal StrSelect As String) As ADODB.Recordset

    ' this function executes the SQL Query and returns the diconnected recordset
    On Error GoTo ErrorHandler
    
    Dim RS As ADODB.Recordset
    Set RS = CreateObject("ADODB.Recordset")

    With RS
        .Open StrSelect, g_obj_Conn, adOpenForwardOnly, adLockOptimistic
    End With
    
    Set RS.ActiveConnection = Nothing
    Set RetrieveRecords = RS
    
    Set RS = Nothing

    Exit Function

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
    Set RetrieveRecords = Nothing

End Function

' load all the resource strings of the form as a whole
Sub LoadResStrings(frm As Form)

    On Error Resume Next

    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Integer

    'set the form's caption
    If Len(frm.Tag) > 0 Then
        If IsNumeric(frm.Tag) Then
            If CInt(frm.Tag) = 1905 Then
                frm.Caption = gTit                         ''''前期 - 入試システム or 後期 - 入試システム
             Else
                frm.Caption = LoadResString(CInt(frm.Tag)) ''''1905 - 前期 - 入試システム
             End If
        End If
    End If

    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In frm.Controls
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "CommandButton" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "OptionButton" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = LoadResString(CInt(obj.Tag))
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = LoadResString(CInt(obj.Tag))
            Next
        Else
            nVal = 0
            nVal = Val(ctl.Tag)
            If nVal > 0 Then ctl.Caption = LoadResString(nVal)
            nVal = 0
            nVal = Val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = LoadResString(nVal)
        End If
    Next
    
End Sub

'*******************************************************************************
' this functions forms the search query, based on the user inputs,
' and then passes this query to the retieve records function to execute it
'*******************************************************************************
Public Sub SearchRecords(Optional NewRow As Boolean)

    On Error GoTo ErrorHandler

    Dim StrWhereClause  As String
    Dim strTableName    As String
    Dim StrSelectString As String
    Dim StrLabelValue   As String
    Dim intPos          As Integer
    Dim objcls          As FieldDetail
    Dim objCombo        As ComboDetail ' added by team
    Dim lngRow          As Long
    
    Dim ctl             As Control
    Dim sCtlType        As String
    Dim l_str_Sql       As String
    Dim l_obj_Rst       As New ADODB.Recordset
    
    ' form the where clause first
    ' go through each and every control in the form
    ' The fields displaying database field values will
    ' have the field name specified int he tag property of the control
    ' usingt he tag value and the value in the control form the where crietia string
    
    If NewRow <> True Then
        If CheckDirty = False Then
            Exit Sub
        End If
    End If
 
   
    For Each ctl In fMainForm.ActiveForm.Controls
        With ctl
            If .Tag <> "" Then
                sCtlType = TypeName(ctl)
                Select Case sCtlType
                    ' right now code only for textbox is taken into consideration
                    ' code needs to be wriiten for other types of control types also
                    Case "TextBox"  ' changed on 281101
                        If Len(Trim$(.Text)) <> 0 Then       'control has value entered
                            If Len(StrWhereClause) <> 0 Then
                                If UCase(.Tag) = "[IZIPCODEID]" Then
                                    ' specific check for the ExamineeProfile form
                                    ' to display the zip code , instead of the zipcode id
                                    l_str_Sql = "SELECT iZipCodeId FROM tbSTEZipCodeMaster"
                                    l_str_Sql = l_str_Sql & " WHERE vZipCodeName='" & Trim(.Text) & "'"
                                    l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
                                    If Not l_obj_Rst.EOF Then
'                                        StrWhereClause = StrWhereClause & "  and " & UCase(.Tag) & " like '%" & UCase(l_obj_Rst("iZipCodeId")) & "%'"
                                        StrWhereClause = StrWhereClause & "  and " & UCase(.Tag) & " like '" & UCase(l_obj_Rst("iZipCodeId")) & "%'"
                                    End If
                                    l_obj_Rst.Close
                                    Set l_obj_Rst = Nothing
                                ElseIf UCase(.Tag) = "[IHIGHSCHOOLID]" Then
                                    ' specific check for the ExamineeProfile form
                                    ' to display the highschool code, instead of the highschool id
                                    l_str_Sql = "SELECT iHighSchoolId FROM tbSTEHighSchoolType"
                                    l_str_Sql = l_str_Sql & " WHERE vHighSchoolCode='" & Trim(.Text) & "'"
                                    l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
                                    If Not l_obj_Rst.EOF Then
'                                        StrWhereClause = StrWhereClause & "  and " & UCase(.Tag) & " like '%" & UCase(l_obj_Rst("iHighSchoolId")) & "%'"
                                        StrWhereClause = StrWhereClause & "  and " & UCase(.Tag) & " like '" & UCase(l_obj_Rst("iHighSchoolId")) & "%'"
                                    End If
                                    l_obj_Rst.Close
                                    Set l_obj_Rst = Nothing
'2002/12/17 frmExamineeProfile画面で受験番号での検索時％検索しないように修正
'2003/02/12 frmExamineeCheck画面も追加
                                ElseIf UCase(.Tag) = "[IJUKENNUMBER]" Then
                                    If fMainForm.ActiveForm.Name = "frmExamineeProfile" Or fMainForm.ActiveForm.Name = prvcExamCheck Then
                                        StrWhereClause = StrWhereClause & "  and " & UCase(.Tag) & " like '" & UCase(.Text) & "'"
                                    Else
                                        StrWhereClause = StrWhereClause & "  and " & UCase(.Tag) & " like '" & UCase(.Text) & "%'"
                                    End If
                                Else
'                                    StrWhereClause = StrWhereClause & "  and " & UCase(.Tag) & " like '%" & UCase(.Text) & "%'"
                                    StrWhereClause = StrWhereClause & "  and " & UCase(.Tag) & " like '" & UCase(.Text) & "%'"
                                End If
                            Else
'2002/12/17 frmExamineeProfile画面で受験番号での検索時％検索しないように修正
'2003/02/12 frmExamineeCheck画面も追加
                                If UCase(.Tag) = "[IJUKENNUMBER]" Then
                                    If fMainForm.ActiveForm.Name = "frmExamineeProfile" Or fMainForm.ActiveForm.Name = prvcExamCheck Then
                                        StrWhereClause = UCase(.Tag) & " like '" & UCase(.Text) & "'"
                                    Else
'                                        StrWhereClause = .Tag & " like '%" & .Text & "%'"
                                        StrWhereClause = .Tag & " like '" & .Text & "%'"
                                    End If
                                Else
'                                    StrWhereClause = .Tag & " like '%" & .Text & "%'"
                                    StrWhereClause = .Tag & " like '" & .Text & "%'"
                                End If
                            End If
                        End If
                        
                    Case "ComboBox"
                        If Len(Trim$(.Text)) <> 0 Then       'control has value entered
                            If Len(StrWhereClause) <> 0 Then
                                ' added by dileep
                                For lngRow = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
'入試実施時の不具合No1対応  2004/01/24
'                                    If Trim$(fMainForm.ActiveForm.m_ComboDetails.Item(lngRow).Description) = Trim$(.Text) Then
                                    If Trim$(fMainForm.ActiveForm.m_ComboDetails.Item(lngRow).Description) = Trim$(.Text) And UCase(Trim$(fMainForm.ActiveForm.m_ComboDetails.Item(lngRow).GroupId)) = UCase(Trim$(.Tag)) Then
                                        ' matching record found in collection, find out its group
                                        'another for loop on collection, find out entry with group,and description
                                        'pick up the id
'                                        intGroupID = fMainForm.ActiveForm.m_ComboDetails.Item(lngRow).GroupId
                                        
                                        StrWhereClause = StrWhereClause & "  and " & .Tag & " = '" & Trim(fMainForm.ActiveForm.m_ComboDetails.Item(lngRow).Value) & "'"
                                        Exit For
                                    End If
                                Next
                            Else
                                ' added by dileep
                                For lngRow = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
'入試実施時の不具合No1対応  2004/01/24
'                                    If Trim$(fMainForm.ActiveForm.m_ComboDetails.Item(lngRow).Description) = Trim$(.Text) Then
                                    If Trim$(fMainForm.ActiveForm.m_ComboDetails.Item(lngRow).Description) = Trim$(.Text) And UCase(Trim$(fMainForm.ActiveForm.m_ComboDetails.Item(lngRow).GroupId)) = UCase(Trim$(.Tag)) Then
                                        StrWhereClause = .Tag & " = '" & Trim(fMainForm.ActiveForm.m_ComboDetails.Item(lngRow).Value) & "'"
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                        
                    Case "CheckBox"
                        If Len(StrWhereClause) <> 0 Then
                            StrWhereClause = StrWhereClause & "  and " & .Tag & " = " & .Value
                        Else
                            StrWhereClause = .Tag & " = " & .Value
                        End If
                        
                End Select
            End If
        End With
    Next
    
    ' in every child form there is a collection named m_colFieldDetails
    ' this collection holds the display value and database field value for the tabel
    ' the variable m_TableName holds the name of the table corrsponding to the collection
    
    strTableName = fMainForm.ActiveForm.m_TableName
    
    StrSelectString = ""
    For Each objcls In fMainForm.ActiveForm.m_colFieldDetails
        StrSelectString = StrSelectString & "," & objcls.DBReadFieldName & " as " & objcls.SCRFieldName
    Next

    If Len(Trim$(StrSelectString)) <> 0 And Len(Trim$(strTableName)) <> 0 Then
        If fMainForm.ActiveForm.Name = prvcExamCheck Then
            StrSelectString = "Select Top " & gsExamIDTo & " '" & LoadResString(2201) & "'" & StrSelectString & " from " & strTableName
        Else
            StrSelectString = "Select '" & LoadResString(2201) & "'" & StrSelectString & " from " & strTableName
        End If
    Else
        fMainForm.ActiveForm.lblErrorMsg.Caption = "フォームが正しく設定されていません。"     ''''LoadResString(1119)
        fMainForm.ActiveForm.lblErrorMsg.Visible = True
    End If
    
    If NewRow = True Then
        StrSelectString = StrSelectString & " where 1 <> 1" ''''2022.01.05 del jhi
''''    StrSelectString = StrSelectString & " where 1 = 1"  ''''2022.01.05 add jhi
        Call FillGrid(StrSelectString, True)
    Else
        If fMainForm.ActiveForm.Name = prvcExamCheck Then
Dim lsWhere As String
            lsWhere = " iNendo = " & gsExamCheckNendo
            lsWhere = lsWhere & " and iJukenNumber >= " & gsExamIDFrom & " "
            If Len(StrWhereClause) <> 0 Then
                StrSelectString = StrSelectString & " where " & StrWhereClause & " and " & lsWhere
            Else
                StrSelectString = StrSelectString & " where " & lsWhere
            End If
'入試試験実施時不具合No2対応
            ' 2005/01/16 受験番号順にソートする。　寺尾
            StrSelectString = StrSelectString & " order by  iJukenNumber "
        Else
            If Len(StrWhereClause) <> 0 Then
                StrSelectString = StrSelectString & " where " & StrWhereClause
            End If
        End If
        FillGrid StrSelectString
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

    ' this function initializes the grid in all the master maintenance screens
Public Sub InitializeSearchGrid()

    On Error GoTo ErrorHandler

    With fMainForm.ActiveForm.hfgSearchGrid
        .Visible = False
        .BackColor = &HFFFFFF
        .BackColorBkg = &HFFFFFF
        .BackColorFixed = &H8000000F
        .BackColorSel = &H800000
        .FixedCols = 0
        .FontFixed.Bold = False
        .Font.Bold = False
        .ForeColorFixed = &H80000008
        .ForeColor = &H800000
        '.CellTextStyle = "0"
        .GridLinesFixed = flexGridInset
        .GridColor = &H808080
        .AllowUserResizing = flexResizeColumns
        .Visible = True
    End With

    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

Public Sub FillGrid(ByVal StrSelect As String, Optional NewRow As Boolean)

    ' function which populates the grid with the values returned from the database query
    On Error GoTo ErrorHandler

    Dim RS As ADODB.Recordset
    Dim lngRow As Long
    Dim lngCol As Long
    
    Set RS = RetrieveRecords(StrSelect)

    If NewRow Then RS.AddNew

    If Not RS.EOF Then
        With fMainForm.ActiveForm.hfgSearchGrid
            .Visible = False
            Set .DataSource = RS
            fMainForm.ActiveForm.f_int_PrevRow = 1
            .Col = 0
            If Not NewRow Then
                For lngRow = 0 To .Rows - 1
                    .Row = lngRow
                    .CellFontUnderline = True
                    .CellForeColor = QBColor(4)
                Next
            End If

            .ColWidth(0) = 800
            RS.MoveFirst
            
            ' set the width of individual columns
            For lngCol = 1 To .cols - 1
                'New Code added to make width of combo type columns in grid zero (Mahesh) 20/5/2002
                If UCase(fMainForm.ActiveForm.m_colFieldDetails.Item(lngCol).strDataType) = "COMBO" Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = fMainForm.ActiveForm.m_colFieldDetails.Item(lngCol).ColWidth
                End If
                'New Code Ends
            Next
            .Visible = True

            If NewRow Then
                NewData
            Else
                'set mode to query
                fMainForm.ActiveForm.m_bMode = "QUERY"
                'enable delete option
                fMainForm.mnuToolsDelete.Enabled = True
                
                fMainForm.Toolbar1.Buttons("Delete").Enabled = True
            End If
                
        End With
        Call ComboConversion    ' call the function to display combo description instead of value
        
    Else
        SearchRecords True  ' to clear the grid incase no records are found in the search
        fMainForm.ActiveForm.lblErrorMsg.Visible = True
''''    fMainForm.ActiveForm.lblErrorMsg.Caption = LoadResString(1120) 'レコードを表示できません
        fMainForm.ActiveForm.lblErrorMsg.Caption = LoadResString(1964) 'レコードがありませんでした。
    End If

    Set RS = Nothing
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

Public Sub AssignValues(ByVal bGridToForm As Boolean, Optional bNoDirtyCheck As Boolean)

    On Error GoTo ErrorHandler

    Dim retval As Integer
    Dim ctl As Control
    Dim sCtlType As String
    Dim objCombo As ComboDetail
    Dim lngRow As Long
    
    Dim l_str_Cap As String 'Master maint caption (5.4.2002 Mahesh)
    Dim l_int_position  As Integer 'Position from which the caption to be changed
    'New Code 20/5/02 Mahesh
    Dim l_int_Counter As Long
    Dim i_int_lngCol As Integer 'ctr to loop thru entire m_colFieldDetails collection to search the groupid and foreign key ID
    'New Code ends
    ' populate the value of the grid to the textboxes
    ' go through each and every control in the form
    ' The fields displaying database field values will
    ' have the field name specified in the tag property of the control
    ' using the tag value
    ' bGridToForm = true means from gris to form else form from to grid

    If bNoDirtyCheck = False Then
        If CheckDirty = False Then Exit Sub
    End If
    
    fMainForm.ActiveForm.m_bChangeOn = True     ' this var used to set dirty flag
    With fMainForm.ActiveForm.hfgSearchGrid
        fMainForm.ActiveForm.m_lngCurrentRow = .Row
        For Each ctl In fMainForm.ActiveForm.Controls
            If ctl.Tag <> "" Then
                sCtlType = TypeName(ctl)
                Select Case sCtlType
                
                    Case "TextBox"
                        If bGridToForm Then
                            If IsNull(Trim(.TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos))) Or Trim(.TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos)) = "" Then
                                ctl.Text = ""
                            Else
'                                If ctl.Name = "txtdtBirthDay" Or ctl.Name = "txtSecondDayExam" Then
'                                    ctl.Text = Format(Trim(.TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos)), "gggee年mm月dd日")       ' for english, use this and comment the above one
                                If fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).strDataFormat <> "" Then
                                    ctl.Text = Trim(Format(.TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos), fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).strDataFormat))             ' for english, use this and comment the above one
                                Else
                                    ctl.Text = Trim(.TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos))
                                End If
                            End If
                        Else
'                            If ctl.Name = "txtdtBirthDay" Or ctl.Name = "txtSecondDayExam" Then
'                                .TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos) = Format(ctl.Text, "gggee年mm月dd日")             ' for english, use this and comment the above one
                            If fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).strDataFormat <> "" Then
                                .TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos) = Format(ctl.Text, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).strDataFormat)             ' for english, use this and comment the above one
                            Else
                                .TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos) = ctl.Text
                            End If
                        End If
                        
                    Case "ComboBox"
                        If bGridToForm Then
                            If IsNull(Trim(.TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos))) Or Trim(.TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos)) = "" Then
                                ctl.ListIndex = 0
                            Else
                                'New Code Start 20/5/2002 Mahesh
                                        For l_int_Counter = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
                                            'compare group id(which is also db field name) with fMainForm.ActiveForm.m_colFieldDetails.DbFieldName field name
                                            If Trim(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Value) = Trim(.TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos)) And UCase(Trim(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).GroupId)) = UCase(ctl.Tag) Then      'UCase(fMainForm.ActiveForm.m_colFieldDetails.Item(i_int_lngCol).DBFieldName) Then
                                                If fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Description <> "" Then
                                                    ctl.Text = fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Description
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                'New Code Ends
                            'Commented on 20/5/02
                            End If
                        Else
                            ' code added on 06/06/02 to display id values in the grid instead of the description
                            For l_int_Counter = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
'入試実施時の不具合No1対応  2004/01/24
'                                If Trim$(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Description) = Trim(ctl.Text) Then
                                If Trim$(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Description) = Trim(ctl.Text) And UCase(Trim(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).GroupId)) = UCase(Trim(ctl.Tag)) Then
                                    .TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos) = Trim(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Value)
                                    Exit For
                                End If
                            Next
                        End If
                        
                    Case "CheckBox"
                        If bGridToForm Then
                            If Len(.TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos)) = 0 Then
                                ctl.Value = 0
                            Else
                                ctl.Value = .TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos)
                            End If
                        Else
                            .TextMatrix(.Row, fMainForm.ActiveForm.m_colFieldDetails.Item(ctl.Tag).GridColPos) = ctl.Value
                        End If
                        
                End Select
            End If
        Next
    End With
    fMainForm.ActiveForm.m_bChangeOn = False
    fMainForm.ActiveForm.m_bMode = "QUERY"      ' added on 07/11/01 by dileep
    If bGridToForm Then ' added by mahesh
        fMainForm.mnuToolsDelete.Enabled = True
        fMainForm.Toolbar1.Buttons("Delete").Enabled = True
    End If
    'New Code 28/3/2002
    l_str_Cap = fMainForm.ActiveForm.Caption
    l_int_position = InStr(1, l_str_Cap, "_")
    If l_int_position > 0 Then
        l_str_Cap = Mid(l_str_Cap, 1, l_int_position - 1)
    End If
    l_str_Cap = l_str_Cap & "_" & LoadResString(2465)  '"_Edit"
    fMainForm.ActiveForm.Caption = l_str_Cap
    'New Code
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Sub

'-------------------------------------------------------------------------------
' this function checks whether the user has made
' any changes to the form after it has gone into the query mode
'-------------------------------------------------------------------------------
Public Function CheckDirty() As Boolean

    On Error GoTo ErrorHandler
    Dim retval As Integer
    
    CheckDirty = True

    If fMainForm.ActiveForm.m_bDirty = True Then

        'prompt to save the changes
''''    retval = MsgBox(LoadResString(1118), vbYesNoCancel)
''''    retval = MsgBox("変更を保存しますか？", vbYesNoCancel) ''''2022.01.05 del jhi

        retval = vbNo ''''2022.01.05 add jhi

        Select Case retval
            Case vbCancel
                ClearErrorMsg

                CheckDirty = False
                Exit Function
            Case vbYes
                ' validate the control values
                If ValidateAndSaveData() = False Then
                    If fMainForm.ActiveForm.m_lngCurrentRow > 0 Then
                        fMainForm.ActiveForm.hfgSearchGrid.Row = fMainForm.ActiveForm.m_lngCurrentRow
                    End If
                    CheckDirty = False
                    Exit Function
                End If
            Case vbNo
                ' proceed with populating the new row values
                ClearErrorMsg
                CheckDirty = True
                fMainForm.ActiveForm.m_bDirty = False
                fMainForm.mnuToolsSave.Enabled = False
                fMainForm.mnuToolsCancel.Enabled = False
                
                fMainForm.Toolbar1.Buttons("Save").Enabled = False
                fMainForm.Toolbar1.Buttons("Cancel").Enabled = False
      
        End Select
        
    End If

    Exit Function

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Function

Public Function SaveData() As Boolean
On Error GoTo ErrorHandler
' if fails then return false else true
    SaveData = True
    Dim strSaveString As String
    Dim strSaveString1 As String
    Dim strSaveString2 As String
    Dim strTableName As String
    Dim strValue As String
    Dim strFieldName As String
    Dim strMode As String
    Dim bPK As Boolean
    Dim lngRow As Long
    Dim l_int_Counter As Integer
    strTableName = fMainForm.ActiveForm.m_TableName
    strSaveString = ""
    
    'for inseration from the insert string - Insert tablename (
    ' and the insert value string values(
    '
    ' for update - update tablename
    ' set statement
    ' where primary key
    
    strSaveString1 = ""
    strSaveString2 = ""
    
    strMode = fMainForm.ActiveForm.m_bMode
    'go throught the collection populated and form the string
    With fMainForm.ActiveForm.m_colFieldDetails
        For lngRow = 1 To .Count
            strValue = .Item(lngRow).strValue
            strFieldName = .Item(lngRow).DBFieldName
            bPK = .Item(lngRow).PrimaryKey
            Select Case UCase(.Item(lngRow).strDataType)
                Case "STRING"
                    'concatenate the value between single quotes
                     strValue = "'" & strValue & "'"
                     
                Case "DATE"
                    'concatenate the value between single quotes
                     strValue = "'" & Format(strValue, "MM/DD/YYYY") & "'"
                
                Case "INTEGER", "LONG"
                    If Trim(strValue) = "" Then
                        strValue = 0
                    Else
                        strValue = strValue
                    End If
                    
                Case "COMBO"
                    If strValue = "" Or IsNull(strValue) Then
                        strValue = "NULL"
                    Else
                        For l_int_Counter = 1 To fMainForm.ActiveForm.m_ComboDetails.Count
'入試実施時の不具合No1対応  2004/01/24
                            If UCase(Trim(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).GroupId)) = UCase(Trim(strFieldName)) Then
                                If Trim$(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Description) = Trim(strValue) Then
                                    strValue = Trim(fMainForm.ActiveForm.m_ComboDetails.Item(l_int_Counter).Value)
                                    Exit For
                                End If
                            End If
                        Next
                    End If
            End Select
            'check the mode fMainForm.ActiveForm.m_bMode
            Select Case strMode
                Case "NEW"
                    'need to form the insert string over here
                    If Len(strSaveString1) = 0 Then
                        strSaveString1 = "Insert " & strTableName & " (" & strFieldName
                        strSaveString2 = " values (" & strValue
                    Else
                        strSaveString1 = strSaveString1 & "," & strFieldName
                        strSaveString2 = strSaveString2 & "," & strValue
                    End If
                    If lngRow = .Count Then 'last item in the collection
                        strSaveString1 = strSaveString1 & ")"
                        strSaveString2 = strSaveString2 & ")"
                    End If
                Case "QUERY"
                        'need to form the update string over here
                    If bPK = True Then
                        If Len(strSaveString2) = 0 Then
                            strSaveString2 = " where " & strFieldName & " = " & strValue
                        Else
                            strSaveString2 = strSaveString2 & " and " & strFieldName & " = " & strValue
                        End If
                    Else
                        If Len(strSaveString1) = 0 Then
                            strSaveString1 = "Update " & strTableName & " set " & strFieldName & " = " & strValue
                        Else
                            strSaveString1 = strSaveString1 & ", " & strFieldName & " = " & strValue
                        End If
                    End If
                Case "DELETE"
                        'need to form the update string over here
                    strSaveString1 = "Delete " & strTableName & " "
                    If bPK = True Then
                        If Len(strSaveString2) = 0 Then
                            strSaveString2 = " where " & strFieldName & " = " & strValue
                        Else
                            strSaveString2 = strSaveString2 & " and " & strFieldName & " = " & strValue
                        End If
                    End If
            End Select
        Next
    End With
    
    If Len(strSaveString1) <> 0 And Len(strSaveString2) <> 0 Then
        g_obj_Conn.Execute strSaveString1 & strSaveString2

        fMainForm.ActiveForm.lblErrorMsg.Visible = True
        fMainForm.ActiveForm.lblErrorMsg.Caption = "データを正常に保存しました。" ''''LoadResString(1121)

        If strMode = "NEW" Then
            'need to add this as a new row in the grid
            With fMainForm.ActiveForm.hfgSearchGrid
                If .TextMatrix(1, 0) <> "" Then .Rows = .Rows + 1
                .Row = .Rows - 1
                AssignValues False, True
                .Col = 0                                    ' 04/10/01
                .TextMatrix(.Row, 0) = LoadResString(2201)
                .CellFontUnderline = True
                .CellForeColor = QBColor(4)
            End With
            'set mode to query
            fMainForm.ActiveForm.m_bMode = "QUERY"
            'enable delete option
            fMainForm.mnuToolsDelete.Enabled = True
            fMainForm.Toolbar1.Buttons("Delete").Enabled = True
            
        ElseIf strMode = "QUERY" Then
            ' need to update the changes in the grid
            lngRow = fMainForm.ActiveForm.hfgSearchGrid.Row
            fMainForm.ActiveForm.hfgSearchGrid.Row = fMainForm.ActiveForm.m_lngCurrentRow
            AssignValues False, True
            If lngRow <> fMainForm.ActiveForm.m_lngCurrentRow Then
                fMainForm.ActiveForm.hfgSearchGrid.Row = lngRow
                AssignValues True, True
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    If Err.Number = -2147217873 Then
        ' constraint violation which checks which ensures that the
        ' juken number will be unique for a given Nendo

        fMainForm.ActiveForm.lblErrorMsg.Visible = True
        fMainForm.ActiveForm.lblErrorMsg.Caption = LoadResString(2445)
        SaveData = False
    Else
        MsgBox Err.Description, vbInformation, LoadResString(1729)
    End If
End Function

Public Function PopulateCollection()
    ' this function populates the strValue field of the field details collection
    ' with the values to be updated in the database
    Dim ctl As Control
    Dim sCtlType As String
    Dim strValue As String
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    
    On Error GoTo ErrorHandler
        
    For Each ctl In fMainForm.ActiveForm.Controls
        strValue = ""
        With ctl
            If .Tag <> "" Then
                sCtlType = TypeName(ctl)
                Select Case sCtlType
                    'right now code only for textbox is taken into consideration
                    'code needs to be wriiten for other types of control types also
                    Case "TextBox"  ' changed on 281101
                        
                        If UCase(.Tag) = "[IZIPCODEID]" And (UCase(fMainForm.ActiveForm.Name) = "FRMEXAMINEEPROFILE" Or UCase(fMainForm.ActiveForm.Name) = "FRMHIGHSCHOOLTYPE") Then
                            ' specific condition for the examinee profile table to take the value of zipcode field
                            l_str_Sql = "SELECT iZipCodeId FROM tbSTEZipCodeMaster"
                            l_str_Sql = l_str_Sql & " WHERE vZipCodeName='" & Trim(.Text) & "'"
                            l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
                            If Not l_obj_Rst.EOF Then
                                strValue = l_obj_Rst("iZipCodeId")
                            End If
                            l_obj_Rst.Close
                            Set l_obj_Rst = Nothing
                        ElseIf UCase(.Tag) = "[IHIGHSCHOOLID]" And UCase(fMainForm.ActiveForm.Name) = "FRMEXAMINEEPROFILE" Then
                            ' specific condition for the examinee profile table to take the value of highschool field
                            l_str_Sql = "SELECT iHighSchoolId FROM tbSTEHighSchoolType"
                            l_str_Sql = l_str_Sql & " WHERE vHighSchoolCode='" & Trim(.Text) & "'"
                            l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
                            If Not l_obj_Rst.EOF Then
                                strValue = l_obj_Rst("iHighSchoolId")
                            End If
                            l_obj_Rst.Close
                            Set l_obj_Rst = Nothing
                        Else
                            strValue = Trim(.Text)
                        End If
                        'the corresponding field from the collection
                        'the tag will have the database field name an the corresponding
                        ' field from the colletion can be retrived using the tage value
                        fMainForm.ActiveForm.m_colFieldDetails.Item(.Tag).strValue = strValue
                        
                    Case "ComboBox"
                        strValue = Trim(.Text)
                        'the corresponding field from the collection
                        'the tag will have the database field name an the corresponding
                        ' field from the colletion can be retrived using the tage value
                        
                        fMainForm.ActiveForm.m_colFieldDetails.Item(.Tag).strValue = strValue
                        
                    Case "CheckBox"
                        strValue = .Value
                        'the corresponding field from the collection
                        'the tag will have the database field name an the corresponding
                        ' field from the colletion can be retrived using the tage value
                        
                        fMainForm.ActiveForm.m_colFieldDetails.Item(.Tag).strValue = strValue
                End Select
            End If
        End With
    Next
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Function

Public Function ValidateControl() As Boolean
    ' this function validates the controls on the form
    ' checks whether all required fields are filled
    On Error GoTo ErrorHandler
    ' if goes through then validatecontrols is true if fails then false
    ' assumption that all the values will be entered
    Dim strValue As String
    Dim lngRow As Long
    
    ValidateControl = True
    ClearErrorMsg
    PopulateCollection
    
    With fMainForm.ActiveForm.m_colFieldDetails
        For lngRow = 1 To .Count
            .Item(lngRow).strErrorMsg = ""
            'incase of any error populate the strErrorMsg
            'of the corrsponding item in the collection
            If .Item(lngRow).bMandatory Then ' check for mandatory
                If Len(Trim$(.Item(lngRow).strValue)) = 0 Then
                    .Item(lngRow).strErrorMsg = Mid(.Item(lngRow).SCRFieldName, 2, Len(.Item(lngRow).SCRFieldName) - 2) & " " & LoadResString(1117)
                    ValidateControl = False
                End If
            End If
            'check for correctness of the datatype
            Select Case UCase(.Item(lngRow).strDataType)
                Case "STRING"
                        ' no special  validation for string
                        'this is just to demonstrate where to put the
                        'datatype check
                        
            End Select
        Next
    End With
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Public Function ClearErrorMsg()
    ' function which clears the displayed error messages
    Dim lngRow As Long
    With fMainForm.ActiveForm
        .lblErrorMsg.Caption = ""
        For lngRow = 1 To .m_colFieldDetails.Count
            .lblErrIndicator(lngRow).Visible = False
            .m_colFieldDetails.Item(lngRow).strErrorMsg = ""
        Next
        .lblErrorMsg.Visible = False
    End With
End Function

Public Function DisplayErrorMsg()
    ' to display the error messages
    On Error GoTo ErrorHandler
    Dim lngRow As Long
    Dim lnglines As Long
    Dim strMsg As String
    lnglines = 0
    
    With fMainForm.ActiveForm
        .lblErrorMsg.Caption = ""
        For lngRow = 1 To .m_colFieldDetails.Count
            .lblErrIndicator(lngRow).Visible = False
            strMsg = .m_colFieldDetails.Item(lngRow).strErrorMsg
            If Len(strMsg) <> 0 Then
                .lblErrorMsg.Caption = LoadResString(2436)
                .lblErrIndicator(lngRow).Visible = True
                lnglines = 1    ' general error message, with error indicator alongside all the fields which are not filled
            End If
        Next
        If Len(.lblErrorMsg.Caption) <> 0 Then
            .lblErrorMsg.Visible = True
        Else
            .lblErrorMsg.Visible = False
        End If
    End With
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Public Function ValidateAndSaveData() As Boolean
On Error GoTo ErrorHandler
    ValidateAndSaveData = True
    ' validate the control values
    If ValidateControl() = False Then
        DisplayErrorMsg
        ValidateAndSaveData = False
        Exit Function
    Else
        ' got for form specific extra validation
        If fMainForm.ActiveForm.ExtraValidation() = False Then
            DisplayErrorMsg
            ValidateAndSaveData = False
            Exit Function
        Else
            'if validation successful then save the changes
            'if savedata  successful then proceeed else exit sub
            If SaveData = False Then
                ValidateAndSaveData = False
                Exit Function
            Else
                ValidateAndSaveData = True
                fMainForm.ActiveForm.m_bDirty = False
                fMainForm.mnuToolsSave.Enabled = False
                fMainForm.mnuToolsCancel.Enabled = False
                
                fMainForm.Toolbar1.Buttons("Save").Enabled = False
                fMainForm.Toolbar1.Buttons("Cancel").Enabled = False
            End If
        End If
    End If
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Public Function DeleteData()

    On Error GoTo ErrorHandler

    Dim strMode As String
    Dim retval As Integer

    'prompt to delete the data
''''retval = MsgBox(LoadResString(1122), vbYesNo)                      ''''2022.01.05 del jhi
    retval = MsgBox("選択しましたレコードを削除しますか？", vbYesNo)   ''''2022.01.05 add jhi
 
   Select Case retval
        Case vbYes
            'form the delete statement
            'for delete - delete tableaname
            'where priamry key
            strMode = fMainForm.ActiveForm.m_bMode
            fMainForm.ActiveForm.m_bMode = "DELETE"
            PopulateCollection
            If SaveData = False Then

                fMainForm.ActiveForm.lblErrorMsg.Visible = True
                fMainForm.ActiveForm.lblErrorMsg.Caption = LoadResString(1123)
            Else
                'clear the control and populate the first record in the grid in the control
                With fMainForm.ActiveForm.hfgSearchGrid
                    'delete the row from the grid
                    .Row = fMainForm.ActiveForm.m_lngCurrentRow
                    'remove this row
                    fMainForm.ActiveForm.m_bDirty = False
                    fMainForm.ActiveForm.m_bMode = strMode
                    If .Rows > 2 Then
                        .RemoveItem fMainForm.ActiveForm.m_lngCurrentRow
                        ClearData
                        .Row = 1    'jump to the first row
                        AssignValues True
                    Else
                        NewData
                        SearchRecords True
                    End If
                    fMainForm.ActiveForm.m_bDirty = False
                End With
            End If
            Exit Function
        Case vbNo
            'nothing to be done
    End Select
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Public Function DuplicateData()
On Error GoTo ErrorHandler
        NewData True
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Public Function NewData(Optional bDuplicateData As Boolean)

    Dim l_int_RetVal As Long
    Dim l_str_FormName As String
    
    Dim l_str_Cap As String 'Master maint caption
    Dim l_int_position  As Integer 'Position from which the caption to be changed

    On Error GoTo ErrorHandler

'    fMainForm.ActiveForm.lblErrorMsg.Caption = ""
    'check if dirty - prompt for saving
    If CheckDirty = False Then Exit Function
    
    'now as all data is saved the controls need to be cleared
    If bDuplicateData = False Then
        ClearData
    Else
        ' let the data remain on the screen
    End If

    'no need to clear the grid. if this new row is saved then it will be
    'added to the end of the grid row
    
    fMainForm.ActiveForm.m_bMode = "NEW"
    fMainForm.ActiveForm.m_lngCurrentRow = -1   'AS THERE is no corresponding record in the grid
    fMainForm.mnuToolsDelete.Enabled = False
    
    fMainForm.Toolbar1.Buttons("Delete").Enabled = False

     l_int_RetVal = f_lng_CreateNewId()
     If l_int_RetVal <> -1 Then
         l_str_FormName = fMainForm.ActiveForm.Name

         Select Case l_str_FormName
         Case "frmHighSchoolType"
             fMainForm.ActiveForm.txtHighSchoolID.Text = l_int_RetVal

         Case "frmInterviewerProfile"
             fMainForm.ActiveForm.txtInterviewerProfileId.Text = l_int_RetVal

         Case "frmInterviewRoomProfile"
             fMainForm.ActiveForm.txtInterviewRoomProfileId.Text = l_int_RetVal

         Case "frmZipCode"
             fMainForm.ActiveForm.txtZipCodeId.Text = l_int_RetVal

         Case "frmSubjectProfile"
             fMainForm.ActiveForm.txtSubjectProfileId.Text = l_int_RetVal

         Case "frmRoomProfile"
             fMainForm.ActiveForm.txtRoomProfileId.Text = l_int_RetVal

         Case "frmExamineeProfile"
             fMainForm.ActiveForm.txtExamineeProfileID.Text = l_int_RetVal

         Case "frmSubjectQuestionProfile"
             fMainForm.ActiveForm.txtSubjectQuestionId.Text = l_int_RetVal

        'Added new case for new MM
         Case "frmInterviewGroupProfile"
             fMainForm.ActiveForm.txtInterviewGroupProfileId.Text = l_int_RetVal
         End Select
 
    Else
        fMainForm.ActiveForm.lblErrorMsg.Visible = True
        fMainForm.ActiveForm.lblErrorMsg.Caption = LoadResString(1116)
     End If

    'New Code 5/4/2002 to display current mode in Master Maint Forms
    l_str_Cap = fMainForm.ActiveForm.Caption
    l_int_position = InStr(1, l_str_Cap, "_")

    If l_int_position > 0 Then
        l_str_Cap = Mid(l_str_Cap, 1, l_int_position - 1)
    End If

    l_str_Cap = l_str_Cap & "_" & "新規" ''''LoadResString(1041)  '"_New"
    fMainForm.ActiveForm.Caption = l_str_Cap
    'New Code

    Exit Function

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)

End Function

Public Function ClearData()
    ' clear the data for fresh entry
    Dim icnt As Integer
    Dim sDBFieldName As String
    Dim ctl As Control
    Dim sCtlType As String
    
On Error GoTo ErrorHandler
    If CheckDirty = False Then Exit Function
    
    fMainForm.ActiveForm.m_bChangeOn = True     ' this var used to set dirty flag
    For Each ctl In fMainForm.ActiveForm.Controls
        With ctl
            If .Tag <> "" Then
                sCtlType = TypeName(ctl)
                Select Case sCtlType
                    'right now code only for textbox is taken into consideration
                    'code needs to be wriiten for other types of control types also
                    Case "TextBox"
                            .Text = ""
                    Case "ComboBox"
                        If .ListCount > 0 Then
                           .ListIndex = 0
                        End If
                    Case "CheckBox"
                        .Value = 0
                End Select
            End If
        End With
    Next
    fMainForm.ActiveForm.m_bChangeOn = False     ' this var used to set dirty flag
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Public Function CancelData()
    ' cancel the ongoing editing/data entry
    On Error GoTo ErrorHandler
    If fMainForm.ActiveForm.m_lngCurrentRow = -1 Then ' that means it is a new record
        ClearData
    Else
       AssignValues True, True
       'Newly added on 21/5/2002 Mahesh
       fMainForm.ActiveForm.m_bDirty = False
       'New code ends
    End If

    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Private Sub ComboConversion()
    ' function which converts the integer fields of the comboboxes to the respective description
    Dim lngRow As Integer
    Dim lngCol As Integer
    Dim l_int_Counter As Integer
    Dim l_str_Sql As String
    Dim l_obj_Rst As New ADODB.Recordset
    
    fMainForm.ActiveForm.hfgSearchGrid.Visible = False
    Screen.MousePointer = vbArrowHourglass
       
    With fMainForm.ActiveForm.m_colFieldDetails
        For lngCol = 1 To .Count
            ' loop through all the columns
            If UCase(.Item(lngCol).strDataType) = "COMBO" Then
            'Commented to stop combo conversion on 20/5/2002 Mahesh
            ElseIf UCase(.Item(lngCol).strDataType) = "INTEGER" And (UCase(fMainForm.ActiveForm.Name) = "FRMEXAMINEEPROFILE" Or UCase(fMainForm.ActiveForm.Name) = "FRMHIGHSCHOOLTYPE") Then ' changed on 281101
                For lngRow = 1 To fMainForm.ActiveForm.hfgSearchGrid.Rows - 1
                    fMainForm.ActiveForm.hfgSearchGrid.Row = lngRow
                    fMainForm.ActiveForm.hfgSearchGrid.Col = .Item(lngCol).GridColPos
                    
                    If UCase(.Item(lngCol).DBFieldName) = "[IZIPCODEID]" And Trim(fMainForm.ActiveForm.hfgSearchGrid.Text) <> "" Then
                        l_str_Sql = "SELECT vZipCodeName FROM tbSTEZipCodeMaster"
                        l_str_Sql = l_str_Sql & " WHERE iZipCodeId=" & fMainForm.ActiveForm.hfgSearchGrid.Text
                        l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
                        If Not l_obj_Rst.EOF Then
                            fMainForm.ActiveForm.hfgSearchGrid.Text = l_obj_Rst("vZipCodeName")
                        End If
                        l_obj_Rst.Close
                        Set l_obj_Rst = Nothing
                    ElseIf UCase(.Item(lngCol).DBFieldName) = "[IHIGHSCHOOLID]" And Trim(fMainForm.ActiveForm.hfgSearchGrid.Text) <> "" Then
'                        l_str_Sql = "SELECT vHighSchoolCode FROM tbSTEHighSchoolType"
'                        l_str_Sql = l_str_Sql & " WHERE iHighSchoolId=" & fMainForm.ActiveForm.hfgSearchGrid.Text
'                        l_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
'                        If Not l_obj_Rst.EOF Then
'                            fMainForm.ActiveForm.hfgSearchGrid.Text = l_obj_Rst("vHighSchoolCode")
'                        End If
'                        l_obj_Rst.Close
'                        Set l_obj_Rst = Nothing
                    End If
                Next
            ElseIf UCase(.Item(lngCol).strDataType) = "DATE" Then
                For lngRow = 1 To fMainForm.ActiveForm.hfgSearchGrid.Rows - 1
                    fMainForm.ActiveForm.hfgSearchGrid.Row = lngRow
                    fMainForm.ActiveForm.hfgSearchGrid.Col = .Item(lngCol).GridColPos
                    
                    If Trim(fMainForm.ActiveForm.hfgSearchGrid.Text) <> "" Then
                        fMainForm.ActiveForm.hfgSearchGrid.Text = g_dt_ConvertDate(fMainForm.ActiveForm.hfgSearchGrid.Text)
                    End If
                Next
            End If
        Next
    End With
    Screen.MousePointer = vbArrow
    fMainForm.ActiveForm.hfgSearchGrid.Visible = True
    
End Sub


Public Function f_lng_CreateNewId() As Long
    ' function to generate the next id(primary key value)
    ' when going for a new row of data
    Dim l_str_Sql As String
    Dim m_obj_Rst As New Recordset
    Dim l_str_FldName As String
    Dim l_str_TblName As String
    Dim l_int_Id As Long
    Dim l_str_txtBoxName As String
    
    Dim ctl As Control
    Dim sCtlType As String
    
    On Error GoTo ErrorHandler
    
    'find the highest existing id value
    
    l_str_FldName = fMainForm.ActiveForm.m_colFieldDetails.Item(1).DBFieldName
    l_str_TblName = fMainForm.ActiveForm.m_TableName
    l_str_txtBoxName = Left(l_str_FldName, Len(l_str_FldName) - 1)
    l_str_txtBoxName = "txt" & Right(l_str_txtBoxName, Len(l_str_txtBoxName) - 2)
    
    l_str_Sql = "Select " & l_str_FldName & " from " & l_str_TblName & _
        " ORDER BY " & l_str_FldName
    
    Set m_obj_Rst = New ADODB.Recordset
    m_obj_Rst.Open l_str_Sql, g_obj_Conn, adOpenStatic, adLockReadOnly
    
    If Not (m_obj_Rst.EOF) Then
        ' add 1 to the highest recordid to get the new recordID value
        m_obj_Rst.MoveLast
        l_int_Id = m_obj_Rst(0).Value
        l_int_Id = l_int_Id + 1
    Else
         ' This is the first value for recordid
        Dim l_obj_Rst As New ADODB.Recordset
        
        l_str_Sql = "Select iTableCounterIdMapping from tbSTETableIdMapping where vTableName = '" & l_str_TblName & "'"
        Set l_obj_Rst = g_obj_Conn.Execute(l_str_Sql)
        If Not l_obj_Rst.EOF Then
            l_int_Id = l_obj_Rst(0)
        Else
            fMainForm.ActiveForm.lblErrorMsg.Visible = True
            fMainForm.ActiveForm.lblErrorMsg.Caption = LoadResString(1124) & " - " & LoadResString(1125)
            f_lng_CreateNewId = -1
            Exit Function
        End If
        
        Set l_obj_Rst = Nothing
    End If
    
    ' free the object variable
    Set m_obj_Rst = Nothing
    f_lng_CreateNewId = l_int_Id
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
    f_lng_CreateNewId = -1
End Function

' this is to display the date in japanese format
Public Function g_dt_ConvertDate_Del20211130(l_dt_TempDate As Date) As String
 '   g_dt_ConvertDate             = Format(DateValue(l_dt_TempDate), "MM/DD/YYYY")         ' for english, use this and comment the below one
     g_dt_ConvertDate_Del20211130 = Format(DateValue(l_dt_TempDate), "gggee年mm月dd日")    ' for japanese, use this and comment the above one
End Function

' this is to display the date in japanese format
Public Function g_dt_ConvertDate(l_dt_TempDate As Date) As String

 '   g_dt_ConvertDate = Format(DateValue(l_dt_TempDate), "MM/DD/YYYY")         ' for english, use this and comment the below one

''''2021.10.14 del jhi 和暦
'''' g_dt_ConvertDate = Format(DateValue(l_dt_TempDate), "gggee年mm月dd日")    ' for japanese, use this and comment the above one


''''2021.10.14 add jhi 西暦にする
     g_dt_ConvertDate = Format(DateValue(l_dt_TempDate), "yyyy年mm月dd日")     '西暦にする

End Function

''''2020.01.27 add jhi
Public Function g_dt_ConvertDate_Seireki(l_dt_TempDate As Date) As String

     g_dt_ConvertDate_Seireki = Format(DateValue(l_dt_TempDate), "yyyy年mm月dd日")    ' for english, use this and comment the above one

End Function

Public Function g_void_SetFontProperties(l_frm_Form As Form)

    Dim l_ctl_TempCtl As Control
    Dim l_str_ctlType As String
    
    For Each l_ctl_TempCtl In l_frm_Form

        With l_ctl_TempCtl
            
            
            l_str_ctlType = TypeName(l_ctl_TempCtl)
            
            Select Case l_str_ctlType

                Case "ComboBox", "OptionButton", "TabStrip"
                    .Font.Size = 12
                    .Font.Name = "ＭＳ Ｐゴシック"        ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.Font.Name = "Verdana"               ' for english, use this and comment the above one
                    .Font.Charset = 128
                Case "VSFlexGrid", "MSFlexGrid", "ListBox"
                    .Font.Size = 10
                    .Font.Name = "ＭＳ ゴシック"          ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.Font.Name = "Verdana"               ' for english, use this and comment the above one
                    .Font.Charset = 128
                Case "DTPicker"
                    .Font.Size = 12
                    .Font.Name = "ＭＳ Ｐゴシック"        ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.Font.Name = "Verdana"               ' for english, use this and comment the above one
                    .Font.Charset = 128
                    .Format = 3
''''                .CustomFormat = "gggee年mm月dd日"      ' for japanese, use this(give proper font name in japanese) and comment the below one
                    .CustomFormat = "yyyy年mm月dd日"       ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.CustomFormat = "MM/dd/yyyy"          ' for english, use this and comment the above one
                Case "Label"
                    .Font.Size = 12
                    .Font.Name = "ＭＳ Ｐゴシック"      ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.Font.Name = "Verdana"               ' for english, use this and comment the above one
                    .Font.Charset = 128
                    If InStr(1, UCase(.Name), "ERR") > 0 Then
                    
                    Else
'                        .Alignment = 1
'受験生データ確認画面の高校情報は情報ラベルのため、色変更しない
                        If .ForeColor <> &H8000& Then
                            .ForeColor = &H800000
                        End If
                    End If
                    .BackStyle = 0
                    .Height = 450
                    
                Case "CommandButton", "MDIForm"
                    .Font.Size = 12
                    .Font.Name = "ＭＳ Ｐゴシック"      ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.Font.Name = "Verdana"               ' for english, use this and comment the above one
                    .Font.Charset = 128

                Case "TreeView"
                    .Font.Size = 10
                    .Font.Name = "ＭＳ Ｐゴシック"      ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.Font.Name = "Verdana"             ' for english, use this and comment the above one
                    .Font.Charset = 128

                Case "TextBox"
                    .Font.Size = 12
                    .Font.Name = "ＭＳ Ｐゴシック"      ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.Font.Name = "Verdana"               ' for english, use this and comment the above one
                    .Font.Charset = 128
                    If .MultiLine = False Then .Height = 350
                    If .Locked = True Then .BackColor = &HE0E0E0
                Case "CheckBox"
                    '.Font.Size = 12
                    '.Font.Name = "ＭＳ Ｐゴシック"      ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.Font.Name = "Verdana"               ' for english, use this and comment the above one
                    '.Font.Charset = 128
                    .Height = 200
                    .Width = 200
                Case "MSFlexGrid", "MSHFlexGrid"
                    .Font.Size = 10
                    .Font.Name = "ＭＳ Ｐゴシック"      ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.Font.Name = "Verdana"               ' for english, use this and comment the above one
                    .Font.Charset = 128
                    .FocusRect = 0
                    .HighLight = 2
                Case "VSFlexGrid", "VSHFlexGrid"
                    .Font.Size = 10
                    .Font.Name = "ＭＳ Ｐゴシック"      ' for japanese, use this(give proper font name in japanese) and comment the below one
                    '.Font.Name = "Verdana"               ' for english, use this and comment the above one
                    .Font.Charset = 128
                    .FocusRect = 0
                    .HighLight = flexHighlightWithFocus
            End Select
        End With
        
    Next

End Function

Public Sub g_void_HighlightRow(l_int_CurRow As Long, ByVal l_int_PrevRow As Long)
' logic changed on 28/08/02
    Dim l_int_ColCounter As Integer
    Dim l_int_OldCol As Integer
    ' show the currently selected row in a different color
    With fMainForm.ActiveForm.hfgSearchGrid
        l_int_OldCol = .Col
        .FocusRect = 1
        
        ' clear the current selection
        If l_int_PrevRow <> 0 Then
            If l_int_PrevRow >= .Rows Then l_int_PrevRow = .Rows - 1
            .Row = l_int_PrevRow
            For l_int_ColCounter = 0 To .cols - 1
                .Col = l_int_ColCounter
                .CellBackColor = .BackColor ' &HFFFFFF
                .CellForeColor = .ForeColor
            Next
        End If
        
        ' highlight the new row
        .Row = l_int_CurRow
        For l_int_ColCounter = 0 To .cols - 1
            .Col = l_int_ColCounter
            .CellBackColor = .ForeColor ' &H800000
            .CellForeColor = .BackColor
        Next
        
        .Row = l_int_CurRow
        .Col = l_int_OldCol                         ' set back the old column
        
        .FocusRect = 0
    End With
End Sub

' refresh the serial number of the grids after removing an item
Public Sub g_void_RefreshGrid(MyGrid As VSFlexGrid)

    Dim i As Integer

    With MyGrid
        For i = 1 To MyGrid.Rows - 1
            .Row = i
            .Col = 0
            .Text = i
        Next

    End With

End Sub

Public Function g_void_CloseChildForm() As Boolean

    Dim l_frm_Form As Form
    Dim l_int_Counter As Integer
    Dim l_int_index As Integer
    On Error GoTo ErrorHandler
    
    For Each l_frm_Form In Forms
        l_int_Counter = l_int_Counter + 1
    Next

    If l_int_Counter <= 2 Then
        For l_int_index = 1 To fMainForm.Toolbar1.Buttons.Count
            ' disable the toolbar buttons
           fMainForm.Toolbar1.Buttons(l_int_index).Enabled = False
        Next
        fMainForm.mnuTools.Enabled = False

    End If

    Exit Function

ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Function

Public Function g_str_LPad(ByVal str As String, ByVal iLen As Integer) As String
    '-------------------------------------------------------------
    'Left pads a string with 0 up to iLen.
    '-------------------------------------------------------------
    Select Case iLen
    Case 1
        g_str_LPad = "000" & str
    Case 2
        g_str_LPad = "00" & str
    Case 3
        g_str_LPad = "0" & str
    Case 4
        g_str_LPad = str
    End Select

End Function

Public Function gflInsChoseiJoken(ByVal piNendo As Integer _
                                , ByVal piSubjectProfileID As Integer _
                                , ByVal piChoseiJokenType As Integer _
                                , ByVal pdChouseiStartScore As Double _
                                , ByVal pdChouseiEndScore As Double _
                                , ByVal psTaishoBi As String _
                                , ByVal piRoomID As Integer _
                                , ByVal pdChouseiScore As Double _
                                ) As Long

Dim sSQL As String

    gflInsChoseiJoken = -99

On Error GoTo ErrProc

    sSQL = "Insert Into tbSTEChoseiJoken ( "
    sSQL = sSQL & "  iNendo "
    sSQL = sSQL & ", iSubjectProfileID "
    sSQL = sSQL & ", iChoseiJokenType "
    sSQL = sSQL & ", fChoseiStartScore "
    sSQL = sSQL & ", fChoseiEndScore "
    sSQL = sSQL & ", dtTaishoBi "
    sSQL = sSQL & ", iRoomID "
    sSQL = sSQL & ", fChoseiScore "
    sSQL = sSQL & " ) values ( "
    sSQL = sSQL & "  " & str(piNendo) & " "
    sSQL = sSQL & ", " & str(piSubjectProfileID) & " "
    sSQL = sSQL & ", " & str(piChoseiJokenType) & " "
    sSQL = sSQL & ", " & IIf(pdChouseiStartScore = -99, "NULL", Format(pdChouseiStartScore, "##0.0")) & " "
    sSQL = sSQL & ", " & IIf(pdChouseiEndScore = -99, "NULL", Format(pdChouseiEndScore, "##0.0")) & " "
    sSQL = sSQL & ", " & IIf(psTaishoBi = "-1", "NULL", "'" & psTaishoBi & "'") & " "
    sSQL = sSQL & ", " & IIf(piRoomID = -1, "NULL", str(piRoomID)) & " "
    sSQL = sSQL & ", " & str(pdChouseiScore) & ""
    sSQL = sSQL & ") "

    g_obj_Conn.Execute sSQL

gflInsChoseiJoken = 0

Exit Function

ErrProc:

End Function

Public Function gf_CheckPhase(iNendo, iPhase)

Dim sSQL As String
Dim oRs As ADODB.Recordset

'Phase0はfChoseiScoreがはいっていたら修正不可にする
'tbSTEExamineeProfileでiNendo、iExameineeStatusが入力されたものより大きいものがあった場合、
'Falseをかえす。

    sSQL = "SELECT "
    sSQL = sSQL & "  MAX(iNendo)"
    sSQL = sSQL & ", MAX(iExamineeStatus)"
    sSQL = sSQL & " FROM tbSTEExamineeProfile "
    sSQL = sSQL & " WHERE iNendo >= " & str(iNendo)

    Set oRs = g_obj_Conn.Execute(sSQL)

    If oRs.Fields(0) > iNendo Or oRs.Fields(1) > iPhase Then
    End If

End Function

Public Function gflDelChoseiJoken(ByVal piNendo As Integer _
                                , ByVal piSubjectProfileID As Integer _
                                , ByVal piChoseiJokenType As Integer _
                                ) As Long

Dim sSQL As String

    gflDelChoseiJoken = -99

On Error GoTo ErrProc

    sSQL = "Delete From tbSTEChoseiJoken "
    sSQL = sSQL & "WHERE iSubjectProfileID = " & str(piSubjectProfileID) & " "
    sSQL = sSQL & "AND   iChoseiJokenType = " & str(piChoseiJokenType) & " "
    sSQL = sSQL & "AND   iNendo = " & str(piNendo) & " "

    g_obj_Conn.Execute sSQL

gflDelChoseiJoken = 0

Exit Function

ErrProc:

End Function

'----------------------------------------------------
' 入力制限処理(0～9 & Period ) 評定値用
'
'----------------------------------------------------
Public Sub NumericPeriodABVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case "A", "B", "a", "b"
            Exit Sub            '--- A,a,B,bは入力可
        Case "."
            If InStr(ovsfGrd.EditText, ".") = 0 Then
                Exit Sub            '--- .(ピリオド)は１度だけ入力可
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' 入力制限処理(0～9 ) 欠席日数用
'
'----------------------------------------------------
Public Sub NumericABVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case "A", "B", "a", "b"
            Exit Sub            '--- A,a,B,bは入力可
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' 入力制限処理(0～9 ) 欠席日数用
'
'----------------------------------------------------
Public Sub NumericVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0～9は入力可
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpaceは入力可
    End Select
    pKeyAscii = 0
    
End Sub

Public Function getNewId(sTableName As String, sIDColName As String, lNewId As Long) As Boolean

Dim sSQL As String
Dim oRs As New ADODB.Recordset

    getNewId = False

    sSQL = "SELECT ISNULL( MAX( " & sIDColName & " ) , -1 )  FROM " & sTableName
    oRs.Open sSQL, g_obj_Conn, adOpenStatic, adLockReadOnly
    If oRs.Fields(0) > 0 Then
        lNewId = oRs.Fields(0) + 1
        ' release the object variable
        oRs.Close
        Set oRs = Nothing
    Else
        oRs.Close
        sSQL = "SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName='" & sTableName & "'"
        oRs.Open sSQL, g_obj_Conn, adOpenStatic, adLockReadOnly
        If Not oRs.EOF Then
            lNewId = oRs.Fields(0)
            oRs.Close
        Else
            lNewId = 1
        End If
        Set oRs = Nothing
    End If

    getNewId = True

End Function
