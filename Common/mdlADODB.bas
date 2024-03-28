Attribute VB_Name = "mdlADODB"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Type puPrm_Type
    pName       As String
    pType       As Long
    pDirection  As Long
    pSize       As Long
    pValue      As Variant
End Type


'‰Šú‰»ƒtƒ@ƒCƒ‹‚ÌƒfƒtƒHƒ‹ƒgBˆø”‚ª‚È‚¢‚Æ‚«‚ÉŽg—p
Private Const prvsProfileName As String = "FANET.Ini"

Private Function lf_LongCheck(psNum As String) As Boolean

Dim lWk As Long

On Error GoTo ErrProc

lf_LongCheck = False

    lWk = CLng(psNum)

lf_LongCheck = True

Exit Function

ErrProc:

End Function

Private Function pf_StrNullCut(psInStr As String) As String

Dim lPos As Long

    lPos = InStr(1, psInStr, vbNullChar)

    If lPos > 0 Then
        pf_StrNullCut = Left$(psInStr, lPos - 1)
    Else
        pf_StrNullCut = psInStr
    End If

End Function

'‚c‚a‚É‚½‚¢‚µ‚Ä‚`‚c‚n‚ð‰î‚µ‚Ä‚r‚p‚k‚ðŽÀs‚·‚é
'ƒGƒ‰[Žž‚Í|‚Q(ƒRƒlƒNƒVƒ‡ƒ“‚ÌŠm—§‚à‚µ‚Ä‚¢‚È‚¢)
'‚ð‚©‚¦‚µAƒŒƒR[ƒh‚È‚µŽž‚Í|‚P‚ð‚©‚¦‚µA³íŽž‚Í‚O‚ð‚©‚¦‚·
Public Function pf_ExecSQL(poCn As ADODB.Connection, poRs As ADODB.Recordset, psSQL As String, psErrMsg As String, Optional pbBeginTrans As Boolean, Optional psIniFileName As String) As Integer

Dim iErrPos As Integer
Dim sDSN As String
Dim sDatabase As String
Dim sUID As String
Dim sPWD As String
Dim bOpen As Boolean
Dim sQueryTimeOut As String

iErrPos = 0
bOpen = False

pf_ExecSQL = -2

On Error GoTo ErrProc

    If poCn.State = adStateClosed Then
        bOpen = True
        If IsEmpty(psIniFileName) Then
            Call psADODB_ODBCIniRead(sDSN, sDatabase, sUID, sPWD, sQueryTimeOut)
        Else
            Call psADODB_ODBCIniRead(sDSN, sDatabase, sUID, sPWD, sQueryTimeOut, psIniFileName)
        End If
        poCn.ConnectionString = "DSN=" & sDSN & ";UID=" & sUID & ";PWD=" & sPWD & ";database=" & sDatabase
        poCn.CommandTimeout = sQueryTimeOut
        poCn.Open
    End If

iErrPos = 1

    If pbBeginTrans Then poCn.BeginTrans

    Set poRs = poCn.Execute(psSQL)

iErrPos = 2

    If poRs.BOF Then
        'ƒŒƒR[ƒh‚È‚µ
        pf_ExecSQL = -1
        Exit Function
    End If

pf_ExecSQL = 0

Exit Function

ErrProc:

    psErrMsg = "ƒf[ƒ^ƒx[ƒX‚Ö‚Ì–â‚¢‡‚í‚¹‚ÉŽ¸”s‚µ‚Ü‚µ‚½" & vbCrLf & "ErrNo=&H" & Hex$(Err.Number) & vbCrLf & "ErrDescription=" & Err.Description

    If iErrPos > 0 Then
    '‚b‚‚Ž‚Ž‚…‚ƒ‚”‚‰‚‚Ž‚ÍŠm—§‚µ‚Ä‚¢‚é
        If bOpen Then
            poCn.Close
        Else
            pf_ExecSQL = -3
        End If
        If iErrPos > 1 Then
            'ƒŒƒR[ƒhƒZƒbƒg‚ªŠJ‚¢‚Ä‚¢‚é
            poRs.Close
        End If
    End If

End Function

'‚c‚a‚É‚½‚¢‚µ‚Ä‚`‚c‚n‚ð‰î‚µ‚Ä‚r‚p‚k‚ðŽÀs‚·‚é
'ƒGƒ‰[Žž‚Í|‚Q(ƒRƒlƒNƒVƒ‡ƒ“‚ÌŠm—§‚à‚µ‚Ä‚¢‚È‚¢)
'‚ð‚©‚¦‚µAƒŒƒR[ƒh‚È‚µŽž‚Í|‚P‚ð‚©‚¦‚µA³íŽž‚Í‚O‚ð‚©‚¦‚·
Public Function pf_OpenSQL(poCn As ADODB.Connection, poRs As ADODB.Recordset, psSQL As String, psErrMsg As String, Optional pbBeginTrans As Boolean, Optional plCursorType As Long, Optional plLockType As Long, Optional psIniFileName As String) As Integer

Dim iErrPos As Integer
Dim sDSN As String
Dim sDatabase As String
Dim sUID As String
Dim sPWD As String
Dim bOpen As Boolean
Dim sQueryTimeOut As String

iErrPos = 0
bOpen = False

pf_OpenSQL = -2

On Error GoTo ErrProc

    If poCn.State = adStateClosed Then
        bOpen = True
        If IsEmpty(psIniFileName) Then
            Call psADODB_ODBCIniRead(sDSN, sDatabase, sUID, sPWD, sQueryTimeOut)
        Else
            Call psADODB_ODBCIniRead(sDSN, sDatabase, sUID, sPWD, sQueryTimeOut, psIniFileName)
        End If
        poCn.ConnectionString = "DSN=" & sDSN & ";UID=" & sUID & ";PWD=" & sPWD & ";database=" & sDatabase
        poCn.CommandTimeout = sQueryTimeOut
        poCn.Open
    End If

iErrPos = 1

    If pbBeginTrans Then poCn.BeginTrans

'    poRs.CursorType = plCursorType
'    Set poRs = poCn.Execute(psSQL)
    poRs.Open psSQL, poCn, plCursorType, plLockType

iErrPos = 2

    If poRs.BOF Then
        'ƒŒƒR[ƒh‚È‚µ
        pf_OpenSQL = -1
        Exit Function
    End If

pf_OpenSQL = 0

Exit Function

ErrProc:

    psErrMsg = "ƒf[ƒ^ƒx[ƒX‚Ö‚Ì–â‚¢‡‚í‚¹‚ÉŽ¸”s‚µ‚Ü‚µ‚½" & vbCrLf & "ErrNo=&H" & Hex$(Err.Number) & vbCrLf & "ErrDescription=" & Err.Description

    If iErrPos > 0 Then
    '‚b‚‚Ž‚Ž‚…‚ƒ‚”‚‰‚‚Ž‚ÍŠm—§‚µ‚Ä‚¢‚é
        If bOpen Then
            poCn.Close
        Else
            pf_OpenSQL = -3
        End If
        If iErrPos > 1 Then
            'ƒŒƒR[ƒhƒZƒbƒg‚ªŠJ‚¢‚Ä‚¢‚é
            poRs.Close
        End If
    End If

End Function

'‚c‚a‚É‚½‚¢‚µ‚Ä‚`‚c‚n‚ð‰î‚µ‚Ä‚r‚p‚k‚ðŽÀs‚·‚é
'ƒGƒ‰[Žž‚Í|‚Q(ƒRƒlƒNƒVƒ‡ƒ“‚ÌŠm—§‚à‚µ‚Ä‚¢‚È‚¢)
'‚ð‚©‚¦‚µAƒŒƒR[ƒh‚È‚µŽž‚Í|‚P‚ð‚©‚¦‚µA³íŽž‚Í‚O‚ð‚©‚¦‚·
Public Function pf_ExecSQL_NoRtn(poCn As ADODB.Connection, poRs As ADODB.Recordset, psSQL As String, psErrMsg As String, Optional pbBeginTrans As Boolean, Optional psIniFileName As String) As Integer

Dim iErrPos As Integer
Dim sDSN As String
Dim sDatabase As String
Dim sUID As String
Dim sPWD As String
Dim sQueryTimeOut As String
Dim lRecordsAffected As Long
Dim bOpen As Boolean

iErrPos = 0
bOpen = False

pf_ExecSQL_NoRtn = -2

On Error GoTo ErrProc

    If poCn.State = adStateClosed Then
        bOpen = True
        If IsEmpty(psIniFileName) Then
            Call psADODB_ODBCIniRead(sDSN, sDatabase, sUID, sPWD, sQueryTimeOut)
        Else
            Call psADODB_ODBCIniRead(sDSN, sDatabase, sUID, sPWD, sQueryTimeOut, psIniFileName)
        End If
        poCn.ConnectionString = "DSN=" & sDSN & ";UID=" & sUID & ";PWD=" & sPWD & ";database=" & sDatabase
        poCn.CommandTimeout = sQueryTimeOut
        poCn.Open
    End If

iErrPos = 1

    lRecordsAffected = 0

'    Set poRs = poCn.Execute(psSQL, lRecordsAffected)
    poCn.Execute psSQL, lRecordsAffected

    If lRecordsAffected = 0 Then
        pf_ExecSQL_NoRtn = -1
        Exit Function
    End If

pf_ExecSQL_NoRtn = 0

Exit Function

ErrProc:

    psErrMsg = "ƒf[ƒ^ƒx[ƒX‚Ö‚Ì–â‚¢‡‚í‚¹‚ÉŽ¸”s‚µ‚Ü‚µ‚½" & vbCrLf & "ErrNo=&H" & Hex$(Err.Number) & vbCrLf & "ErrDescription=" & Err.Description

    If iErrPos > 0 Then
    '‚b‚‚Ž‚Ž‚…‚ƒ‚”‚‰‚‚Ž‚ÍŠm—§‚µ‚Ä‚¢‚é
        If bOpen Then
            poCn.Close
        End If
    End If

End Function

'‚c‚a‚É‚½‚¢‚µ‚Ä‚`‚c‚n‚ð‰î‚µ‚Ä‚r‚p‚k‚ðŽÀs‚·‚é
'ƒGƒ‰[Žž‚Í|‚Q(ƒRƒlƒNƒVƒ‡ƒ“‚ÌŠm—§‚à‚µ‚Ä‚¢‚È‚¢)
'‚ð‚©‚¦‚µAƒŒƒR[ƒh‚È‚µŽž‚Í|‚P‚ð‚©‚¦‚µA³íŽž‚Í‚O‚ð‚©‚¦‚·
Public Function pf_CmdExecSQL_Stored(poCn As ADODB.Connection, poCmd As ADODB.Command, puPrm_() As puPrm_Type, psSQL As String, psErrMsg As String, Optional pbBeginTrans As Boolean = False, Optional psIniFileName As String) As Integer

Dim iErrPos As Integer
Dim sDSN As String
Dim sDatabase As String
Dim sUID As String
Dim sPWD As String
Dim bOpen As Boolean
Dim oPrm()    As ADODB.Parameter
Dim ii As Integer
Dim sQueryTimeOut As String

iErrPos = 0
bOpen = False

pf_CmdExecSQL_Stored = -2

On Error GoTo ErrProc

'ƒRƒlƒNƒVƒ‡ƒ“ŽÀs
    If poCn.State = adStateClosed Then
        bOpen = True
        If IsEmpty(psIniFileName) Then
            Call psADODB_ODBCIniRead(sDSN, sDatabase, sUID, sPWD, sQueryTimeOut)
        Else
            Call psADODB_ODBCIniRead(sDSN, sDatabase, sUID, sPWD, sQueryTimeOut, psIniFileName)
        End If
        poCn.ConnectionString = "DSN=" & sDSN & ";UID=" & sUID & ";PWD=" & sPWD & ";database=" & sDatabase
        poCn.CommandTimeout = sQueryTimeOut
        poCn.Open
    End If

    If pbBeginTrans Then
        poCn.BeginTrans
    End If

iErrPos = 1

    Set poCmd = New ADODB.Command

'ŽÀsƒRƒ}ƒ“ƒh‚ÌŠi”[
    poCmd.CommandText = psSQL
    poCmd.CommandType = adCmdStoredProc

'ƒpƒ‰ƒ[ƒ^‚ÌŠi”[
    ReDim oPrm(LBound(puPrm_) To UBound(puPrm_)) As ADODB.Parameter
    For ii = LBound(puPrm_) To UBound(puPrm_)
        Set oPrm(ii) = poCmd.CreateParameter(puPrm_(ii).pName, puPrm_(ii).pType, puPrm_(ii).pDirection, puPrm_(ii).pSize)
        poCmd.Parameters.Append oPrm(ii)
        oPrm(ii).Value = puPrm_(ii).pValue
    Next

    ' Create recordset by executing the command.
    Set poCmd.ActiveConnection = poCn
    poCmd.Execute

'    Set poRs = poCn.Execute(psSQL, lRecordsAffected)

pf_CmdExecSQL_Stored = 0

Exit Function

ErrProc:

    psErrMsg = "ƒf[ƒ^ƒx[ƒX‚Ö‚Ì–â‚¢‡‚í‚¹‚ÉŽ¸”s‚µ‚Ü‚µ‚½" & vbCrLf & "ErrNo=&H" & Hex$(Err.Number) & vbCrLf & "ErrDescription=" & Err.Description

    If iErrPos > 0 Then
    '‚b‚‚Ž‚Ž‚…‚ƒ‚”‚‰‚‚Ž‚ÍŠm—§‚µ‚Ä‚¢‚é
        If bOpen Then
            poCn.Close
        End If
    End If

End Function

Public Sub psADODB_ODBCIniRead(psDSN As String, psDATABASE As String, psUID As String, psPWD As String, psQTM As String, Optional psFile As String)

Dim lRtn As Long
Dim sRtn As String
Dim sProfileName As String

On Error Resume Next

    If IsEmpty(psFile) Then
        psFile = prvsProfileName
    Else
        If psFile = "" Then
            psFile = prvsProfileName
        End If
    End If

    If Right(App.Path, 1) = "\" Then
        sProfileName = App.Path & psFile
    Else
        sProfileName = App.Path & "\" & psFile
    End If

    sRtn = Space(40)
    lRtn = GetPrivateProfileString("ODBC", "DSN", "DISCO", sRtn, 40, sProfileName)
    If lRtn > 0 Then
        psDSN = pf_StrNullCut(sRtn)
    End If

    sRtn = Space(40)
    lRtn = GetPrivateProfileString("ODBC", "DATABASE", "DISCO", sRtn, 40, sProfileName)
    If lRtn > 0 Then
        psDATABASE = pf_StrNullCut(sRtn)
    End If

    sRtn = Space(40)
    lRtn = GetPrivateProfileString("ODBC", "UID", "sa", sRtn, 40, sProfileName)
    If lRtn > 0 Then
        psUID = pf_StrNullCut(sRtn)
    End If

    sRtn = Space(40)
    lRtn = GetPrivateProfileString("ODBC", "PWD", "", sRtn, 40, sProfileName)
    If lRtn > 0 Then
        psPWD = pf_StrNullCut(sRtn)
    End If

    sRtn = Space(40)
    lRtn = GetPrivateProfileString("ODBC", "QUERYTIMEOUT", "60", sRtn, 40, sProfileName)
    If lRtn > 0 Then
        psQTM = pf_StrNullCut(sRtn)
        If Not lf_LongCheck(psQTM) Then
            psQTM = 60
        End If
    End If

End Sub

Public Sub psAttendantInfoIniRead(psAttendantID As String, psAttendantPWD As String)

Dim lRtn As Long
Dim sRtn As String
Dim sProfileName As String

On Error Resume Next

    If Right(App.Path, 1) = "\" Then
        sProfileName = App.Path & prvsProfileName
    Else
        sProfileName = App.Path & "\" & prvsProfileName
    End If

    sRtn = Space(20)
    lRtn = GetPrivateProfileString("ATTENDANT", "ATTENDANTID", "999999", sRtn, 20, sProfileName)
    If lRtn > 0 Then
        psAttendantID = pf_StrNullCut(sRtn)
    End If

    sRtn = Space(20)
    lRtn = GetPrivateProfileString("ATTENDANT", "ATTENDANTPWD", "999999", sRtn, 20, sProfileName)
    If lRtn > 0 Then
        psAttendantPWD = pf_StrNullCut(sRtn)
    End If

End Sub

Public Sub psErrOut(psErrMsg As String, psOutputFileNm As String)

Dim FileNo As Long
Dim sPutStr As String

    sPutStr = Format(Now, "YYYY/MM/DD HH:NN:SS") & "  " & Replace(psErrMsg, vbCrLf, ",")

    FileNo = FreeFile
    Open psOutputFileNm For Append Access Write As #FileNo
    Print #FileNo, sPutStr
    Close #FileNo

End Sub
