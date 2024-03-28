Attribute VB_Name = "mdlWinUser"
Option Explicit

Private Const UNLEN = 256

Private Declare Function GetUserName Lib "advapi32.dll" _
    Alias "GetUserNameA" ( _
            ByVal lpBuffer As String, _
            nSize As Long) As Long

Public Function gf_GetLoginUser() As String
Dim strUser As String
    strUser = String(UNLEN, vbNullChar)
    GetUserName strUser, UNLEN - 1
    strUser = Left(strUser, InStr(1, strUser, vbNullChar) - 1)
    gf_GetLoginUser = strUser
End Function

