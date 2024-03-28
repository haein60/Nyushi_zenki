Attribute VB_Name = "mdlComm"
'*******************************************************************************
'* ���ʊ֐��Q                                                                  *
'* �쐬�� : 2021.12.10                                                         *
'* �쐬�� : jyon hein                                                          *
'*******************************************************************************

Option Explicit

Private Const GWL_STYLE                  As Long = (-16)
Private Const TVS_HASLINES               As Long = 2
Private Const TV_FIRST                   As Long = &H1100
Private Const TVM_SETBKCOLOR             As Long = (TV_FIRST + 29)

Private Declare Function SendMessage Lib "user32" _
                         Alias "SendMessageA" _
                         (ByVal hwnd As Long, _
                          ByVal wMsg As Long, _
                          ByVal wParam As Long, _
                          lParam As Any) As Long

Private Declare Function GetWindowLong Lib "user32" _
                         Alias "GetWindowLongA" _
                         (ByVal hwnd As Long, _
                          ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
                         Alias "SetWindowLongA" _
                         (ByVal hwnd As Long, _
                          ByVal nIndex As Long, _
                          ByVal dwNewLong As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32" _
                         (ByVal clr As OLE_COLOR, _
                          ByVal hPal As Long, _
                          dwRGB As Long) As Long


'*******************************************************************************
'* ���l�ϊ����čēx������ɕϊ�����0���폜������@                             *
'* Val�֐���Double�^�ɕϊ���������CStr�֐��ŕ�����ɖ߂��֐��ł��             *
'* �����ɕϊ��������������n���Ďg���܂��                                     *
'*******************************************************************************
Public Function DelZero(s As String) As String
    
    Dim ret As String
    
    ret = CStr(Val(s))
    
    DelZero = ret
    
End Function

'*******************************************************************************
'* �����P�F������                                                              *
'* �����Q�F�폜������                                                          *
'* �߂�l�F�폜��̕�����                                                      *
'*******************************************************************************
Public Function CutLeft(s As String, i As Long) As Variant
    
    Dim iLen    As Long     '������
    
    
    '������ł͂Ȃ��ꍇ
'    If VarType(s) <> vbString Then
'        Exit Function
'    End If
    
    iLen = Len(s)
    
    ' �����񒷂��w�蕶�������傫���ꍇ
    If iLen < i Then
        Exit Function
    End If
    
    
    ' �w�蕶�������폜���ĕԂ�
    If (Mid(s, 1, 1) = "0") Then
        CutLeft = CVar(Right(s, iLen - i))
    Else
        CutLeft = CVar(s)
    End If
    
End Function

'===============================================================================
' �w��̕������ɂȂ�܂Ő擪�𕶎��Ŗ��߂܂��B
'
' @Param    stTarget    �����ΏۂƂȂ镶����B
' @Param    iLength     �����̒����B
' @Param    [chOne]     ���߂镶���B
' @Return               �擪���w��̕����� iLength �̒����܂Ŗ��߂�ꂽ������B
'===============================================================================
Public Function PadLeft(stTarget, iLength, chOne)
   
   Do While (Len(stTarget) < iLength)
       stTarget = chOne & stTarget
   Loop

   PadLeft = Right(stTarget, iLength)

End Function

'===============================================================================
'���p�E�S�p�����݂���悤�ȏꍇ�́A���L�̊֐����g�p����
'2021.12.16 add jhi
'===============================================================================
Public Function fPadLeft(ByVal myData As String, ByVal CutLen As Long, ByVal CutStr As String) As String

    '�������E�񂹂��A�w�肵��������̕������ɂȂ�܂ō����Ɏw�肵������(0 �� " " ��)�𖄂ߍ��݂܂��B
    Dim tmp As String

    tmp = StrConv(RightB$(StrConv(String$(CutLen, CutStr) & myData, vbFromUnicode), CutLen), vbUnicode)
    fPadLeft = tmp

End Function


'*******************************************************************************
'* debug�p log�֐�                                                             *
'* (�ȈՔ�:parameter��������string 1�̂�)                                    *
'*******************************************************************************
Public Sub log(ByVal str As String)

    ''''system date and time���擾����
    Dim strDateTime    As String
    Dim sYM            As String
    Dim fName          As String
    
    Dim FileNumber     As Integer

    
    strDateTime = Format(Now, "yyyy/MM/dd HH:mm:ss") & " "
    
    '�V�X�e�����t���N�����擾����
    sYM = Format(Now, "yyyymm")
    'Debug.Print sYM

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    
    '�t�@�C����Append���[�h�ŊJ���܂��B
    fName = App.Path & "\log_" & sYM & ".txt"
    Open fName For Append As #FileNumber
    
    Print #FileNumber, strDateTime & str
    
    Close #FileNumber

End Sub

'*******************************************************************************
'* �`�F�b�N���ʂ��o�͂��� csv�֐�                                              *
'*******************************************************************************
Public Sub logcsv(title_flag As Integer, ByVal str As String)

    'system date and time���擾����
    Dim strDateTime    As String
    Dim sYM            As String
    Dim sYMD           As String
    Dim fName          As String
    
    Dim FileNumber     As Integer

    
    strDateTime = Format(Now, "yyyy/MM/dd HH:mm:ss") & ","
    
    '�V�X�e�����t���N�����擾����
    sYM = Format(Now, "yyyymm")
    sYMD = Format(Now, "yyyymmdd")
''''gYMD = sYMD
    
    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    
    '�t�@�C����Append���[�h�ŊJ���܂��B
    fName = ThisWorkbook.Path & "\log_" & sYMD & ".csv"
    Open fName For Append As #FileNumber
    
    'title_flag���A�擪��strDateTime���ȗ����邩�A�o�͂���
    If (title_flag = 1) Then
        Print #FileNumber, str
    Else
        Print #FileNumber, strDateTime & str
    End If
    
    Close #FileNumber

End Sub

'*******************************************************************************
'* �`�F�b�N���ʂ��o�͂��� csv�֐�                                              *
'*******************************************************************************
Public Sub logcsv_2(fn As String, title_flag As Integer, ByVal str As String)

    'system date and time���擾����
    Dim strDateTime    As String
    Dim sYM            As String
    Dim sYMD           As String
    Dim fName          As String

    Dim FileNumber     As Integer

    
    strDateTime = Format(Now, "yyyy/MM/dd HH:mm:ss") & ","

    '�V�X�e�����t���N�����擾����
    sYM = Format(Now, "yyyymm")
    sYMD = Format(Now, "yyyymmdd")
''''gYMD = sYMD
    
    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    
    '�t�@�C����Append���[�h�ŊJ���܂��B
    fName = ThisWorkbook.Path & "\log" & fn & "_" & sYMD & ".csv"
    Open fName For Append As #FileNumber
    
    'title_flag���A�擪��strDateTime���ȗ����邩�A�o�͂���
    If (title_flag = 1) Then
        Print #FileNumber, str
    Else
        Print #FileNumber, strDateTime & str
    End If
    
    Close #FileNumber

End Sub

'*******************************************************************************
'* ���ʃt�@�C�����폜���ď�����������                                          *
'*******************************************************************************
Public Sub Del_Csvfile(dummy As String)
    
    Dim FSO      As Object
    
    Dim sYMD     As String
    Dim fName    As String
    
    
    On Error Resume Next
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    
    '�V�X�e�����t���N�����擾����
    sYMD = Format(Now, "yyyymmdd")
        
    fName = ThisWorkbook.Path & "\log_" & sYMD & ".csv"
    FSO.DeleteFile fName
    
    Set FSO = Nothing


End Sub
'*******************************************************************************
'* ���ʃt�@�C�����폜���ď�����������                                          *
'*******************************************************************************
Public Sub Del_Csvfile_2(fn As String)
    
    Dim FSO      As Object
    
    Dim sYMD     As String
    Dim fName    As String
    
    
    On Error Resume Next
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    
    '�V�X�e�����t���N�����擾����
    sYMD = Format(Now, "yyyymmdd")
        
    fName = ThisWorkbook.Path & "\log" & fn & "_" & sYMD & ".csv"
    FSO.DeleteFile fName
    
    Set FSO = Nothing


End Sub

'*******************************************************************************
'* �I���������j���[����V�X�e���萔Table tbSTESystemProfile�Ƀt�F�[�Yflag��    *
'* �Z�b�g����                                                                  *
'*-----------------------------------------------------------------------------*
'* 2021.12.09 add jhi                                                          *
'*******************************************************************************
Public Sub Phase_FlagSet(phno As Integer)

    On Error GoTo ErrorHandler

    Dim l_obj_Rst      As New ADODB.Recordset
    Dim sSQL           As String
    Dim rinf           As Integer
    


    '-----------------------------------------------------------------------
    ' tbSTESystemProfile table�Ƀt�F�[�Yflag���Z�b�g����
    '-----------------------------------------------------------------------
    sSQL = ""
    sSQL = "update tbSTESystemProfile set iCurrentPhase=" & phno & " where iActiveFlag=1"
    g_obj_Conn.Execute (sSQL)
        
    '
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
    End If

    Exit Sub


ErrorHandler:
    MsgBox Err.Description

End Sub

'*******************************************************************************
'* myMsgBox                                                                    *
'*-----------------------------------------------------------------------------*
'* 2021.12.09 add jhi                                                          *
'*******************************************************************************
Public Function myMsgBox(sMsg As String, sTit) As Long

    Dim rinf As Long

    rinf = MsgBox(sMsg, vbOKCancel, sTit)
    
    myMsgBox = rinf

End Function

'*******************************************************************************
'* Treeview Background change                                                  *
'*-----------------------------------------------------------------------------*
'* 2021.12.09 add jhi                                                          *
'*******************************************************************************
Public Sub SetTVBackColor(pobjTV As TreeView, plngBackColor As Long)
 
    Dim lngTVHwnd   As Long
    Dim lngStyle    As Long
    Dim objTVNode   As Node
    

    lngTVHwnd = pobjTV.hwnd
    
    ' Change the background
    Call SendMessage(lngTVHwnd, TVM_SETBKCOLOR, 0, ByVal plngBackColor)
    
    ' Set the backcolor of the nodes ...
    For Each objTVNode In pobjTV.Nodes
        objTVNode.BackColor = plngBackColor
    Next
 
    ' Reset the treeview style so the tree lines appear properly ...
    lngStyle = GetWindowLong(lngTVHwnd, GWL_STYLE)
    
    ' If the treeview has lines, temporarily remove them so the back
    ' repaints to the selected colour, then restore ...
    If lngStyle And TVS_HASLINES Then
       Call SetWindowLong(lngTVHwnd, GWL_STYLE, lngStyle Xor TVS_HASLINES)
       Call SetWindowLong(lngTVHwnd, GWL_STYLE, lngStyle)
    End If

    
End Sub


Public Function StrNullCut(psInStr As String) As String

    Dim lPos As Long


    lPos = InStr(1, psInStr, vbNullChar)

    If lPos > 0 Then
        StrNullCut = Left$(psInStr, lPos - 1)
    Else
        StrNullCut = psInStr
    End If

End Function

'*******************************************************************************
'�y�@�\�z    �t�@�C����ʂ̏ꏊ�փR�s�[���܂��B
'CopyFile ���\�b�h
'[�Q�Ɛݒ�]
'Microsoft Scripting Runtime (scrrun.dll)
'2022.02.08 add jhi
'*******************************************************************************
Public Sub fCopy(strSrcName As String, strDestName As String)

    On Error GoTo ErrorHandler

    'FileSystemObject�C���X�^���X�𐶐�
    Dim FSO As Object


    Set FSO = CreateObject("Scripting.FileSystemObject")

    '�t�@�C�����R�s�[
    FSO.CopyFile strSrcName, strDestName, True '�㏑�����̏ꍇ

    '�I�u�W�F�N�g�̉��
    Set FSO = Nothing
    Exit Sub


ErrorHandler:
    MsgBox Err.Description, vbInformation, "�G���["

End Sub

