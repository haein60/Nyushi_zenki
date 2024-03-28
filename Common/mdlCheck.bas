Attribute VB_Name = "mdlCheck"
Option Explicit

Public Function gf_IntCheck(psNum As String) As Boolean

Dim iWk As Integer

On Error GoTo ErrProc

gf_IntCheck = False

    iWk = CInt(psNum)

gf_IntCheck = True

Exit Function

ErrProc:

End Function

Public Function gf_LongCheck(psNum As String) As Boolean

Dim lWk As Long

On Error GoTo ErrProc

gf_LongCheck = False

    lWk = CLng(psNum)

gf_LongCheck = True

Exit Function

ErrProc:

End Function

Public Function gf_DblCheck(psNum As String) As Boolean

Dim lWk As Long

On Error GoTo ErrProc

gf_DblCheck = False

    lWk = CDbl(psNum)

gf_DblCheck = True

Exit Function

ErrProc:

End Function
Public Function gf_FileCheck(psFile As String) As Boolean

Dim lWk As Long
Dim sWk As String
Dim sCkStr As String

On Error GoTo ErrProc

gf_FileCheck = False

    lWk = InStrRev(psFile, "\")

    If lWk = 0 Then Exit Function
    If lWk = Len(psFile) Then Exit Function

    sWk = StrConv(Mid(psFile, lWk + 1), vbUpperCase)

    sCkStr = Dir(Left(psFile, lWk), vbNormal)

    Do Until sCkStr = ""
        If StrConv(sCkStr, vbUpperCase) = sWk Then
            gf_FileCheck = True
            Exit Function
        End If
        sCkStr = Dir
    Loop

Exit Function

ErrProc:

End Function

'----------------------------------------------------
' ���͐�������(0�`9)
'
'----------------------------------------------------
Public Sub NumericOnly(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0�`9�͓��͉�
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpace�͓��͉�
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' ���͐�������(0�`9 & Period)
'
'----------------------------------------------------
Public Sub NumericPeriod(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0�`9�͓��͉�
        Case "."
            If InStr(F.ActiveControl, ".") = 0 Then
                Exit Sub            '--- .(�s���I�h)�͂P�x�������͉�
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpace�͓��͉�
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' ���͐�������(0�`9 & Period & '-')
'
'----------------------------------------------------
Public Sub NumericPeriodMinus(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0�`9�͓��͉�
        Case "."
            If InStr(F.ActiveControl, ".") = 0 Then
                Exit Sub            '--- .(�s���I�h)�͂P�x�������͉�
            End If
        Case "-"
            If Len(F.ActiveControl.Text) = 0 _
            Or F.ActiveControl.SelLength = Len(F.ActiveControl.Text) Then
                Exit Sub            '--- -(�}�C�i�X)�͂P�x�����A�擪�̂ݓ��͉�
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpace�͓��͉�
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' ���͐�������(0�`9 & '-')
'
'----------------------------------------------------
Public Sub NumericMinus(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
                Exit Sub            '--- 0�`9�͓��͉�
        Case "-"
            If InStr(F.ActiveControl, "-") = 0 Then
                If Len(F.ActiveControl.Text) = 0 _
                Or F.ActiveControl.SelLength = Len(F.ActiveControl.Text) Then
                    Exit Sub            '--- -(�}�C�i�X)�͓��͉�
                End If
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpace�͓��͉�
    End Select
    pKeyAscii = 0
    
End Sub

'--------------------------------------------------
' �t�H�[�J�X�擾�ǐՏ���
'
'--------------------------------------------------
Public Sub GotFocusTracking(pForm As Form, pControl As Control)
', pInitIMEMode() As Integer
    Dim wIndex   As Integer
    
    wIndex = -1: On Error Resume Next: wIndex = pControl.index: On Error GoTo 0
    Call AllControlGotFocus(pForm, pControl, wIndex)
'    Call SetIMEMode(pForm, pInitIMEMode())
    
End Sub

'--------------------------------------------------
' �t�H�[�J�X�r���ǐՏ���
'
'--------------------------------------------------
Public Sub LostFocusTracking(pForm As Form, pControl As Control)

    Dim wIndex   As Integer
    
    wIndex = -1: On Error Resume Next: wIndex = pControl.index: On Error GoTo 0
    Call AllControlLostFocus(pForm, pControl, wIndex)
    
End Sub


'--------------------------------------------------
' �S�R���g���[���t�H�[�J�X�擾������
'
'--------------------------------------------------
Private Sub AllControlGotFocus(pForm As Form, pGotControl As Control, pGotIndex As Integer)

    Dim strLostControl As String
    Dim strGotControl  As String
    
    strGotControl = pGotControl.Name
    
    On Error Resume Next
    If TypeOf pGotControl Is CommandButton Or _
       TypeOf pGotControl Is OptionButton Or _
       TypeOf pGotControl Is CheckBox Then
        '--- �R�}���h�{�^��/�I�v�V�����{�^��/�`�F�b�N�{�b�N�X�̏ꍇ��
        '    �t�H���g�̃{�[���h�ݒ�� On/Off ����
        If pGotIndex = -1 Then
            pForm.Controls(strGotControl).Font.Bold = True
        Else
            pForm.Controls(strGotControl)(pGotIndex).Font.Bold = True
        End If
    ElseIf TypeOf pGotControl Is TextBox Or _
           TypeOf pGotControl Is ListBox Or _
           TypeOf pGotControl Is ComboBox Then
        '--- �e�L�X�g�{�b�N�X/���X�g�{�b�N�X/�R���{�{�b�N�X�̏ꍇ��
        '    �o�b�N�J���[��ύX����
'        If pGotIndex = -1 Then
'            If pForm.Controls(strGotControl).Locked = False Then
'                pForm.Controls(strGotControl).BackColor = gGotFocusBackColor
'                Call FTCOverWriteMode(pForm)
'            End If
'        Else
'            If pForm.Controls(strGotControl)(pGotIndex).Locked = False Then
'                pForm.Controls(strGotControl)(pGotIndex).BackColor = gGotFocusBackColor
'                Call FTCOverWriteMode(pForm)
'            End If
'        End If
    End If
    On Error GoTo 0

End Sub

'--------------------------------------------------
' �S�R���g���[���t�H�[�J�X�r��������
'
'--------------------------------------------------
Private Sub AllControlLostFocus(pForm As Form, pLostControl As Control, pLostIndex As Integer)

    Dim strLostControl As String
    Dim strGotControl  As String
    
    strLostControl = pLostControl.Name
    
    On Error Resume Next
    If TypeOf pLostControl Is CommandButton Or _
       TypeOf pLostControl Is OptionButton Or _
       TypeOf pLostControl Is CheckBox Then
        '--- �R�}���h�{�^��/�I�v�V�����{�^��/�`�F�b�N�{�b�N�X�̏ꍇ��
        '    �t�H���g�̃{�[���h�ݒ�� On/Off ����
        If pLostIndex = -1 Then
            pForm.Controls(strLostControl).Font.Bold = False
        Else
            pForm.Controls(strLostControl)(pLostIndex).Font.Bold = False
        End If
    ElseIf TypeOf pLostControl Is TextBox Or _
           TypeOf pLostControl Is ListBox Or _
           TypeOf pLostControl Is ComboBox Then
        '--- �e�L�X�g�{�b�N�X/���X�g�{�b�N�X/�R���{�{�b�N�X�̏ꍇ��
        '    �o�b�N�J���[��ύX����
'        If pLostIndex = -1 Then
'            pForm.Controls(strLostControl).BackColor = gLostFocusBackColor
'        Else
'            pForm.Controls(strLostControl)(pLostIndex).BackColor = gLostFocusBackColor
'        End If
    End If
    On Error GoTo 0
    
End Sub

'-----------------------------------------------------
' �N���e�B�J���Z�b�V�����J�n����
'
'-----------------------------------------------------
Public Function BeginCriticalSession(pSessionKey As String) As Integer

    Dim fp As Integer

    fp = FreeFile

    On Error Resume Next
    
    '--- OS�̃t�@�C�����b�N�𗘗p���ă��b�N���s��
    Open pSessionKey & ".LCK" For Output Lock Write As #fp
    
    
    Close #fp
    Open pSessionKey & ".LCK" For Output Lock Write As #fp


    Do Until Err = 0
        Err = 0
        Close #fp
        DoEvents
        Open pSessionKey & ".LCK" For Output Lock Write As #fp
    Loop
    Print #fp, App.EXEName
    BeginCriticalSession = fp
    
End Function

'-----------------------------------------------------
' �N���e�B�J���Z�b�V�����I������
'
'-----------------------------------------------------
Public Sub EndCriticalSession(pFp As Integer)

    Close #pFp
    
End Sub

'--------------------------------------------------------
'   �֐���  : GetFileExistence
'   �p�r    : �t�@�C�������݂��邩�ǂ������ׂ�
'   ����    : strPathName �t�@�C���E�f�B���g�N��(�p�X)��
'   �߂�l  : True �t�@�C���͑��݂���
'             False �t�@�C���͑��݂��Ȃ�
'--------------------------------------------------------
Public Function GetFileExistence(strPathName As String) As Boolean
    
    '�����̃T�C�Y���i�[/�t�@�C���ԍ����i�[
    Dim lngPNameSize As Long

    '�G���[�𖳌��ɂ��Ă���
    On Error Resume Next

    If strPathName = "" Then
        '�����̃t�@�C�����E�p�X�����Z�b�g����Ă��Ȃ�
        'Null���Z�b�g����
        GetFileExistence = ""
        '�֐��𔲂���
        Exit Function
    End If

    '�p�X���̍Ō�Ƀf�B���N�g���L��������ꍇ�͍폜
    If Right(strPathName, 1) = "\" Then
        
        '�p�X�̃T�C�Y-1���i�[
        lngPNameSize = Len(strPathName) - 1
        '�Ō�̈ꕶ������菜��
        strPathName = Left(strPathName, lngPNameSize)
    
    End If
    
    '�t�@�C�����J���āA�G���[���ǂ����m���߂�
    '���ݎg�p�\�ȃt�@�C���ԍ�������U��
    lngPNameSize = FreeFile
    
    '�ł́A�J��
    Open strPathName For Input As lngPNameSize
    
    '�G���[�ԍ��𒲂ׂ�B0�́u�t�@�C�����������v
    If Err = 0 Then
        '�u�t�@�C��������܂����v���Z�b�g
        GetFileExistence = True
    Else
        '�u�t�@�C���́A�Ȃ�������v���Z�b�g
        GetFileExistence = False
    End If
    
    Close lngPNameSize
    
    '�G���[�l��������
    Err = 0

End Function

'----------------------------------------------------
' ���͐�������(0�`9 & /)
'
'----------------------------------------------------
Public Sub gf_ChkDayInput(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0�`9�͓��͉�
        Case "/"
'            If InStr(F.ActiveControl, "/") = 0 Then
                Exit Sub            '--- /(�s���I�h)�͓��͉�
'            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpace�͓��͉�
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' ���͐�������(0�`9 & :)
'
'----------------------------------------------------
Public Sub gf_ChkTimeInput(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0�`9�͓��͉�
        Case ":"
'            If InStr(F.ActiveControl, ":") = 0 Then
                Exit Sub            '--- :(�s���I�h)�͓��͉�
'            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpace�͓��͉�
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' ���͐�������(0�`9 & / & :)
'
'----------------------------------------------------
Public Sub gf_ChkDateInput(F As Form, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0�`9�͓��͉�
        Case "/"
'            If InStr(F.ActiveControl, "/") = 0 Then
                Exit Sub            '--- /(�s���I�h)�͓��͉�
'            End If
        Case ":"
'            If InStr(F.ActiveControl, ":") = 0 Then
                Exit Sub            '--- :(�s���I�h)�͓��͉�
'            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpace�͓��͉�
    End Select
    pKeyAscii = 0
    
End Sub

Public Function gfNullChkStr(ByVal psInStr As Variant) As String

    If IsNull(psInStr) Then
        gfNullChkStr = ""
    Else
        gfNullChkStr = psInStr
    End If

End Function

Public Function gfNullChkStrTrim(ByVal psInStr As Variant) As String

    If IsNull(psInStr) Then
        gfNullChkStrTrim = ""
    Else
        gfNullChkStrTrim = Trim(psInStr)
    End If

End Function

'�ϊ��s�͂O��߂��H
Public Function gfNullZeroChkInt(ByVal psInStr As Variant) As String

    If IsNull(psInStr) Then
        gfNullZeroChkInt = ""
    Else
        gfNullZeroChkInt = Trim(psInStr)
    End If

End Function

'----------------------------------------------------
' ���͐�������(0�`9 & Period )
'
'----------------------------------------------------
Public Sub NumericPeriodVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0�`9�͓��͉�
        Case "."
            If InStr(ovsfGrd.EditText, ".") = 0 Then
                Exit Sub            '--- .(�s���I�h)�͂P�x�������͉�
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpace�͓��͉�
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' ���͐�������(0�`9 & '-')
'
'----------------------------------------------------
Public Sub NumericMinusVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0�`9�͓��͉�
        Case "-"
            If Len(ovsfGrd.EditText) = 0 _
            Or ovsfGrd.EditSelLength = Len(ovsfGrd.EditText) Then
                Exit Sub            '--- -(�}�C�i�X)�͂P�x�����A�擪�̂ݓ��͉�
            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpace�͓��͉�
    End Select
    pKeyAscii = 0
    
End Sub

'----------------------------------------------------
' ���͐�������(0�`9 & Period & '-')
'
'----------------------------------------------------
Public Sub NumericPeriodMinusVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer)

    Select Case Chr(pKeyAscii)
        Case "0" To "9"
            Exit Sub            '--- 0�`9�͓��͉�
        Case "."
            If InStr(ovsfGrd.EditText, ".") = 0 Then
                Exit Sub            '--- .(�s���I�h)�͂P�x�������͉�
            End If
        Case "-"
            If MinusCheckVsfGrd(ovsfGrd, pKeyAscii) Then Exit Sub
'            If Len(ovsfGrd.EditText) = 0 _
'            Or ovsfGrd.EditSelLength = Len(ovsfGrd.EditText) Then
'                Exit Sub            '--- -(�}�C�i�X)�͂P�x�����A�擪�̂ݓ��͉�
'            End If
        Case Chr(vbKeyBack)
            Exit Sub            '--- BackSpace�͓��͉�
    End Select
    pKeyAscii = 0

End Sub

Private Function MinusCheckVsfGrd(ovsfGrd As VSFlexGrid, pKeyAscii As Integer) As Boolean
    MinusCheckVsfGrd = False
    If Len(ovsfGrd.EditText) = 0 _
    Or ovsfGrd.EditSelLength = Len(ovsfGrd.EditText) Then
        MinusCheckVsfGrd = True
        Exit Function            '--- -(�}�C�i�X)�͂P�x�����A�擪�̂ݓ��͉�
    End If
End Function
