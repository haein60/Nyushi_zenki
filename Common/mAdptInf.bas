Attribute VB_Name = "mAdptInf"
'=========================================================================
'   �݁`����p�p��API�T���v��
'-------------------------------------------------------------------------
'   [IPHelper] Ethernet �A�_�v�^���̎擾
'-------------------------------------------------------------------------
'   [�쐬��]    2000�N03��25��
'-------------------------------------------------------------------------
'   [����m�F��]
'       Windows 98 Second Edition
'       Windows 2000 Professional
'       �� ���ׂ� VB 6.0 SP3 �Ŏ��s
'-------------------------------------------------------------------------
'   [NOTE]
'       Windows 98, Windows 2000 �ȍ~�̊��ł̂ݓ��삵�܂��B
'       (Windows 95, NT4�ł͓��삵�܂���)
'=========================================================================
Option Explicit

Public Const NO_ERROR = 0
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const ERROR_INVALID_PARAMETER = 87
Public Const ERROR_NO_DATA = 232
Public Const ERROR_NOT_SUPPORTED = 50

Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 128  '// arb.
Public Const MAX_ADAPTER_NAME_LENGTH = 256         '// arb.
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8        '// arb.

Public Type IP_ADDRESS_STRING
    Addr(15)        As Byte
End Type

Public Type IP_ADDR_STRING
    pNext           As Long
    IpAddress       As IP_ADDRESS_STRING
    IpMask          As IP_ADDRESS_STRING
    Context         As Long
End Type

Public Type IP_ADAPTER_INFO
    pNext                   As Long
    ComboIndex              As Long
    AdapterName(MAX_ADAPTER_NAME_LENGTH + 3)          As Byte
    Description(MAX_ADAPTER_DESCRIPTION_LENGTH + 3)   As Byte
    AddressLength           As Long
    Address(MAX_ADAPTER_ADDRESS_LENGTH - 1)           As Byte
    dwIndex                 As Long
    uType                   As Long
    bDhcpEnabled            As Long
    pCurrentIpAddress       As Long
    IpAddressList           As IP_ADDR_STRING
    GatewayList             As IP_ADDR_STRING
    DhcpServer              As IP_ADDR_STRING
    bHaveWins               As Long
    PrimaryWinsServer       As IP_ADDR_STRING
    SecondaryWinsServer     As IP_ADDR_STRING
    LeaseObtained           As Long 'time_t
    LeaseExpires            As Long 'time_t
End Type

Public Declare Function GetAdaptersInfo Lib "IPHLPAPI.DLL" ( _
    pAdapterInfo As Byte, _
    pOutBufLen As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)


''''2022.01.24 add jhi
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private udtIPAdaptInfo()    As IP_ADAPTER_INFO

'�߂�l�F�A�_�v�^���̐��A�}�C�i�X�̓G���[
Public Function gfLoadAdptData(psRetMsg As String) As Long

Dim bytAdaptInfo()  As Byte
Dim lngBufLen       As Long
Dim lngRet          As Long
Dim lngPtr          As Long
Dim i               As Integer

    ' �f�[�^�̎擾�o�b�t�@�̏�����
    ReDim udtIPAdaptInfo(0)
    ReDim bytAdaptInfo(0)

    ' �A�_�v�^�����擾���邽�߂ɕK�v�ȃo�b�t�@�T�C�Y���擾����
    lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)

    Select Case lngRet
    Case ERROR_BUFFER_OVERFLOW  ' �o�b�t�@�T�C�Y��������
        ' �擾�����o�b�t�@�T�C�Y���K�v�ȃo�b�t�@���m��
        ReDim bytAdaptInfo(lngBufLen)
        ' �A�_�v�^�����擾����
        lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)
        If lngRet = NO_ERROR Then
            ' �擾�����A�_�v�^�������[�U�[��`�^�ϐ��փR�s�[
            CopyMemory udtIPAdaptInfo(0), bytAdaptInfo(0), LenB(udtIPAdaptInfo(0))
            lngPtr = udtIPAdaptInfo(0).pNext
            i = 0
            Do While Not lngPtr = 0 ' �z��̎��̃A�h���X�ւ̃|�C���^���`�F�b�N����
                i = i + 1
                ReDim Preserve udtIPAdaptInfo(i)
                CopyMemory udtIPAdaptInfo(i), ByVal lngPtr, LenB(udtIPAdaptInfo(0))
                lngPtr = udtIPAdaptInfo(i).pNext
            Loop
        End If
        gfLoadAdptData = UBound(udtIPAdaptInfo) + 1
    ' �G���[�\��
    Case ERROR_INVALID_PARAMETER
        psRetMsg = "�o�̓o�b�t�@�ւ̃A�N�Z�X��������Ă��Ȃ����A�p�����[�^���s���ł�"
        gfLoadAdptData = -1
    Case ERROR_NO_DATA
        psRetMsg = "���̃R���s���[�^�ɂ̓A�_�v�^��񂪑��݂��܂���"
        gfLoadAdptData = -2
    Case ERROR_NOT_SUPPORTED
        psRetMsg = "���̃o�[�W������OS�ł��̋@�\�̓T�|�[�g����Ă��܂���"
        gfLoadAdptData = -3
    Case Else
        psRetMsg = "�s���ȃG���[�ł��B" & str(lngRet)
        gfLoadAdptData = -4
    End Select

End Function

Public Function getMacAddr(plIndex As Long) As String

Dim sMacAddr As String

On Error GoTo ErrProc

    getMacAddr = ""

    ' �A�_�v�^�A�h���X
    sMacAddr = ByteToMACAddr(udtIPAdaptInfo(plIndex).Address, udtIPAdaptInfo(plIndex).AddressLength)

    getMacAddr = sMacAddr

Exit Function

ErrProc:

End Function

Public Sub getAdptData(plIndex As Long)

Dim sMacAddr As String
Dim sIPAddr As String
Dim sSubMask As String
Dim sGW As String
Dim sDHCP As String
Dim sWINS1 As String
Dim sWINS2 As String
Dim sLeaseObtain As String
Dim sLeaseExpire As String

    ' �A�_�v�^�A�h���X
    sMacAddr = ByteToMACAddr(udtIPAdaptInfo(plIndex).Address, udtIPAdaptInfo(plIndex).AddressLength)
    ' IP�A�h���X
    sIPAddr = ByteToStr(udtIPAdaptInfo(plIndex).IpAddressList.IpAddress.Addr)
    ' �T�u�l�b�g�}�X�N
    sSubMask = ByteToStr(udtIPAdaptInfo(plIndex).IpAddressList.IpMask.Addr)
    ' �Q�[�g�E�F�C
    sGW = ByteToStr(udtIPAdaptInfo(plIndex).GatewayList.IpAddress.Addr)
    ' DHCP
    sDHCP = ByteToStr(udtIPAdaptInfo(plIndex).DhcpServer.IpAddress.Addr)
    ' �v���C�}��WINS
    sWINS1 = ByteToStr(udtIPAdaptInfo(plIndex).PrimaryWinsServer.IpAddress.Addr)
    ' �Z�J���_��WINS
    sWINS2 = ByteToStr(udtIPAdaptInfo(plIndex).SecondaryWinsServer.IpAddress.Addr)
    
    If udtIPAdaptInfo(plIndex).bDhcpEnabled = 0 Or _
       sIPAddr = "0.0.0.0" Then
        sLeaseObtain = ""
        sLeaseExpire = ""
    Else
    ' DHCP���[�X�擾��
        sLeaseObtain = Format$(DateAdd("s", udtIPAdaptInfo(plIndex).LeaseObtained, #1/1/1970#), "yyyy/mm/dd hh:nn:ss")
    ' DHCP���[�X����
        sLeaseExpire = Format$(DateAdd("s", udtIPAdaptInfo(plIndex).LeaseExpires, #1/1/1970#), "yyyy/mm/dd hh:nn:ss")
    End If

End Sub

' �o�C�g�z����A�_�v�^�̃A�h���X�𕶎���^�ɐ�������
Private Function ByteToMACAddr(bytDest As Variant, BufLen As Long) As String
    Dim strBuf  As String
    Dim i       As Integer
    On Error Resume Next
    For i = 0 To BufLen - 1
        If i > 0 Then strBuf = strBuf & "-"
        strBuf = strBuf & Right$("00" & Hex$(bytDest(i)), 2)
    Next
    ByteToMACAddr = strBuf
End Function

' �o�C�g�z����NULL�I�[�܂ł̃f�[�^�ŕ�����^�𐶐�
Private Function ByteToStr(bytDest As Variant) As String
    Dim strBuf  As String
    strBuf = StrConv(bytDest, vbUnicode)
    ByteToStr = Left$(strBuf, InStr(1, strBuf, vbNullChar) - 1)
End Function

