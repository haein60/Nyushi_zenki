Attribute VB_Name = "mAdptInf"
'=========================================================================
'   み～くんパパのAPIサンプル
'-------------------------------------------------------------------------
'   [IPHelper] Ethernet アダプタ情報の取得
'-------------------------------------------------------------------------
'   [作成日]    2000年03月25日
'-------------------------------------------------------------------------
'   [動作確認環境]
'       Windows 98 Second Edition
'       Windows 2000 Professional
'       ※ すべて VB 6.0 SP3 で実行
'-------------------------------------------------------------------------
'   [NOTE]
'       Windows 98, Windows 2000 以降の環境でのみ動作します。
'       (Windows 95, NT4では動作しません)
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

'戻り値：アダプタ情報の数、マイナスはエラー
Public Function gfLoadAdptData(psRetMsg As String) As Long

Dim bytAdaptInfo()  As Byte
Dim lngBufLen       As Long
Dim lngRet          As Long
Dim lngPtr          As Long
Dim i               As Integer

    ' データの取得バッファの初期化
    ReDim udtIPAdaptInfo(0)
    ReDim bytAdaptInfo(0)

    ' アダプタ情報を取得するために必要なバッファサイズを取得する
    lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)

    Select Case lngRet
    Case ERROR_BUFFER_OVERFLOW  ' バッファサイズが小さい
        ' 取得したバッファサイズより必要なバッファを確保
        ReDim bytAdaptInfo(lngBufLen)
        ' アダプタ情報を取得する
        lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)
        If lngRet = NO_ERROR Then
            ' 取得したアダプタ情報をユーザー定義型変数へコピー
            CopyMemory udtIPAdaptInfo(0), bytAdaptInfo(0), LenB(udtIPAdaptInfo(0))
            lngPtr = udtIPAdaptInfo(0).pNext
            i = 0
            Do While Not lngPtr = 0 ' 配列の次のアドレスへのポインタをチェックする
                i = i + 1
                ReDim Preserve udtIPAdaptInfo(i)
                CopyMemory udtIPAdaptInfo(i), ByVal lngPtr, LenB(udtIPAdaptInfo(0))
                lngPtr = udtIPAdaptInfo(i).pNext
            Loop
        End If
        gfLoadAdptData = UBound(udtIPAdaptInfo) + 1
    ' エラー表示
    Case ERROR_INVALID_PARAMETER
        psRetMsg = "出力バッファへのアクセスが許可されていないか、パラメータが不正です"
        gfLoadAdptData = -1
    Case ERROR_NO_DATA
        psRetMsg = "このコンピュータにはアダプタ情報が存在しません"
        gfLoadAdptData = -2
    Case ERROR_NOT_SUPPORTED
        psRetMsg = "このバージョンのOSでこの機能はサポートされていません"
        gfLoadAdptData = -3
    Case Else
        psRetMsg = "不明なエラーです。" & str(lngRet)
        gfLoadAdptData = -4
    End Select

End Function

Public Function getMacAddr(plIndex As Long) As String

Dim sMacAddr As String

On Error GoTo ErrProc

    getMacAddr = ""

    ' アダプタアドレス
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

    ' アダプタアドレス
    sMacAddr = ByteToMACAddr(udtIPAdaptInfo(plIndex).Address, udtIPAdaptInfo(plIndex).AddressLength)
    ' IPアドレス
    sIPAddr = ByteToStr(udtIPAdaptInfo(plIndex).IpAddressList.IpAddress.Addr)
    ' サブネットマスク
    sSubMask = ByteToStr(udtIPAdaptInfo(plIndex).IpAddressList.IpMask.Addr)
    ' ゲートウェイ
    sGW = ByteToStr(udtIPAdaptInfo(plIndex).GatewayList.IpAddress.Addr)
    ' DHCP
    sDHCP = ByteToStr(udtIPAdaptInfo(plIndex).DhcpServer.IpAddress.Addr)
    ' プライマリWINS
    sWINS1 = ByteToStr(udtIPAdaptInfo(plIndex).PrimaryWinsServer.IpAddress.Addr)
    ' セカンダリWINS
    sWINS2 = ByteToStr(udtIPAdaptInfo(plIndex).SecondaryWinsServer.IpAddress.Addr)
    
    If udtIPAdaptInfo(plIndex).bDhcpEnabled = 0 Or _
       sIPAddr = "0.0.0.0" Then
        sLeaseObtain = ""
        sLeaseExpire = ""
    Else
    ' DHCPリース取得日
        sLeaseObtain = Format$(DateAdd("s", udtIPAdaptInfo(plIndex).LeaseObtained, #1/1/1970#), "yyyy/mm/dd hh:nn:ss")
    ' DHCPリース期限
        sLeaseExpire = Format$(DateAdd("s", udtIPAdaptInfo(plIndex).LeaseExpires, #1/1/1970#), "yyyy/mm/dd hh:nn:ss")
    End If

End Sub

' バイト配列よりアダプタのアドレスを文字列型に生成する
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

' バイト配列よりNULL終端までのデータで文字列型を生成
Private Function ByteToStr(bytDest As Variant) As String
    Dim strBuf  As String
    strBuf = StrConv(bytDest, vbUnicode)
    ByteToStr = Left$(strBuf, InStr(1, strBuf, vbNullChar) - 1)
End Function

