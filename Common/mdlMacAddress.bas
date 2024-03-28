Attribute VB_Name = "mdlMacAddress"
Option Explicit

Private Declare Sub memmoveAL Lib "kernel32" Alias "RtlMoveMemory" (dst As Any, ByVal src As Long, ByVal num As Long)
Private Declare Sub memset Lib "kernel32" Alias "RtlFillMemory" (dst As Any, ByVal length As Long, ByVal fill As Byte)

' Global Memory Flags
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Type T_ADAPTER_STATUS
    adapter_address(5)  As Byte
    rev_major           As Byte
    reserved0           As Byte
    adapter_type        As Byte
    rev_minor           As Byte
    duration            As Long
    frmr_recv           As Long
    frmr_xmit           As Long
    iframe_recv_err     As Long
    xmit_aborts         As Long
    xmit_success        As Long
    recv_success        As Long
    iframe_xmit_err     As Long
    recv_buff_unavail   As Long
    t1_timeouts         As Long
    ti_timeouts         As Long
    reserved1           As Long
    free_ncbs           As Long
    max_cfg_ncbs        As Long
    max_ncbs            As Long
    xmit_buf_unavail    As Long
    max_dgram_size      As Long
    pending_sess        As Long
    max_cfg_sess        As Long
    max_sess            As Long
    max_sess_pkt_size   As Long
    name_count          As Long
End Type

Private Const NCBNAMSZ  As Long = 16    'absolute length of a net name

Type T_NAME_BUFFER
    name_(NCBNAMSZ - 1) As Byte
    name_num            As Byte
    name_flags          As Byte
End Type

Type T_ASTAT
    adapt           As T_ADAPTER_STATUS
    NameBuff(29)    As T_NAME_BUFFER
End Type

Type T_NCB
    ncb_command                 As Byte
    ncb_retcode                 As Byte
    ncb_lsn                     As Byte
    ncb_num                     As Byte
    ncb_buffer                  As Long
    ncb_length                  As Integer
    ncb_callname(NCBNAMSZ - 1)  As Byte
    ncb_name(NCBNAMSZ - 1)      As Byte
    ncb_rto                     As Byte
    ncb_sto                     As Byte
    ncb_post                    As Long
    ncb_lana_num                As Byte
    ncb_cmd_cplt                As Byte
    ncb_reserve(9)              As Byte
    ncb_event                   As Long
End Type

Private Const MAX_LANA      As Long = 254   'lana's in range 0 to MAX_LANA inclusive

Type LANA_ENUM
    length          As Byte
    lana(MAX_LANA)  As Byte
End Type

'NCB Command codes
Private Const NCBRESET  As Byte = &H32
Private Const NCBASTAT  As Byte = &H33
Private Const NCBENUM   As Byte = &H37
Private Declare Function Netbios Lib "netapi32.dll" (pncb As Any) As Byte

Public Function GetMacAddress() As String

    Dim sRet        As String
    sRet = ""

    Dim Ncb         As T_NCB
    Dim uRetCode    As Byte
    Dim lenum       As LANA_ENUM
    Dim nNcbLen     As Long
    nNcbLen = (4 * 3) + 2 + 11

    memset Ncb, 0, nNcbLen
    With Ncb
        .ncb_command = NCBENUM
        .ncb_length = 256
        .ncb_buffer = GlobalAllocPtr(GHND, .ncb_length)
    End With
    uRetCode = Netbios(Ncb)
    If uRetCode = 0 Then
        memmoveAL lenum, Ncb.ncb_buffer, Ncb.ncb_length

        Dim Adapter     As T_ASTAT
        Dim nAdpLen     As Long
        Dim pAdpAdr     As Long
        nAdpLen = ((22 * 4) + 5) + (18 * 30)
        pAdpAdr = GlobalAllocPtr(GHND, nAdpLen)

        If lenum.length > 0 Then
            memset Ncb, 0, nNcbLen
            With Ncb
                .ncb_command = NCBRESET
                .ncb_lana_num = lenum.lana(0)
            End With
            uRetCode = Netbios(Ncb)
            If uRetCode = 0 Then
                memset Ncb, 0, nNcbLen
                With Ncb
                    .ncb_command = NCBASTAT
                    .ncb_lana_num = lenum.lana(0)
                    .ncb_callname(0) = Asc("*")
                    memset .ncb_callname(1), Asc(" "), 15
                    .ncb_name(0) = 0
                    .ncb_buffer = pAdpAdr
                    .ncb_length = nAdpLen
                End With
                uRetCode = Netbios(Ncb)
                If uRetCode = 0 Then
                    memmoveAL Adapter, pAdpAdr, nAdpLen
                    sRet = Right("00" & Hex(Adapter.adapt.adapter_address(0)), 2) & _
                           Right("00" & Hex(Adapter.adapt.adapter_address(1)), 2) & _
                           Right("00" & Hex(Adapter.adapt.adapter_address(2)), 2) & _
                           Right("00" & Hex(Adapter.adapt.adapter_address(3)), 2) & _
                           Right("00" & Hex(Adapter.adapt.adapter_address(4)), 2) & _
                           Right("00" & Hex(Adapter.adapt.adapter_address(5)), 2)
                End If
            End If
        End If
        GlobalFreePtr pAdpAdr
    End If
    GlobalFreePtr Ncb.ncb_buffer

    GetMacAddress = sRet

End Function

Private Function GlobalAllocPtr(ByVal flags As Long, ByVal cb As Long) As Long
    GlobalAllocPtr = GlobalLock(GlobalAlloc(flags, (cb)))
End Function

Private Function GlobalPtrHandle(ByVal lp As Long) As Long
    GlobalPtrHandle = GlobalHandle(lp)
End Function

Private Function GlobalUnlockPtr(ByVal lp As Long) As Long
    GlobalUnlockPtr = GlobalUnlock(GlobalPtrHandle(lp))
End Function

Private Sub GlobalFreePtr(ByVal lp As Long)
    GlobalUnlockPtr lp
    GlobalFree GlobalPtrHandle(lp)
End Sub


