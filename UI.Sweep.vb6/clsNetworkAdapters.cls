VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetworkAdapters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const MAX_HOSTNAME_LEN = 132
Const MAX_DOMAIN_NAME_LEN = 132
Const MAX_SCOPE_ID_LEN = 260
Const MAX_ADAPTER_NAME_LENGTH = 260
Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Const ERROR_BUFFER_OVERFLOW = 111
Const MIB_IF_TYPE_ETHERNET = 6
Const MIB_IF_TYPE_TOKENRING = 9
Const MIB_IF_TYPE_FDDI = 15
Const MIB_IF_TYPE_PPP = 23
Const MIB_IF_TYPE_LOOPBACK = 24
Const MIB_IF_TYPE_SLIP = 28

Private Type IP_ADDR_STRING
            Next As Long
            IPAddress As String * 16
            IpMask As String * 16
            Context As Long
End Type

Private Type IP_ADAPTER_INFO
            Next As Long
            ComboIndex As Long
            AdapterName As String * MAX_ADAPTER_NAME_LENGTH
            Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
            AddressLength As Long
            Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
            Index As Long
            Type As Long
            DhcpEnabled As Long
            CurrentIpAddress As Long
            IpAddressList As IP_ADDR_STRING
            GatewayList As IP_ADDR_STRING
            DHCPServer As IP_ADDR_STRING
            HaveWins As Byte
            PrimaryWinsServer As IP_ADDR_STRING
            SecondaryWinsServer As IP_ADDR_STRING
            LeaseObtained As Long
            LeaseExpires As Long
End Type

Private Type FIXED_INFO
            HostName As String * MAX_HOSTNAME_LEN
            DomainName As String * MAX_DOMAIN_NAME_LEN
            CurrentDnsServer As Long
            DnsServerList As IP_ADDR_STRING
            NodeType As Long
            ScopeId  As String * MAX_SCOPE_ID_LEN
            EnableRouting As Long
            EnableProxy As Long
            EnableDns As Long
End Type

Private Type NetAdapter
    NodeType As String
    DNSServers As String
    HostName As String
    Name As String
    Description As String
    DHCP As Boolean
    IPAddress As String
    SubnetMask As String
    Gateway As String
    PrimaryWINS As String
    SecondaryWINS As String
    HWAddress As String
    DHCPServer As String
End Type

Private Declare Function GetNetworkParams Lib "IPHlpApi.dll" (FixedInfo As Any, pOutBufLen As Long) As Long
Private Declare Function GetAdaptersInfo Lib "IPHlpApi.dll" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
       
       

'local variable(s) to hold property value(s)
Private mvarCount As Integer 'local copy
Private mvarDomain As String
Private mvarHostName As String
Private mvarNodeType As String
Private mvarDNSCount As Integer
Private mvarDNSServer As String
Private mvarIPRouting As Boolean
Private mvarWINSProxy As Boolean
Private mvarNetBIODNS As Boolean
    
Private aryAdapters() As NetAdapter


Public Property Let Count(ByVal vData As Integer)
    mvarCount = vData
End Property


Public Property Get Count() As Integer
    Count = mvarCount
End Property

Public Property Get NodeType() As String
    NodeType = mvarNodeType
End Property

Public Property Get HostName() As String
    HostName = mvarHostName
End Property

Public Property Get Domain() As String
    Domain = mvarDomain
End Property

Public Property Get DNSCount() As Integer
    DNSCount = mvarDNSCount
End Property

Public Property Get IPRouting() As Boolean
    IPRouting = mvarIPRouting
End Property

Public Property Get WINSProxy() As Boolean
    WINSProxy = mvarWINSProxy
End Property

Public Property Get NetBIODNS() As Boolean
    NetBIODNS = mvarNetBIODNS
End Property

Public Property Get AdapterName(ID As Integer) As String
    AdapterName = aryAdapters(ID).Name
End Property

Public Property Get AdapterDescription(ID As Integer) As String
    AdapterDescription = aryAdapters(ID).Description
End Property

Public Property Get AdapterDHCP(ID As Integer) As Boolean
    AdapterDHCP = aryAdapters(ID).DHCP
End Property

Public Property Get AdapterDHCPServer(ID As Integer) As String
    AdapterDHCPServer = aryAdapters(ID).DHCPServer
End Property

Public Property Get AdapterDNSServers(ID As Integer) As String
    AdapterDNSServers = aryAdapters(ID).DNSServers
End Property

Public Property Get AdapterGateway(ID As Integer) As String
    AdapterGateway = aryAdapters(ID).Gateway
End Property

Public Property Get AdapterMAC(ID As Integer) As String
    AdapterMAC = aryAdapters(ID).HWAddress
End Property

Public Property Get AdapterIPAddress(ID As Integer) As String
    AdapterIPAddress = aryAdapters(ID).IPAddress
End Property

Public Property Get AdapterNodeType(ID As Integer) As String
    AdapterNodeType = aryAdapters(ID).NodeType
End Property

Public Property Get AdapterPrimaryWINS(ID As Integer) As String
    AdapterPrimaryWINS = aryAdapters(ID).PrimaryWINS
End Property

Public Property Get AdapterSecondaryWINS(ID As Integer) As String
    AdapterSecondaryWINS = aryAdapters(ID).SecondaryWINS
End Property

Public Property Get AdapterSubnetMask(ID As Integer) As String
    AdapterSubnetMask = aryAdapters(ID).SubnetMask
End Property

Public Sub Refresh()
    Call Start
End Sub

Private Sub Class_Initialize()
    Call Start
End Sub

Public Function MaskFromIP(IP As String) As String
    Dim x As Integer
    For x = 0 To UBound(aryAdapters)
        If aryAdapters(x).IPAddress = IP Then
            MaskFromIP = aryAdapters(x).SubnetMask
            Exit Function
        End If
    Next
End Function

'converts a Long  to a string
Private Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Private Sub Start()
    Dim error As Long
    Dim FixedInfoSize As Long
    Dim AdapterInfoSize As Long
    Dim i As Integer
    Dim PhysicalAddress  As String
    Dim NewTime As Date
    Dim AdapterInfo As IP_ADAPTER_INFO
    Dim AddrStr As IP_ADDR_STRING
    Dim FixedInfo As FIXED_INFO
    Dim Buffer As IP_ADDR_STRING
    Dim pAddrStr As Long
    Dim pAdapt As Long
    Dim Buffer2 As IP_ADAPTER_INFO
    Dim FixedInfoBuffer() As Byte
    Dim AdapterInfoBuffer() As Byte
    Dim iCount As Integer

    ' Get the main IP configuration information for this machine
    ' using a FIXED_INFO structure.
    FixedInfoSize = 0
    error = GetNetworkParams(ByVal 0&, FixedInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           Debug.Print "GetNetworkParams sizing failed with error " & error
           Exit Sub
        End If
    End If
    ReDim FixedInfoBuffer(FixedInfoSize - 1)

    error = GetNetworkParams(FixedInfoBuffer(0), FixedInfoSize)
    If error = 0 Then
            CopyMemory FixedInfo, FixedInfoBuffer(0), FixedInfoSize
            
            mvarHostName = TrimNullStr(FixedInfo.HostName)
            mvarDNSServers = TrimNullStr(FixedInfo.DnsServerList.IPAddress)
            
            pAddrStr = FixedInfo.DnsServerList.Next
            Do While pAddrStr <> 0
                  CopyMemory Buffer, ByVal pAddrStr, LenB(Buffer)
                  mvarDNSServers = mvarDNSServers & "," & TrimNullStr(Buffer.IPAddress)
                  pAddrStr = Buffer.Next
            Loop

            Select Case FixedInfo.NodeType
                       Case 1
                                  mvarNodeType = "Broadcast"
                       Case 2
                                  mvarNodeType = "Peer to peer"
                       Case 4
                                  mvarNodeType = "Mixed"
                       Case 8
                                  mvarNodeType = "Hybrid"
                       Case Else
                                  mvarNodeType = "Unknown node type"
            End Select

            'Debug.Print "NetBIOS Scope ID:  " & FixedInfo.ScopeId
            mvarIPRouting = FixedInfo.EnableRouting
            mvarWINSProxy = FixedInfo.EnableProxy
            mvarNetBIODDNS = FixedInfo.EnableDns
    Else
            Debug.Print "GetNetworkParams failed with error " & error
            Exit Sub
    End If

    ' Enumerate all of the adapter specific information using the
    ' IP_ADAPTER_INFO structure.
    ' Note:  IP_ADAPTER_INFO contains a linked list of adapter entries.

    AdapterInfoSize = 0
    error = GetAdaptersInfo(ByVal 0&, AdapterInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           Debug.Print "GetAdaptersInfo sizing failed with error " & error
           Exit Sub
        End If
    End If
    ReDim AdapterInfoBuffer(AdapterInfoSize - 1)

    ' Get actual adapter information
    iCount = 0
    error = GetAdaptersInfo(AdapterInfoBuffer(0), AdapterInfoSize)
    If error <> 0 Then
       Debug.Print "GetAdaptersInfo failed with error " & error
       Exit Sub
    End If

    ' Allocate memory
     CopyMemory AdapterInfo, AdapterInfoBuffer(0), AdapterInfoSize
    pAdapt = AdapterInfo.Next

    Do
     ReDim Preserve aryAdapters(iCount)
     CopyMemory Buffer2, AdapterInfo, AdapterInfoSize
       Select Case Buffer2.Type
              Case MIB_IF_TYPE_ETHERNET
                   aryAdapters(iCount).Name = "Ethernet adapter"
              Case MIB_IF_TYPE_TOKENRING
                   aryAdapters(iCount).Name = "Token Ring adapter"
              Case MIB_IF_TYPE_FDDI
                   aryAdapters(iCount).Name = "FDDI adapter"
              Case MIB_IF_TYPE_PPP
                   aryAdapters(iCount).Name = "PPP adapter"
              Case MIB_IF_TYPE_LOOPBACK
                   aryAdapters(iCount).Name = "Loopback adapter"
              Case MIB_IF_TYPE_SLIP
                   aryAdapters(iCount).Name = "Slip adapter"
              Case Else
                   aryAdapters(iCount).Name = "Other adapter"
       End Select
       aryAdapters(iCount).Description = Buffer2.Description

       PhysicalAddress = ""
       For i = 0 To Buffer2.AddressLength - 1
           PhysicalAddress = PhysicalAddress & Hex(Buffer2.Address(i))
           If i < Buffer2.AddressLength - 1 Then
              PhysicalAddress = PhysicalAddress & "-"
           End If
       Next
       aryAdapters(iCount).HWAddress = PhysicalAddress
       aryAdapters(iCount).DHCP = Buffer2.DhcpEnabled

       aryAdapters(iCount).IPAddress = TrimNullStr(Buffer2.IpAddressList.IPAddress)
       aryAdapters(iCount).SubnetMask = TrimNullStr(Buffer2.IpAddressList.IpMask)
       
       pAddrStr = Buffer2.IpAddressList.Next
       Do While pAddrStr <> 0
          CopyMemory Buffer, Buffer2.IpAddressList, LenB(Buffer)
          aryAdapters(iCount).IPAddress = aryAdapters(iCount).IPAddress & "," & TrimNullStr(Buffer.IPAddress)
          aryAdapters(iCount).SubnetMask = aryAdapters(iCount).SubnetMask & "," & TrimNullStr(Buffer.IpMask)
          pAddrStr = Buffer.Next
          If pAddrStr <> 0 Then
             CopyMemory Buffer2.IpAddressList, ByVal pAddrStr, _
                        LenB(Buffer2.IpAddressList)
          End If
       Loop

       aryAdapters(iCount).Gateway = TrimNullStr(Buffer2.GatewayList.IPAddress)
       
       pAddrStr = Buffer2.GatewayList.Next
       Do While pAddrStr <> 0
          CopyMemory Buffer, Buffer2.GatewayList, LenB(Buffer)
          aryAdapters(iCount).Gateway = aryAdapters(iCount).Gateway & "," & TrimNullStr(Buffer.IPAddress)
          pAddrStr = Buffer.Next
          If pAddrStr <> 0 Then
             CopyMemory Buffer2.GatewayList, ByVal pAddrStr, _
                        LenB(Buffer2.GatewayList)
          End If
       Loop

       aryAdapters(iCount).DHCPServer = TrimNullStr(Buffer2.DHCPServer.IPAddress)
       aryAdapters(iCount).PrimaryWINS = TrimNullStr(Buffer2.PrimaryWinsServer.IPAddress)
       aryAdapters(iCount).SecondaryWINS = TrimNullStr(Buffer2.SecondaryWinsServer.IPAddress)

       ' Display time.
       'NewTime = DateAdd("s", Buffer2.LeaseObtained, #1/1/1970#)
       'Debug.Print "Lease Obtained: " & _
       '       CStr(Format(NewTime, "dddd, mmm d hh:mm:ss yyyy"))

       'NewTime = DateAdd("s", Buffer2.LeaseExpires, #1/1/1970#)
       'Debug.Print "Lease Expires :  " & _
       '       CStr(Format(NewTime, "dddd, mmm d hh:mm:ss yyyy"))
       pAdapt = Buffer2.Next
       If pAdapt <> 0 Then
           CopyMemory AdapterInfo, ByVal pAdapt, AdapterInfoSize
        End If
        iCount = iCount + 1
      Loop Until pAdapt = 0
      
      mvarCount = iCount
End Sub

Private Function TrimNullStr(pszString As String) As String
    Dim b As String, cs As String, bl As Integer
    Dim c As Integer, ca As Integer, cb As Integer, cf As Boolean
    '
    b$ = Trim$(pszString)
    bl% = Len(b$) ' Get length of String
    cb% = bl% + 1
    '
    ' recognized valid characters
    cs$ = Chr$(126)


    For c% = 32 To 96: cs$ = cs$ & Chr$(c%): Next


        For c% = 160 To 223: cs$ = cs$ & Chr$(c%): Next
            '


            For c% = 1 To bl%


                If Not cf And InStr(UCase$(cs$), UCase$(Mid$(b$, c%, 1))) Then
                    ca% = c% 'First character after preceding Nullstring
                    cf = True
                End If


                If cf And InStr(UCase$(cs$), UCase$(Mid$(b$, c%, 1))) = 0 Then
                    cb% = c% 'Terminating Nullstring after last character
                    Exit For
                End If
            Next c%
            If cf Then TrimNullStr = Trim$(Mid$(b$, ca%, (cb% - ca%)))
            '
        End Function

