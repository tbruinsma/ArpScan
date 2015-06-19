VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdapters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarCount As Integer 'local copy
Private aryAdapters() As String

Const MAX_IP = 5   'To make a buffer... i dont think you have more than 5 ip on your pc..

Private Type IPINFO
     dwAddr As Long   ' IP address
    dwIndex As Long '  interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Private Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO  'array of IP address entries
End Type

Private Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long


Public Property Let Count(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Count = 5
    mvarCount = vData
End Property


Public Property Get Count() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Count
    Count = mvarCount
End Property



Public Function Mask(Interface As Integer) As String
    If Interface <= UBound(aryAdapters, 2) Then Mask = aryAdapters(2, Interface)
End Function

Public Function MaskFromIP(IP As String) As String
    Dim x As Integer
    For x = 0 To UBound(aryAdapters, 2)
        If aryAdapters(1, x) = IP Then
            MaskFromIP = aryAdapters(2, x)
            Exit Function
        End If
    Next
End Function

Public Function IP(Interface As Integer) As String
    If Interface <= UBound(aryAdapters, 2) Then IP = aryAdapters(1, Interface)
End Function

Public Sub Refresh()
    Call Start
End Sub

Private Sub Class_Initialize()
    Call Start
End Sub

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
Dim Ret As Long
Dim bBytes() As Byte
Dim Listing As MIB_IPADDRTABLE
Dim iBound As Integer

    

On Error GoTo END1
    GetIpAddrTable ByVal 0&, Ret, True

    If Ret <= 0 Then Exit Sub
    ReDim bBytes(0 To Ret - 1) As Byte
    'retrieve the data
    GetIpAddrTable bBytes(0), Ret, False
      
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    mvarCount = Listing.dEntrys
    ReDim Preserve aryAdapters(2, Listing.dEntrys)
    
    For tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
        CopyMemory Listing.mIPInfo(tel), bBytes(4 + (tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(tel))
        aryAdapters(1, tel) = ConvertAddressToString(Listing.mIPInfo(tel).dwAddr)
        aryAdapters(2, tel) = ConvertAddressToString(Listing.mIPInfo(tel).dwMask)
    Next

Exit Sub
END1:
'MsgBox "ERROR RESOLVING ADAPTERS"
End Sub
