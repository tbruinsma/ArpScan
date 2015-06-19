Attribute VB_Name = "IP"
Public sAdapter() As String


Public Function Start(ByVal ip1 As String, ByVal SubnetMask As String, ByVal ip2 As String) As Boolean
    Dim home_network(4)
    Dim remote_network(4)
    Dim Mask(4)
    If ip1 = "" Or SubnetMask = "" Then
        Start = False
        Exit Function
    End If
    ' split the subnetmask into it's 4 octet
    '     s
    Mask(0) = convert_to_binary(Split(SubnetMask, ".", -1, 1)(0))
    Mask(1) = convert_to_binary(Split(SubnetMask, ".", -1, 1)(1))
    Mask(2) = convert_to_binary(Split(SubnetMask, ".", -1, 1)(2))
    Mask(3) = convert_to_binary(Split(SubnetMask, ".", -1, 1)(3))
    
    home_network(0) = convert_to_decimal(logical_and(convert_to_binary(Split(ip1, ".", -1, 1)(0)), Mask(0)))
    home_network(1) = convert_to_decimal(logical_and(convert_to_binary(Split(ip1, ".", -1, 1)(1)), Mask(1)))
    home_network(2) = convert_to_decimal(logical_and(convert_to_binary(Split(ip1, ".", -1, 1)(2)), Mask(2)))
    home_network(3) = convert_to_decimal(logical_and(convert_to_binary(Split(ip1, ".", -1, 1)(3)), Mask(3)))
    
    remote_network(0) = convert_to_decimal(logical_and(convert_to_binary(Split(ip2, ".", -1, 1)(0)), Mask(0)))
    remote_network(1) = convert_to_decimal(logical_and(convert_to_binary(Split(ip2, ".", -1, 1)(1)), Mask(1)))
    remote_network(2) = convert_to_decimal(logical_and(convert_to_binary(Split(ip2, ".", -1, 1)(2)), Mask(2)))
    remote_network(3) = convert_to_decimal(logical_and(convert_to_binary(Split(ip2, ".", -1, 1)(3)), Mask(3)))
   
    allowed = False


    For x = 0 To 3


        If home_network(x) <> remote_network(x) Then
            allowed = False
            Exit For
        Else
            allowed = True
        End If
    Next
    Start = allowed ' function returns true or false.
End Function


Public Function convert_to_binary(ByVal str As String) As String
    Dim i As Long, x As Long, Bin As String
    x = Val(str)
    For i = 7 To 0 Step -1

        If x And (2 ^ i) Then
            Bin = Bin + "1"
        Else
            Bin = Bin + "0"
        End If
    Next
    convert_to_binary = Bin
End Function

Public Function GetBinNetID(ByVal strip As String, ByVal StrSubnetMask As String) As String
  Dim pos As Integer, tempnetid As String, x As String, y As String, z As String
  pos = 1
  Do While pos <> Len(strip) + 1
    If Mid(strip, pos, 1) <> "." Then
      x = Mid(strip, pos, 1)
      y = Mid(StrSubnetMask, pos, 1)
      z = (CInt(x) * CInt(y))
      tempnetid = tempnetid & z
    Else
      tempnetid = tempnetid & "."
    End If
    pos = pos + 1
  Loop
  GetBinNetID = tempnetid
End Function

Public Function ConvertBinToIP(ByVal strBin As String) As String
  On Error Resume Next
  Dim pos As Integer, binarray, tempnetid As String, ix As Integer, x As Integer, y As Integer, z As String
  strBin = strBin & "."
  binarray = Split(strBin, ".")
  For ix = 0 To UBound(binarray) - 1
    z = z & convert_to_decimal(binarray(ix)) & "."
  Next ix
  ConvertBinToIP = Left(z, Len(z) - 1)
End Function

Public Function GetBits(strmask As String) As Single
  On Error Resume Next
  Dim tempdata, ix As Integer, pos As Integer, itemp As Single
  strmask = strmask & "."
  tempdata = Split(strmask, ".")
  For ix = 0 To UBound(tempdata) - 1
    Select Case tempdata(ix)
      Case "255"
        itemp = itemp + 8
      Case "128"
        itemp = itemp + 1
      Case "192"
        itemp = itemp + 2
      Case "224"
        itemp = itemp + 3
      Case "240"
        itemp = itemp + 4
      Case "248"
        itemp = itemp + 5
      Case "252"
        itemp = itemp + 6
      Case "254"
        itemp = itemp + 7
    End Select
  Next ix
  GetBits = itemp
End Function

Public Function GetRange(ByVal strmask As String) As Integer
  'uses the couchie method
  Dim tempdata, ix As Integer, itemp As Integer, ipclass As String
  Dim iRange As Integer
  
  strmask = strmask & "."
  tempdata = Split(strmask, ".")
  For ix = 0 To UBound(tempdata) - 1
    Select Case tempdata(ix)
      Case "128"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "192"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "224"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "240"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "248"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "252"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "254"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case Else
        iRange = 256
    End Select
  Next ix
  GetRange = iRange
End Function

Private Function convert_to_decimal(ByVal BinaryValue As String) As String
    
    Dim Bit As Integer
    Dim Value As Integer
    Dim Counter As Integer


    For Counter = Len(BinaryValue) To 1 Step -1
        Bit = Mid(BinaryValue, Counter, 1)


        If Bit = 1 Then
            Value = Value + 2 ^ (Len(BinaryValue) - Counter)
        End If
    Next
    convert_to_decimal = CStr(Value)
End Function

Private Function logical_and(ByVal str1 As String, ByVal str2 As String) As String
    Dim result As String
    For Counter = 1 To Len(str1)
        result = result & (Mid(str1, Counter, 1) And Mid(str2, Counter, 1))
    Next
    
    logical_and = result
End Function


Public Function GetNetStart(sIP As String)
Dim aryIP() As String
aryIP = Split(sIP, ".")

GetNetStart = aryIP(0) & "." & aryIP(1) & "." & aryIP(2) & "." & CStr(CInt(aryIP(3) + 1))
End Function

