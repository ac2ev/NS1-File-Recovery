Attribute VB_Name = "BytesToNumber"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'The WideCharToMultiByte function maps a wide-character string to a new character string.
'The function is faster when both lpDefaultChar and lpUsedDefaultChar are NULL.

'CodePage
Private Const CP_ACP = 0 'ANSI
Private Const CP_MACCP = 2 'Mac
Private Const CP_OEMCP = 1 'OEM
Private Const CP_UTF7 = 65000
Private Const CP_UTF8 = 65001

'dwFlags
Private Const WC_NO_BEST_FIT_CHARS = &H400
Private Const WC_COMPOSITECHECK = &H200
Private Const WC_DISCARDNS = &H10
Private Const WC_SEPCHARS = &H20 'Default
Private Const WC_DEFAULTCHAR = &H40

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
                                                    ByVal dwFlags As Long, _
                                                    ByVal lpWideCharStr As Long, _
                                                    ByVal cchWideChar As Long, _
                                                    ByVal lpMultiByteStr As Long, _
                                                    ByVal cbMultiByte As Long, _
                                                    ByVal lpDefaultChar As Long, _
                                                    ByVal lpUsedDefaultChar As Long) As Long
                                                    
Private Type DOUBLE_T
    d As Double             ' 8 bytes
End Type
Dim mtDouble As DOUBLE_T

Private Type TWOLONGS_T
   al(0 To 1) As Long      ' 2 * 4 bytes
End Type
Private mtLongs As TWOLONGS_T

Private Type FOURINTS_T
   ai(0 To 3) As Integer   ' 4 * 2 bytes
End Type
Private mtInts As FOURINTS_T
Function BytesToNumEx(ByteArray() As Byte, StartRec As Long, _
   EndRec As Long, UnSigned As Boolean) As Double
' ###################################################
' Author                : Imran Zaheer
' Contact               : imraanz@mail.com
' Date                  : January 2000
' Function BytesToNumEx : Convertes the specified byte array
'                         into the corresponding Integer or Long
'                         or any signed/unsigned
'                        ;(non-float) data type.
'
' * BYTES : LIKE NUMBERS(Integer/Long etc.) STORED IN A
' * BINARY FILE

' Parameters :
'  (All parameters are reuuired: No Optional)
'     ByteArray() : byte array containg a number in byte format
'  StartRec    : specify the starting array record within the
                 ' array
'     EndRec      : specify the end array record within the array
'     UnSigned    : when False process bytes for both -ve and
'                   +ve values.
'                   when true only process the bytes for +ve
'                   values.
'
' Note: If both "StartRec" and "EndRec" Parameters are zero,
'       then the complete array will be processed.
'
' Example Calls :
'      dim myArray(1 To 4) as byte
'      dim myVar1 as Integer
'      dim myVar2 as Long
'
'      myArray(1) = 255
'      myArray(2) = 127
'      myVar1 = BytesToNumEx(myArray(), 1, 2, False)
'  after execution of above statement myVar1 will be 32767
'
'      myArray(1) = 0
'      myArray(2) = 0
'      myArray(3) = 0
'      myArray(4) = 128
'      myVar2 = BytesToNumEx(myArray(), 1, 4, False)
'  after execution of above statement myVar2 will be -2147483648
'
'
'####################################################
On Error GoTo ErrorHandler
Dim i As Integer
Dim lng256 As Double
Dim lngReturn As Double
    
    lng256 = 1
    lngReturn = 0
    
    If EndRec < 1 Then
        EndRec = UBound(ByteArray)
    End If
    
    If StartRec > EndRec Or StartRec < 0 Then
        MsgBox _
         "Start record can not be greater then End record...!", _
          vbInformation
        BytesToNumEx = -1
        Exit Function
    End If
    
    lngReturn = lngReturn + (ByteArray(StartRec))
    For i = (StartRec + 1) To EndRec
        lng256 = lng256 * 256
        If i < EndRec Then
            lngReturn = lngReturn + (ByteArray(i) * lng256)
        Else
           ' if -ve

            If ByteArray(i) > 127 And UnSigned = False Then
             lngReturn = (lngReturn + ((ByteArray(i) - 256) _
                  * lng256))
            Else
                lngReturn = lngReturn + (ByteArray(i) * lng256)
            End If
        End If
    Next i
    
    BytesToNumEx = lngReturn
ErrorHandler:
End Function

Public Function DEC2BIN(Value As String, Optional X As Integer) As String
Dim iVal#, temp#, ret%, i%, str$
Dim BinVal%()

iVal = Value
Do
    temp = iVal / 2
    ret = InStr(temp, ".")
    If ret > 0 Then
        temp = Left(temp, ret - 1)
    End If
    ret = iVal Mod 2
    ReDim Preserve BinVal(i)
    BinVal(i) = ret
    i = i + 1
    iVal = temp
Loop While temp > 0
For i = UBound(BinVal) To 0 Step -1
    str = str + CStr(BinVal(i))
Next
If X = 3 Then
    Select Case Len(str) Mod 3
        Case 1:
            str = "00" + str
        Case 2:
            str = "0" + str
    End Select
ElseIf X = 4 Then
    Select Case Len(str) Mod 4
        Case 1:
            str = "000" + str
        Case 2:
            str = "00" + str
        Case 3:
            str = "0" + str
    End Select
End If
DEC2BIN = str

End Function
Private Function ByteString(ByVal pNumber As Long) As String


    Do
        If pNumber Mod 2 = 1 Then ByteString = "1" & ByteString Else ByteString = "0" & ByteString
        pNumber = pNumber \ 2
    Loop Until pNumber = 0

    ByteString = String$(4 - (Len(ByteString) Mod 4), "0") & ByteString
End Function

'==================
'Binary To Decimal
' =================
Function Bin2Dec(BinaryString As String) As Variant
   Dim X As Integer
   For X = 0 To Len(BinaryString) - 1
       Bin2Dec = CDec(Bin2Dec) + Val(Mid(BinaryString, _
                 Len(BinaryString) - X, 1)) * 2 ^ X
   Next
End Function
Public Function Channelbits(channel As String) As String
 Dim channelarry() As String
 Dim i As Integer
 Dim tmpchannel As Long
 If InStr(1, channel, ",") <> 0 Then
 channelarry() = Split(channel, ",")
    Channelbits = "0"
    For i = LBound(channelarry) To UBound(channelarry)
     Channelbits = Format(CDToH(2 ^ (CLng(channelarry(i)))), "00000000") Or Channelbits
    Next
  Else
    If channel <> "" Then
        tmpchannel = 2 ^ (CLng(channel))
        Channelbits = Format(CDToH(tmpchannel), "00000000")
    End If
  End If

End Function

Public Function ByteToBinaryString(byteVal As Byte) As String
Dim BitPower As Long
  'Loop through the bits, testing and concatenating the string
  For BitPower = 7 To 0 Step -1
    If byteVal And 2 ^ BitPower Then
      ByteToBinaryString = ByteToBinaryString & "1"
    Else
      ByteToBinaryString = ByteToBinaryString & "0"
    End If
  Next BitPower
End Function

Public Function ByteArrayToString(bytes() As Byte) As String
    Dim iUnicode As Long, i As Long, j As Long
    
    On Error Resume Next
    i = UBound(bytes)
    
    If (i < 1) Then
        'ANSI, just convert to unicode and return
        ByteArrayToString = StrConv(bytes, vbUnicode)
        Exit Function
    End If
    i = i + 1
    
    'Examine the first two bytes
    CopyMemory iUnicode, bytes(0), 2
    
    If iUnicode = bytes(0) Then 'Unicode
        'Account for terminating null
        If (i Mod 2) Then i = i - 1
        'Set up a buffer to recieve the string
        ByteArrayToString = String$(i / 2, 0)
        'Copy to string
        CopyMemory ByVal StrPtr(ByteArrayToString), bytes(0), i
    Else 'ANSI
        ByteArrayToString = StrConv(bytes, vbUnicode)
    End If
                    
End Function

Public Function StringToByteArray(strInput As String, _
                                Optional bReturnAsUnicode As Boolean = True, _
                                Optional bAddNullTerminator As Boolean = False) As Byte()
    
    Dim lRet As Long
    Dim bytBuffer() As Byte
    Dim lLenB As Long
    
    If bReturnAsUnicode Then
        'Number of bytes
        lLenB = LenB(strInput)
        'Resize buffer, do we want terminating null?
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        'Copy characters from string to byte array
        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
    Else
        lLenB = Len(strInput)
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        lRet = WideCharToMultiByte(CP_ACP, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, 0&)
    End If
    
    StringToByteArray = bytBuffer
    
End Function
Public Function LongToByteArray(ByVal lng As Long) As Byte()

Dim ByteArray(0 To 3) As Byte
CopyMemory ByteArray(0), ByVal VarPtr(lng), Len(lng)
LongToByteArray = ByteArray

End Function
Public Function DoubleToByteArray(ByVal dbl As Double) As Byte()
Dim ByteArray(0 To 7) As Byte
CopyMemory ByteArray(0), ByVal VarPtr(dbl), Len(dbl)
DoubleToByteArray = ByteArray

End Function

Public Sub DoubleToLongs(dSource As Double, alDest() As Long)
   ' Convert an 8-byte double to an array of Longs

   ' put double value into UDT
   mtDouble.d = dSource
   ' copy UDT
   LSet mtLongs = mtDouble

   ' populate long array from UDT
   ReDim alDest(0 To 1)
   alDest(0) = mtLongs.al(0)
   alDest(1) = mtLongs.al(1)

End Sub

Public Sub DoubleToInts(dSource As Double, aiDest() As Integer)
   ' Convert an 8-byte double to an array of Integers

   ' put double value into UDT
   mtDouble.d = dSource
   ' copy UDT
   LSet mtInts = mtDouble

   ' populate integer array from UDT
   ReDim aiDest(0 To 3)
   aiDest(0) = mtInts.ai(0)
   aiDest(1) = mtInts.ai(1)
   aiDest(2) = mtInts.ai(2)
   aiDest(3) = mtInts.ai(3)

End Sub

Public Function ByteArrayToDouble(ByRef bytes() As Byte) As Double
Dim DoubleNum As Double
  'This pretends the next eight bytes of the packet is a Double
  CopyMemory DoubleNum, bytes(0), Len(DoubleNum)
  ByteArrayToDouble = DoubleNum
End Function

Public Function ValToDms(ByVal dblVal As Double, _
                       blnLatitude As Boolean, Optional wiscan As Boolean = False) As String
Dim strDMS As String
Dim intDeg As Integer
Dim intMin As Integer
Dim intSec As Integer
Dim dblmin As Double
'On Error Resume Next

'Create a temporary variable strDms to hold the result of the
'conversion. Build the converted value in steps.
'
'One of the first things to do is get rid of an eventual negative sign,
'and compute the compass point at the same time. So we get a big
If dblVal < 0 Then
   strDMS = IIf(blnLatitude, "S", "W")
Else
   strDMS = IIf(blnLatitude, "N", "E")
End If
'Continue with a positive number
If dblVal < 0 Then dblVal = -dblVal
'Get the degrees
intDeg = Int(dblVal)
If wiscan Then
    strDMS = strDMS & " " & Format(CStr(dblVal), "00.0000000")
Else
    strDMS = strDMS & CStr(intDeg) & Chr(176)
    dblmin = (dblVal - intDeg) * 60
    strDMS = strDMS & Format(CStr(dblmin), "0#.000") & Chr(39)
End If

'intMin = Int(dblmin)
'strDMS = strDMS & CStr(intMin) & "."

'dblmin = dblmin - intMin '; // remove the whole part
'dblmin = mdblminin * 1000# '; // get some places past the decimal.
   


'The last statements in your function should of course be

   ValToDms = strDMS
End Function


Public Sub DoubletoBinaryString(dblvar As Double)
Dim DblByteArray(0 To 7) As Byte
Dim ByteCount As Integer, BinaryString As String
  'Place the Double in a Byte array to prevent overflows in bits > 30
  CopyMemory DblByteArray(0), dblvar, 8
  'Loop through all 8 bytes concatenating our string of bits
  For ByteCount = 7 To 0 Step -1
    BinaryString = BinaryString & " " & _
    ByteToBinaryString(DblByteArray(ByteCount))
  Next ByteCount
  'Display results
End Sub




