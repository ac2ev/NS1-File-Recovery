Attribute VB_Name = "modNs1File"
Option Explicit
'********************************************************************************
' Name: modNs1File
' Created by: 
'
' Ns1 Binary File Format Copyright © Marius Milner, 2003-2004 
'
' Requirements: Bytes2Num.bas
'
' Description: These routines will read and write a Native ns1 file
'
' Usage:       OpenBinFile
'              ret = Read_ns1Header(Offset)
'             'Error checking for proper header
'              Step = 1
'              For ApCount = 1 To NS1.ApCount
'                apDone = False
'                Do While apDone = False
'                    ret = Read_ApInfo(ApCount - BadRecords.Count, Offset, Step, apDone)
'                    If ret <> 0 And ret <> 9999 Then Exit For
'                Loop
'              Next ApCount
'              Closefile
'
'Notes: Some variable types include both the HR (Human Readable) and original byte structure.
'       This allows us to write the structure back to the file unmodified
'********************************************************************************
Global NoSSID As String
Global NoPrint As Long
Public mFileSize As Long

Public Type ByteHRD
    bytes() As Byte
    dbl As Double
End Type
Public Type ByteHRS
    bytes() As Byte
    str As String
End Type
Public Type ByteHRFT
    bytes() As Byte
    File_Time As FILETIME
    Time As Date
End Type

Private Type GPSDATA
    Latitude           As ByteHRD
    Longitude          As ByteHRD
    Altitude           As ByteHRD
    NumSats            As Long
    Speed              As ByteHRD
    Track              As ByteHRD
    MagVariation       As ByteHRD
    Hdop               As ByteHRD
End Type

Public Type APData
    Time                As ByteHRFT
    Signal              As Long
    Noise               As Long
    Location_Source     As Long
    GPSDATA             As GPSDATA
End Type

Public Type apinfo
    SSIDLength          As Long
    SSID                As String
    BSSID               As String
    MaxSignal           As Long
    MinNoise            As Long
    MaxSNR              As Long
    flags               As String
    BeaconInterval      As Long
    firstseen           As ByteHRFT
    lastseen            As ByteHRFT
    BestLat             As ByteHRD
    BestLong            As ByteHRD
    DataCount           As Long
    APData()            As APData
    NameLength          As Long
    Name                As String
    Channels            As ByteHRS
    LastChannel         As Long
    IPAddress           As String
    MinSignal           As Long
    MaxNoise            As Long
    DataRate            As Long
    IPSubnet            As String
    IPMask              As String
    ApFlags             As Long
    IELength            As Long
    InformationElements As Long
    Offsets             As ItemOffset
End Type

Public Type ns1
    dwSignature As String * 4
    dwFileVer As Long
    apcount As Long
    apinfo() As apinfo
End Type
Public Type BadRecord
    Items As ns1
    indexes() As Long
End Type
Public gphIndex As Long
Public BadRecords As BadRecord
Public fNum As Integer
Public ns1 As ns1
Public MergedNs1 As ns1
Dim prevFoundPos As Long
Dim apDataOffset As Long
Dim SkipPattern As Boolean
Dim prevIndex As Long


Public Function IsArrayInit(arTest() As apinfo) As Boolean
    'Check if array is initialized.
    
    On Error GoTo ErrHandler
    Dim intMax As Integer
    
    intMax = UBound(arTest)
    
    IsArrayInit = True
    
exitHandler:
    Exit Function
    
ErrHandler:
    IsArrayInit = False
    Resume exitHandler
End Function

Public Function Read_ns1Header(ByRef Offset As Long, Optional wififofum As Boolean = False) As Long
    Dim bytes() As Byte
    Offset = 1
    If wififofum Then
        ReDim bytes(0)
        Offset = GetBytes(Offset, bytes())
    End If
    ReDim bytes(3)
    Offset = GetBytes(Offset, bytes())
    ns1.dwSignature = StrConv(bytes(), vbUnicode)
    
    If ns1.dwSignature = "NetS" Then
        
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.dwFileVer = BytesToNumEx(bytes, 0, 0, True)
        
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apcount = BytesToNumEx(bytes, 0, 0, True)

     Else
        MsgBox "Not a valid ns1 file", vbCritical + vbOKOnly, "Invalid File Type"
        Read_ns1Header = 99
     End If
End Function

 '******************************************************************************************
 ' Description:
 '               This function must be called recursively from an outside source
 '               It will populate the Ns1 Structure
 '******************************************************************************************

Public Function Read_ApInfo(ByRef Index As Long, ByRef Offset As Long, ByRef step As Long, ByRef apdone As Boolean, ByRef Corrupt As Boolean, Optional wififofum As Boolean = False) As Long
 Dim File_Time As FILETIME
 Dim ret As VbMsgBoxResult
 Dim Temp_Date As Date
 Static bytes() As Byte
 Dim lData As Long
 Dim lSource As Long
 Dim lngTemp As Long
 Dim dblTemp As Double
 Dim i As Long
 Dim x As Long
 Dim tmpStr As String
 Dim NullCnt As Long
 
 On Error Resume Next
 'Version 6 of Netstumbler file has channels in a different position and different format then later versions
 'So, here's what happens. We do steps 1-6, then goto 28 to do the channels then next time around we pick
 'up at step 7. Since version 6 only goes to step 16 (Name) we don't have to worry about hitting step
 '17 because the file version will override it.
  
 On Error GoTo ErrHandler
    Select Case ns1.dwFileVer
        Case 6
            If step = 17 Then step = 99  'Version 6 of the ns1 file format ends at step 17
        Case 8
            If step = 20 Then step = 99  'Version 8 of the ns1 file format ends at step 19
        Case 11
            If step = 25 Then step = 99  'Version 11 of the ns1 file format ends at step 24
     End Select
 'If Index = 304 Then MsgBox "here"
PatternMatch:
 Select Case step
     
     Case 1 'SSIDLength
       ns1.apinfo(Index).Offsets.BeginOffset = Offset
       ReDim bytes(0)
       Offset = GetBytes(Offset, bytes)
       ns1.apinfo(Index).SSIDLength = BytesToNumEx(bytes, 0, 0, True)
       step = step + 1
       
    Case 2 'SSID
      If ns1.apinfo(Index).SSIDLength > 0 Then
      If wififofum Then 'bypass wififofum 0.3.3 error of extra ssid length byte
           ReDim bytes(0)
           Offset = GetBytes(Offset, bytes)
      End If
       
       'For Ministumbler .ns1 files we need to see if SSID is null seperated
       ReDim bytes(ns1.apinfo(Index).SSIDLength - 1)
       Offset = GetBytes(Offset, bytes)
              
       ns1.apinfo(Index).SSID = StrConv(bytes(), vbUnicode)
       Else
        ns1.apinfo(Index).SSIDLength = Len(NoSSID)
        ns1.apinfo(Index).SSID = NoSSID
       End If
       If NoPrint Then
        For x = 1 To ns1.apinfo(Index).SSIDLength
            If x > Len(ns1.apinfo(Index).SSID) Then Exit For
            If Asc(Mid$(ns1.apinfo(Index).SSID, x, 1)) < 32 Or Asc(Mid$(ns1.apinfo(Index).SSID, x, 1)) > 126 Then
                ns1.apinfo(Index).SSID = Left$(ns1.apinfo(Index).SSID, x - 1) & Right$(ns1.apinfo(Index).SSID, Len(ns1.apinfo(Index).SSID) - (x - 1) - 1)
                x = x - 1
            End If
        Next
       End If
       step = step + 1
    
    Case 3 'BSSID
       ReDim bytes(5)
       Offset = GetBytes(Offset, bytes)
       ns1.apinfo(Index).BSSID = CDToH(Val(bytes(0))) & ":" & CDToH(Val(bytes(1))) & ":" & CDToH(Val(bytes(2))) & _
                                 ":" & CDToH(Val(bytes(3))) & ":" & CDToH(Val(bytes(4))) & ":" & CDToH(Val(bytes(5)))
       step = step + 1
       
       If ns1.apinfo(Index).BSSID = "00:00:00:00:00:00" And ns1.apinfo(Index).SSID = "<no ssid>" Then
            ' Bad Record

            ret = MsgBox("Null record found at index " & Index & vbCrLf & _
                        "SSID is: " & ns1.apinfo(Index).SSID & vbCrLf & _
                        "BSSID is: " & ns1.apinfo(Index).BSSID & vbCrLf & _
                         "Do you want to continue? ", vbYesNo, "Bad Record")
            If ret = vbNo Then
                BadRecords.Items.apcount = BadRecords.Items.apcount + 1
                ReDim Preserve BadRecords.Items.apinfo(BadRecords.Items.apcount)
                ReDim Preserve BadRecords.indexes(BadRecords.Items.apcount)
                BadRecords.Items.dwFileVer = ns1.dwFileVer
                BadRecords.Items.dwSignature = ns1.dwSignature
                BadRecords.Items.apinfo(BadRecords.Items.apcount) = ns1.apinfo(Index)
                BadRecords.indexes(BadRecords.Items.apcount) = Index
                apdone = True
                step = 1
                Read_ApInfo = 9999
            End If
       End If

    Case 4 'MaxSignal
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).MaxSignal = BytesToNumEx(bytes, 0, 0, False)
        step = step + 1
    
    Case 5 'MinNoise
       ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).MinNoise = BytesToNumEx(bytes, 0, 0, False)
        step = step + 1
    
    Case 6 'MaxSNR
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).MaxSNR = BytesToNumEx(bytes, 0, 0, False)
        
        If ns1.dwFileVer = 6 Then
            step = 28 'Channels is in a different position in version 6
        Else
            step = step + 1
        End If
    Case 7 'Flags
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).flags = BytesToNumEx(bytes, 0, 0, True)
        step = step + 1
 
    Case 8 'BeaconInterval
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).BeaconInterval = BytesToNumEx(bytes, 0, 0, True) * 1.024 'Kus
        step = step + 1

    Case 9 'FirstSeen
        ReDim bytes(7)
        Offset = GetBytes(Offset, bytes())
        File_Time.dwHighDateTime = BytesToNumEx(bytes, 4, 7, False)
        File_Time.dwLowDateTime = BytesToNumEx(bytes, 0, 3, False)
        'Currently some issues converting between file time and time so I store 3 formats now
        ns1.apinfo(Index).firstseen.File_Time = File_Time
        ns1.apinfo(Index).firstseen.bytes = bytes
        ns1.apinfo(Index).firstseen.Time = FileTimeToDate(UtcToLocalFileTime(File_Time))
        'Debug.Print ns1.APINFO(Index).firstseen.Time, FileTimeToDate(ns1.APINFO(Index).firstseen.File_Time)
        step = step + 1

    Case 10 'LastSeen
        ReDim bytes(7)
        Offset = GetBytes(Offset, bytes())
        File_Time.dwHighDateTime = BytesToNumEx(bytes, 4, 7, False)
        File_Time.dwLowDateTime = BytesToNumEx(bytes, 0, 3, False)
        'Currently some issues converting between file time and time so I store 3 formats now
        ns1.apinfo(Index).lastseen.File_Time = File_Time
        ns1.apinfo(Index).lastseen.bytes = bytes
        ns1.apinfo(Index).lastseen.Time = FileTimeToDate(UtcToLocalFileTime(File_Time))
        
        step = step + 1
    
    Case 11 'BestLat
       ReDim bytes(7)
       Offset = GetBytes(Offset, bytes)
       ns1.apinfo(Index).BestLat.dbl = ByteArrayToDouble(bytes)
       ns1.apinfo(Index).BestLat.bytes = bytes
       step = step + 1
    
    Case 12 'BestLong
       ReDim bytes(7)
       Offset = GetBytes(Offset, bytes)
       ns1.apinfo(Index).BestLong.dbl = ByteArrayToDouble(bytes)
       ns1.apinfo(Index).BestLong.bytes = bytes
       step = step + 1

    Case 13 'DataCount
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).DataCount = BytesToNumEx(bytes, 0, 0, True)
        step = step + 1
    
    Case 14 'ApData
         If ns1.apinfo(Index).DataCount <> 0 Then
         apDataOffset = Offset
         ReDim ns1.apinfo(Index).APData(ns1.apinfo(Index).DataCount)
         For lData = 0 To ns1.apinfo(Index).DataCount - 1
             If IsArrayInit(ns1.apinfo) Then
             With ns1.apinfo(Index).APData(lData)
                 Erase bytes()
                 ReDim bytes(7)
                 Offset = GetBytes(Offset, bytes())
                 .Time.bytes = bytes
                 
                 File_Time.dwHighDateTime = BytesToNumEx(bytes, 4, 7, False)
                 File_Time.dwLowDateTime = BytesToNumEx(bytes, 0, 3, False)
                 .Time.File_Time = File_Time
                 .Time.Time = FileTimeToDate(UtcToLocalFileTime(File_Time))
                 
                 Erase bytes()
                 ReDim bytes(3)
                 Offset = GetBytes(Offset, bytes())
                 .Signal = BytesToNumEx(bytes, 0, 0, False)
                 
                 Erase bytes()
                 ReDim bytes(3)
                 Offset = GetBytes(Offset, bytes())
                 .Noise = BytesToNumEx(bytes, 0, 0, False)
                 
                 Erase bytes()
                 ReDim bytes(3)
                 Offset = GetBytes(Offset, bytes())
                 .Location_Source = BytesToNumEx(bytes, 0, 0, False)
                 If .Location_Source > 31 Then .Location_Source = 1
                 If .Location_Source <> 0 Then
                       With .GPSDATA
                             Erase bytes()
                             ReDim bytes(7)
                             Offset = GetBytes(Offset, bytes())
                             .Latitude.bytes = bytes
                             .Latitude.dbl = ByteArrayToDouble(bytes)
                             
                             Erase bytes()
                             ReDim bytes(7)
                             Offset = GetBytes(Offset, bytes())
                             .Longitude.bytes = bytes
                             .Longitude.dbl = ByteArrayToDouble(bytes)
                      
                             
                             Erase bytes()
                             ReDim bytes(7)
                             Offset = GetBytes(Offset, bytes())
                             .Altitude.dbl = ByteArrayToDouble(bytes)
                             .Altitude.bytes = bytes
                             
                             Erase bytes()
                             ReDim bytes(3)
                             Offset = GetBytes(Offset, bytes())
                             .NumSats = BytesToNumEx(bytes, 0, 0, False)
                             
                             Erase bytes()
                             ReDim bytes(7)
                             Offset = GetBytes(Offset, bytes())
                             .Speed.dbl = ByteArrayToDouble(bytes)
                             .Speed.bytes = bytes
                             
                             Erase bytes()
                             ReDim bytes(7)
                             Offset = GetBytes(Offset, bytes())
                             .Track.dbl = ByteArrayToDouble(bytes)
                             .Track.bytes = bytes
                             
                             Erase bytes()
                             ReDim bytes(7)
                             Offset = GetBytes(Offset, bytes())
                             .MagVariation.dbl = ByteArrayToDouble(bytes)
                             .MagVariation.bytes = bytes
                             
                             Erase bytes()
                             ReDim bytes(7)
                             Offset = GetBytes(Offset, bytes())
                             .Hdop.dbl = ByteArrayToDouble(bytes)
                             .Hdop.bytes = bytes
                         End With
                 End If
             End With
             End If
         Next lData
        End If
        step = step + 1
    
    Case 15 'NameLength
        ReDim bytes(0)
        Offset = GetBytes(Offset, bytes)
        ns1.apinfo(Index).NameLength = BytesToNumEx(bytes, 0, 0, True)
        step = step + 1
        
    Case 16 'Name
       If ns1.apinfo(Index).NameLength <> 0 Then
            ReDim bytes(ns1.apinfo(Index).NameLength - 1)
            Offset = GetBytes(Offset, bytes)
            ns1.apinfo(Index).Name = StrConv(bytes(), vbUnicode)
        End If
        step = step + 1
    
    Case 17 'Channels version 8+
           ReDim bytes(7)
           Offset = GetBytes(Offset, bytes())
    
           ns1.apinfo(Index).Channels.bytes = bytes
           'Split into bits and parse out positons = 1 as channels
                   'bytes() = DoubleToByteArray(BytesToNumEx(bytes, 0, 0, True))
                   
           tmpStr = StrReverse(DEC2BIN(BytesToNumEx(bytes, 0, 0, True)))
           For i = 1 To Len(tmpStr)
            If Mid(tmpStr, i, 1) = 1 Then
                ns1.apinfo(Index).Channels.str = ns1.apinfo(Index).Channels.str & "," & i - 1
            End If
            If Left(ns1.apinfo(Index).Channels.str, 1) = "," Then ns1.apinfo(Index).Channels.str = Right(ns1.apinfo(Index).Channels.str, Len(ns1.apinfo(Index).Channels.str) - 1)
            
           Next i
           'Debug.Print BytesToNumEx(bytes, 0, 0, True), DEC2BIN(BytesToNumEx(bytes, 0, 0, True)), ns1.APINFO(Index).Channels.str
          
          step = step + 1
    
    Case 18 'LastChannel
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).LastChannel = BytesToNumEx(bytes, 0, 0, False)
        step = step + 1
    
    Case 19 'IPAddress
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).IPAddress = Val(bytes(0)) & "." & Val(bytes(1)) & "." & Val(bytes(2)) & _
                                 "." & Val(bytes(3))
        step = step + 1
    
    Case 20 'MinSignal
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).MinSignal = BytesToNumEx(bytes, 0, 0, False)
        step = step + 1
        
    Case 21 'MaxNoise
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).MaxNoise = BytesToNumEx(bytes, 0, 0, False)
        step = step + 1
    
    Case 22 'DataRate
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).DataRate = BytesToNumEx(bytes, 0, 0, True)
        step = step + 1
        
    Case 23 'IPSubnet
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).IPSubnet = Val(bytes(0)) & "." & Val(bytes(1)) & "." & Val(bytes(2)) & _
                                 "." & Val(bytes(3))
        step = step + 1
    
    Case 24 'IPMask
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).IPMask = Val(bytes(0)) & "." & Val(bytes(1)) & "." & Val(bytes(2)) & _
                                 "." & Val(bytes(3))
        step = step + 1
    
    Case 25 'ApFlags
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).ApFlags = BytesToNumEx(bytes, 0, 0, True)
        step = step + 1
        
    Case 26 'IELength
        ReDim bytes(3)
        Offset = GetBytes(Offset, bytes())
        ns1.apinfo(Index).IELength = BytesToNumEx(bytes, 0, 0, True)
        step = step + 1
        
    Case 27 'InformationElements
        If ns1.apinfo(Index).IELength <> 0 Then
            ReDim bytes(ns1.apinfo(Index).IELength - 1)
            Offset = GetBytes(Offset, bytes())
            ns1.apinfo(Index).InformationElements = BytesToNumEx(bytes, 0, 0, True)
        End If
        step = 99
    
    Case 28 'Channels Version 6
           ReDim bytes(3)
           Offset = GetBytes(Offset, bytes())
    
           ns1.apinfo(Index).Channels.bytes = bytes
           'Split into bits and parse out positons = 1 as channels
           tmpStr = StrReverse(DEC2BIN(BytesToNumEx(bytes, 0, 0, True)))
           
           For i = 1 To Len(tmpStr)
            If Mid(tmpStr, i, 1) = 1 Then
                ns1.apinfo(Index).Channels.str = ns1.apinfo(Index).Channels.str & "," & i - 1
            End If
            If Left(ns1.apinfo(Index).Channels.str, 1) = "," Then ns1.apinfo(Index).Channels.str = Right(ns1.apinfo(Index).Channels.str, Len(ns1.apinfo(Index).Channels.str) - 1)
            
           Next i
          step = 7
    Case Else
        ns1.apinfo(Index).Offsets.EndOffset = Offset
        If Corrupt = True Then
                BadRecords.Items.apcount = BadRecords.Items.apcount + 1
                ReDim Preserve BadRecords.Items.apinfo(BadRecords.Items.apcount)
                ReDim Preserve BadRecords.indexes(BadRecords.Items.apcount)
                BadRecords.Items.dwFileVer = ns1.dwFileVer
                BadRecords.Items.dwSignature = ns1.dwSignature
                BadRecords.Items.apinfo(BadRecords.Items.apcount) = ns1.apinfo(Index)
                BadRecords.indexes(BadRecords.Items.apcount) = Index
        Debug.Print "foo"
        End If
        apdone = True
        step = 1
End Select
Exit Function

ErrHandler:

Dim msg As String

'    Dim free As Long
'    free = FreeFile
'
'    'MsgBox "Error Number: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
'           "Please email author ns1 file and " & App.Path & "\Errorlog.txt", vbCritical, "Error"
'    Open App.Path & "\Errorlog.txt" For Output As #free
'    Print #free, "----------------------------------------"
'    Print #free, "File: " & fname
'    Print #free, "Signature: " & ns1.dwSignature
'    Print #free, "Version: " & ns1.dwFileVer
'    Print #free, "Ap Count: " & ns1.apcount
'    Print #free, Err.Description & vbCrLf & "At Step " & step
'    Print #free, "Index: " & Index
'    Print #free, "Offset: " & Offset
'    Print #free, "SSID: " & ns1.apinfo(Index).SSID
'    Print #free, "+++++++++++++++++++++++++++++++++++++++++"
'    Close #free
If Not SkipPattern Then
 msg = "Data corruption found" & vbCrLf & _
       "Record: " & Index & vbCrLf & _
       "Byte: " & Offset & vbCrLf & _
       "Pattern matching recovery will be attempted to recover data"
 ret = MsgBox(msg, vbOKOnly, "Attempt Pattern Matching")
Dim curoffset As Long
'Attempt data recovery
curoffset = Offset
'Pattern = 6E 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00
'Look for the 6E
If HexSearch("6E00000000000000000000000000000000000000", curoffset) Then
    If curoffset + Offset < mFileSize Then
        Offset = curoffset + Offset
    End If
        SeekBytes Offset
        Corrupt = True
        step = 99
Else
    SkipPattern = True
    Offset = apDataOffset
    step = 14
End If


 
 
 Exit Function
End If
    If gDebugMode Then
        Debug.Print Err.Description
        Resume Next
    End If
    Read_ApInfo = Err.Number
    Resume Next
End Function

 '******************************************************************************************
 ' Description:
 '               Writes Ns1 File Header elements
 ' Assumptions: Valid file handle open
 '              Valid Ns1 data structure
 '******************************************************************************************

Public Function Write_Ns1Header(Optional Newcount As Long = 0)
    PutBytes StringToByteArray(ns1.dwSignature, False, False)  'Write the signature
    PutBytes LongToByteArray(ns1.dwFileVer)    'Write the version
    If Newcount = 0 Then
        PutBytes LongToByteArray(UBound(ns1.apinfo) - BadRecords.Items.apcount)    'Write the record count
    Else
        PutBytes LongToByteArray(Newcount)
    End If

End Function

Private Function HexSearch(Pattern As String, ByRef Offset As Long) As Boolean
    On Error Resume Next
    Dim HexCnt As Integer
    Dim i, j
    Dim mMatch As Boolean
    Dim foundStartPos As Long
    Dim arrByte() As Byte
    Dim curoffset As Long
    
    ReDim arrByte(mFileSize - Offset) 'Dimension to what's left
    GetBytes Offset, arrByte()
    
    
    HexCnt = Len(Pattern) / 2
    ReDim arrHexByte(1 To HexCnt)
    For i = 1 To HexCnt
         arrHexByte(i) = CByte("&h" & (Mid(Pattern, (i * 2 - 1), 2)))
    Next i
    foundStartPos = prevFoundPos + 1
    For i = foundStartPos To (UBound(arrByte) - (HexCnt - 1))
         If arrByte(i) = arrHexByte(1) Then
              mMatch = True
                ' Compare rest bytes
              For j = 1 To (HexCnt - 1)
                   If arrByte(i + j) <> arrHexByte(1 + j) Then
                       mMatch = False
                       Exit For
                   End If
              Next j
              If mMatch = True Then
                   Dim k
                   foundStartPos = i
'                   prevFoundPos = i
'                   k = (foundStartPos + 1) / CLng(mPageSize)
'                   k = NoFraction(k)
'                   pageStart = k * mPageSize + 1
'                   pageEnd = pageStart + mPageSize - 1
'                   If pageEnd > mFileSize Then pageEnd = mFileSize
'                   k = foundStartPos + (HexCnt - 1)
'                   If k > pageEnd Then k = pageEnd
'                   updEditByte
'                   ShowPage True, foundStartPos, k, &HFFFF00, vbRed
'                   Screen.MousePointer = vbDefault
                   
                   Offset = foundStartPos + HexCnt
                   HexSearch = mMatch
                   
                   Exit Function
              End If
         End If
    Next i
    prevFoundPos = 0
    MsgBox Pattern & vbCrLf & vbCrLf & "Searched to end."
End Function

 
Public Function Write_ApInfo(ByRef ns1 As ns1, ByRef Index As Long, ByRef Offset As Long, ByRef step As Long, ByRef apdone As Boolean, Optional merge As Boolean = False) As Long
 
 Dim File_Time As FILETIME
 Dim Temp_Date As Date
 Dim bytes() As Byte
 Dim aryLng() As Long
 Dim lData As Long
 Dim lgps As Long
 Dim lngTemp As Long
 Dim dblTemp As Double
 Dim arry() As String
 Dim stemp As String
 
    
 On Error GoTo ErrHandler
 Erase bytes()
 Dim i As Long
 
    Select Case ns1.dwFileVer
        Case 6
            If step = 17 Then step = 99  'Version 6 of the ns1 file format ends at step 17
        Case 8
            If step = 20 Then step = 99  'Version 8 of the ns1 file format ends at step 19
        Case 11
            If step = 25 Then step = 99  'Version 11 of the ns1 file format ends at step 24
     End Select
 
 Select Case step
    Case 1 'SSIDLength
       ReDim bytes(0)
       bytes(0) = ns1.apinfo(Index).SSIDLength
       PutBytes bytes
       step = step + 1
       
    Case 2 'SSID
       If ns1.apinfo(Index).SSIDLength > 0 Then
        PutBytes StringToByteArray(ns1.apinfo(Index).SSID, False, False)
       End If
       step = step + 1
       
    Case 3 'BSSID
       If Len(ns1.apinfo(Index).BSSID) = 12 Then
        stemp = ""
        For i = 1 To Len(ns1.apinfo(Index).BSSID) Step 2
           ' Debug.Print Mid(ns1.apinfo(Index).BSSID, i, 2), sTemp
            If i > 1 Then
            stemp = stemp & ":" & Mid(ns1.apinfo(Index).BSSID, i, 2)
            Else
            stemp = Mid(ns1.apinfo(Index).BSSID, i, 2)
            End If
            'Debug.Print sTemp
        Next
       ns1.apinfo(Index).BSSID = stemp
       End If
       
       arry() = Split(ns1.apinfo(Index).BSSID, ":")
       ReDim Preserve arry(5)
       ReDim bytes(5)
       
       For i = 0 To UBound(arry)
            bytes(i) = CByte(CHToD(arry(i)))
       Next i
       
       PutBytes bytes
       step = step + 1
    
    Case 4 'MaxSignal
       PutBytes LongToByteArray(ns1.apinfo(Index).MaxSignal)
       step = step + 1
    
    Case 5 'MinNoise
       
       PutBytes LongToByteArray(ns1.apinfo(Index).MinNoise)
       step = step + 1
    
    Case 6 'MaxSNR
        
        PutBytes LongToByteArray(ns1.apinfo(Index).MaxSNR)
        If ns1.dwFileVer = 6 Then
            step = 28
        Else
            step = step + 1
        End If
    
    Case 7 'Flags
        ReDim bytes(3)
        If ns1.apinfo(Index).flags = "" Then ns1.apinfo(Index).flags = "0"
        PutBytes LongToByteArray(CLng(ns1.apinfo(Index).flags))
        step = step + 1
    
    Case 8 'BeaconInterval
        
        PutBytes LongToByteArray(ns1.apinfo(Index).BeaconInterval)
        step = step + 1
    
    Case 9 'FirstSeen
        'File_Time = ns1.APINFO(Index).firstseen.File_Time
        File_Time = UtcFromLocalFileTime(FileTimeFromDate(ns1.apinfo(Index).firstseen.Time))
        PutBytes LongToByteArray(File_Time.dwLowDateTime)
        PutBytes LongToByteArray(File_Time.dwHighDateTime)

        step = step + 1
        
    Case 10 'LastSeen
        
        'File_Time = ns1.APINFO(Index).lastseen.File_Time
        File_Time = UtcFromLocalFileTime(FileTimeFromDate(ns1.apinfo(Index).lastseen.Time))
        PutBytes LongToByteArray(File_Time.dwLowDateTime)
        PutBytes LongToByteArray(File_Time.dwHighDateTime)
        step = step + 1
    
    Case 11 'BestLat
        
        PutBytes ns1.apinfo(Index).BestLat.bytes
        step = step + 1
    
    Case 12 'BestLong
        
        PutBytes ns1.apinfo(Index).BestLong.bytes
        step = step + 1
    
    Case 13 'DataCount
        PutBytes LongToByteArray(ns1.apinfo(Index).DataCount)
        step = step + 1
        
    Case 14 'ApData
       If ns1.apinfo(Index).DataCount <> 0 Then
         For lData = 0 To ns1.apinfo(Index).DataCount - 1
            With ns1.apinfo(Index).APData(lData)
            
            PutBytes .Time.bytes
             
            PutBytes LongToByteArray(.Signal)

            PutBytes LongToByteArray(.Noise)

            PutBytes LongToByteArray(.Location_Source)

            If .Location_Source <> 0 Then
                    With .GPSDATA

                         PutBytes .Latitude.bytes
                         
                         PutBytes .Longitude.bytes
                         PutBytes .Altitude.bytes
                         
                         PutBytes LongToByteArray(.NumSats)
                         PutBytes .Speed.bytes
                         
                         PutBytes .Track.bytes
                         
                         PutBytes .MagVariation.bytes
                         
                         PutBytes .Hdop.bytes
                         
                    End With
            End If
            End With
         Next lData
       End If
     step = step + 1
    
    Case 15 'NameLength
    
        ReDim bytes(0)
        bytes(0) = ns1.apinfo(Index).NameLength
        PutBytes bytes
        step = step + 1
    
    Case 16 'Name
        If ns1.apinfo(Index).NameLength > 0 Then PutBytes StringToByteArray(ns1.apinfo(Index).Name, False, False)
        step = step + 1
        
    Case 17 'Channels Version 8+
        
        PutBytes ns1.apinfo(Index).Channels.bytes
        step = step + 1
            
    Case 18 'LastChannel
        bytes = LongToByteArray(CLng(ns1.apinfo(Index).LastChannel))
        PutBytes LongToByteArray(ns1.apinfo(Index).LastChannel)
        step = step + 1

    Case 19 'IPAddress
       If ns1.apinfo(Index).IPAddress = "" Then ns1.apinfo(Index).IPAddress = "0.0.0.0"
       arry() = Split(ns1.apinfo(Index).IPAddress, ".")
       ReDim bytes(UBound(arry))
       
       For i = 0 To UBound(arry)
            bytes(i) = CByte(arry(i))
       Next i
       PutBytes bytes
       step = step + 1
    
    Case 20 'MinSignal
        PutBytes LongToByteArray(ns1.apinfo(Index).MinSignal)
        step = step + 1
    
    Case 21 'MaxNoise
        PutBytes LongToByteArray(ns1.apinfo(Index).MaxNoise)
        step = step + 1

    Case 22 'DataRate
        PutBytes LongToByteArray(ns1.apinfo(Index).DataRate)
        step = step + 1
        
    Case 23 'IPSubnet
       If ns1.apinfo(Index).IPSubnet = "" Then ns1.apinfo(Index).IPSubnet = "0.0.0.0"
       arry() = Split(ns1.apinfo(Index).IPSubnet, ".")
       ReDim bytes(UBound(arry))
       
       For i = 0 To UBound(arry)
            bytes(i) = CByte(arry(i))
       Next i
       PutBytes bytes
       step = step + 1

    Case 24 'IPMask
       If ns1.apinfo(Index).IPMask = "" Then ns1.apinfo(Index).IPMask = "0.0.0.0"
       arry() = Split(ns1.apinfo(Index).IPMask, ".")
       ReDim bytes(UBound(arry))
       
       For i = 0 To UBound(arry)
            bytes(i) = CByte(arry(i))
       Next i
       PutBytes bytes
       step = step + 1

    Case 25 'ApFlags
        PutBytes LongToByteArray(ns1.apinfo(Index).ApFlags)
        step = step + 1

    Case 26 'IELength
        PutBytes LongToByteArray(ns1.apinfo(Index).IELength)
        step = step + 1

    Case 27 'InformationElements
        If ns1.apinfo(Index).IELength <> 0 Then
            ReDim bytes(ns1.apinfo(Index).IELength - 1)
            Offset = GetBytes(Offset, bytes())
            ns1.apinfo(Index).InformationElements = BytesToNumEx(bytes, 0, 0, True)
        End If
        step = 99
    
    Case 28 'Channels Version 6
       
        PutBytes ns1.apinfo(Index).Channels.bytes
        step = 7
    
    Case Else
        ns1.apinfo(Index).Offsets.EndOffset = Offset
        apdone = True
        step = 1
End Select
Exit Function
ErrHandler:
    Dim free
    MsgBox "Error Number: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Please email author ns1 file and " & App.Path & "\Errorlog.txt", vbCritical, "Error" '
    Resume
    On Error Resume Next
    free = FreeFile
    Open App.Path & "\Errorlog.txt" For Output As #free
    Print #free, "----------------------------------------"
    Print #free, "File: " & fname
    Print #free, "Signature: " & ns1.dwSignature
    Print #free, "Version: " & ns1.dwFileVer
    Print #free, "Ap Count: " & ns1.apcount
    Print #free, Err.Description & vbCrLf & "At Step " & step
    Print #free, "Index: " & Index
    Print #free, "Offset: " & Offset
    Print #free, "SSID: " & ns1.apinfo(Index).SSID
    Print #free, "+++++++++++++++++++++++++++++++++++++++++"
    Close #free
    If gDebugMode Then
        Debug.Print Err.Description
        Resume
    End If
    Write_ApInfo = Err.Number
End Function

Public Function OpenBinFile(Optional file As String = "")
    
    fNum = FreeFile
    
    
    If file = "" Then
        file = fname
    End If
    Open file For Binary As #fNum
    mFileSize = LOF(fNum)
End Function
Public Function OpentxtFile(Optional file As String = "")
    
    fNum = FreeFile
    If file = "" Then
        file = fname
    End If
    Open file For Output As #fNum
End Function
Public Function CloseFile()
    Close #fNum
End Function

Public Function GetBytes(StartPOS As Long, ByRef bytes() As Byte) As Long
    Get #fNum, StartPOS, bytes
    GetBytes = StartPOS + UBound(bytes) + 1
End Function

Public Function SeekBytes(position As Long) As Long
    Seek #fNum, position
End Function
Public Function PutBytes(ByRef bytes() As Byte) As Long
    Put #fNum, , bytes
End Function

Public Function PutText(ByRef Text As String) As Long
     Print #fNum, Text
End Function
Public Sub SaveFile(ByRef ns1 As ns1, Optional merged As Boolean = False, Optional SingleItem As Long = 0, Optional count As Long = 0, Optional append As String)
    
    'Save all your hard work!
    Dim FNamebak As String
    Dim Counter As Long
    Dim rfname As String
    Dim apcount As Long
    Dim apdone As Boolean
    Dim Offset As Long
    Dim step As Long
    Dim ret As Long
    Dim blankns1 As ns1
    'Set filenames
    
    
    
    If merged Then
       ns1 = MergedNs1
       MergedNs1 = blankns1
       rfname = Left(fname, InStrRev(fname, "\")) & Format(Now, "YYYYMMDDHHMMSS") & ".ns1"
       'Left(fname, Len(fname) - 4) & "_Merged.ns1"
    Else
        If SingleItem = 0 Then
            rfname = Left(fname, Len(fname) - 4) & "_Recovered" & append & ".ns1"
        Else
            rfname = Left(fname, InStrRev(fname, "\")) & Replace(ns1.apinfo(SingleItem).BSSID, ":", "") & append & ".ns1"
        End If
    End If
    'if the file is there, get rid of it
    'this is a good spot to add some code to backup the original file
    If Dir(rfname) <> "" Then
        Kill rfname
    End If
    
    frmMain.StatusBar1.Panels(1).Text = "Saving Records..."
    
    'Output all data to the file
    
    OpenBinFile rfname
    If SingleItem = 0 Then
        Write_Ns1Header count
    Else
        Write_Ns1Header 1
    End If
    step = 1
    If SingleItem = 0 Then
        'It's much easier to save the entire array directly (without looping),
        'but looping is required to implement a progress bar.
         For apcount = 1 To UBound(ns1.apinfo)
            frmMain.StatusBar1.Panels(1).Text = "Saving Records: " & apcount
            apdone = False
            Do While apdone = False
               ret = Write_ApInfo(ns1, apcount, Offset, step, apdone, merged)
               If ret <> 0 Then Exit For
            Loop
        Next apcount
    Else
        frmMain.StatusBar1.Panels(1).Text = "Saving Records: " & apcount
        apdone = False
        Do While apdone = False
           ret = Write_ApInfo(ns1, SingleItem, Offset, step, apdone, merged)
        Loop
    End If
    CloseFile
    frmMain.mnuOpenNet.Enabled = True
    If count > 0 Then
    frmMain.StatusBar1.Panels(1).Text = count & " records saved successfully... Ready"
    Else
    frmMain.StatusBar1.Panels(1).Text = ns1.apcount - 1 & " records saved successfully... Ready"
    End If
    
End Sub
Public Function RefillListBox(ByRef ns1 As ns1, lstView As ListView, Index As Long, Optional FullLoad As Boolean = False, Optional DataType As Long = 0)

    'Clears, then refills using the array data
    Dim recordno As Long
    Dim lvitem As ListItem
    On Error GoTo ErrHandler
    lstView.Enabled = False
    lstView.Visible = False
    
   If Not FullLoad Then lstView.ListItems.Clear
Select Case DataType
       Case 0 'Apinfo
    With lstView
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , "SSID", "SSID"
        .ColumnHeaders.Add , "BSSID", "BSSID"
        .ColumnHeaders.Add , "MAXSIGNAL", "MaxSignal"
        .ColumnHeaders.Add , "MINNOISE", "MinNoise"
        .ColumnHeaders.Add , "MAXSNR", "MaxSNR"
        .ColumnHeaders.Add , "FLAGS", "flags"
        .ColumnHeaders.Add , "BEACONINTERVAL", "BeaconInterval"
        .ColumnHeaders.Add , "FIRSTSEEN", "FirstSeen"
        .ColumnHeaders.Add , "LASTSEEN", "LastSeen"
        .ColumnHeaders.Add , "BESTLAT", "BestLat"
        .ColumnHeaders.Add , "BESTLONG", "BestLong"
        .ColumnHeaders.Add , "NAME", "Name"
        .ColumnHeaders.Add , "CHANNELS", "Channels"
        .ColumnHeaders.Add , "LASTCHANNEL", "LastChannel"
        .ColumnHeaders.Add , "IPADDRESS", "IPAddress"
        .ColumnHeaders.Add , "MINSIGNAL", "MinSignal"
        .ColumnHeaders.Add , "MAXNOISE", "MaxNoise"
        .ColumnHeaders.Add , "DATARATE", "DataRate"
        .ColumnHeaders.Add , "IPSUBNET", "IPSubnet"
        .ColumnHeaders.Add , "IPMASK", "IPMask"
        .ColumnHeaders.Add , "APFLAGS", "ApFlags"
        .ColumnHeaders.Add , "IE", "InformationElements"
        .ColumnHeaders.Add , "ApDataSize", "ApData Size"
        .ColumnHeaders.Add , "DataSize", "Data Size"
        .ColumnHeaders.Add , "BO", "Begin Offset"
        .ColumnHeaders.Add , "EO", "End Offset"
        
     End With
     If ns1.apinfo(Index).BSSID <> "" Then
        Set lvitem = lstView.ListItems.Add(, ns1.apinfo(Index).BSSID & "|" & Index, ns1.apinfo(Index).SSID)
        With lvitem
           .SubItems(1) = ns1.apinfo(Index).BSSID
           .SubItems(2) = ns1.apinfo(Index).MaxSignal
           .SubItems(3) = ns1.apinfo(Index).MinNoise
           .SubItems(4) = ns1.apinfo(Index).MaxSNR
           .SubItems(5) = Format(CDToH(ns1.apinfo(Index).flags), "00##")
           If Mid(.SubItems(5), 3, 1) = "1" Then
               .Icon = 9
               .SmallIcon = 9
           Else
               .Icon = 6
               .SmallIcon = 6
           End If
           .SubItems(6) = ns1.apinfo(Index).BeaconInterval
           .SubItems(7) = ns1.apinfo(Index).firstseen.Time
           .SubItems(8) = ns1.apinfo(Index).lastseen.Time
           .SubItems(9) = ValToDms(ns1.apinfo(Index).BestLat.dbl, True)
           .SubItems(10) = ValToDms(ns1.apinfo(Index).BestLong.dbl, False)
           .SubItems(11) = ns1.apinfo(Index).Name
           .SubItems(12) = ns1.apinfo(Index).Channels.str
           .SubItems(13) = ns1.apinfo(Index).LastChannel
           .SubItems(14) = ns1.apinfo(Index).IPAddress
           .SubItems(15) = ns1.apinfo(Index).MinSignal
           .SubItems(16) = ns1.apinfo(Index).MaxNoise
           .SubItems(17) = ns1.apinfo(Index).DataRate
           .SubItems(18) = ns1.apinfo(Index).IPSubnet
           .SubItems(19) = ns1.apinfo(Index).IPMask
           .SubItems(20) = Format(CDToH(ns1.apinfo(Index).ApFlags), "00##")
           .SubItems(21) = ns1.apinfo(Index).InformationElements
           
           .SubItems(22) = UBound(ns1.apinfo(Index).APData)
           .SubItems(23) = ns1.apinfo(Index).DataCount
           .SubItems(24) = ns1.apinfo(Index).Offsets.BeginOffset
           .SubItems(25) = ns1.apinfo(Index).Offsets.EndOffset

     End With
      End If
     Case 1 'Apdata
         With lstView
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , "Item", "Item"
            .ColumnHeaders.Add , "TIME", "Time"
            .ColumnHeaders.Add , "SIGNAL", "Signal"
            .ColumnHeaders.Add , "NOISE", "Noise"
            .ColumnHeaders.Add , "Location_Source", "Location Source"
            .ColumnHeaders.Add , "Latitude", "Latitude"
            .ColumnHeaders.Add , "Longitude", "Longitude"
            .ColumnHeaders.Add , "Altitude", "Altitude"
            .ColumnHeaders.Add , "NumSats", "NumSats"
            .ColumnHeaders.Add , "Speed", "Speed"
            .ColumnHeaders.Add , "Track", "Track"
            .ColumnHeaders.Add , "MagVariation", "MagVariation"
            .ColumnHeaders.Add , "Hdop", "Hdop"
         End With
         If ns1.apinfo(Index).DataCount > 0 Then
        For recordno = LBound(ns1.apinfo(Index).APData) To UBound(ns1.apinfo(Index).APData) - 1
            Set lvitem = lstView.ListItems.Add(, Index & "|" & recordno, recordno)
            With lvitem
               .SubItems(1) = ns1.apinfo(Index).APData(recordno).Time.Time
               .SubItems(2) = ns1.apinfo(Index).APData(recordno).Signal
               .SubItems(3) = ns1.apinfo(Index).APData(recordno).Noise
               .SubItems(4) = ns1.apinfo(Index).APData(recordno).Location_Source
               If ns1.apinfo(Index).APData(recordno).Location_Source <> 0 Then
                    .SubItems(5) = ValToDms(ns1.apinfo(Index).APData(recordno).GPSDATA.Latitude.dbl, True)
                    .SubItems(6) = ValToDms(ns1.apinfo(Index).APData(recordno).GPSDATA.Longitude.dbl, False)
                    .SubItems(7) = ns1.apinfo(Index).APData(recordno).GPSDATA.Altitude.dbl
                    .SubItems(8) = ns1.apinfo(Index).APData(recordno).GPSDATA.NumSats
                    .SubItems(9) = ns1.apinfo(Index).APData(recordno).GPSDATA.Speed.dbl
                    .SubItems(10) = ns1.apinfo(Index).APData(recordno).GPSDATA.Track.dbl
                    .SubItems(11) = ns1.apinfo(Index).APData(recordno).GPSDATA.MagVariation.dbl
                    .SubItems(12) = ns1.apinfo(Index).APData(recordno).GPSDATA.Hdop.dbl
              End If
            End With
        Next recordno
        End If
         End Select
     lstView.Enabled = True
     lstView.Visible = True
     Exit Function
ErrHandler:
If gDebugMode Then
   ' Debug.Print "Refill ListBox", Err.Description
    Resume Next
Else
    frmMain.StatusBar1.Panels(1).Text = Err.Description & " error populating listbox"
     Resume Next
    
End If

End Function

