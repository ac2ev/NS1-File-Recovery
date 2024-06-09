Attribute VB_Name = "modKismet"
Option Explicit

Private Type KismetDate
    strDOW     As String
    strMonth   As String
    intDay     As Integer
    intHour    As Integer
    intMin     As Integer
    intSec     As Integer
    intYear    As Integer
End Type

Public Type Kismet_record
    Network     As Long
    NetType     As String
    ESSID       As String
    BSSID       As String
    Info        As String
    Channel     As Integer
    Cloaked     As String
    Encryption  As String
    Decrypted   As String
    MaxRate     As Long
    MaxSeenRate As Long
    Beacon      As Long
    LLC         As Long
    Data        As Long
    Crypt       As Long
    Weak        As Long
    Total       As Long
    Carrier     As String
    Encoding    As String
    FirstTime   As Date
    LastTime    As Date
    BestQuality As Long
    BestSignal  As Long
    BestNoise   As Long
    GPSMinLat   As Double
    GPSMinLon   As Double
    GPSMinAlt   As Double
    GPSMinSpd   As Double
    GPSMaxLat   As Double
    GPSMaxLon   As Double
    GPSMaxAlt   As Double
    GPSMaxSpd   As Double
    GPSBestLat  As Double
    GPSBestLon  As Double
    GPSBestAlt  As Double
    DataSize    As Long
    IPType      As String
    IP          As String
End Type

Public Type Wiscan_record
    Latitude    As String
    Longitude   As String
    ESSID       As String
    NetType     As String
    BSSID       As String
    Time        As String
    SNR         As String
    MaxSignal   As Long
    MaxSNR      As Long
    MaxNoise    As Long
End Type

Public Type Sniffi_record
    ESSID       As String
    BSSID       As String
    Crypt       As String
    BestSignal  As Long
    Latitude    As String
    Longitude   As String
    NetType     As String 'AP Adhoc
    Channels    As String 'Space delimited string
    LastChannel      As String
    Field9      As String
    Field10     As String
    IP          As String
End Type

Public Type OZIWaypoint_Record
    Field1      As String
    ESSID       As String
    Latitude    As String
    Longitude   As String
    TDateTime   As String
    Symbol      As Long
    Status      As Long
    MapFormat   As String
    ForegroundColor As Long
    BackgroundColor As Long
    BSSID       As String
    PointerDir  As String
    GarminFormat As String
    Proximity   As Long
    Altitude As Long
    FontSize As Long
    SymbolSize As Long
    SymbolPos  As Long
    ProximityTime As String
    ProxRouteBoth As String
    FileName As String
    ProxSymbolName As String
End Type

Public Type Kismet_Header
'// Kismet header information
    strFields()           As String
    strXMLVersion       As String
    strXMLEncoding      As String
    strXMLDocType       As String
    strXMLSystem        As String
    strKismetVersion    As String
    strStartTime        As String
    strEndTime          As String
    strFormatType       As String
End Type

Public Kismet_record() As Kismet_record
Public Wiscan_record() As Wiscan_record
Public OZIWaypoint_Record() As OZIWaypoint_Record
Public Sniffi_record() As Sniffi_record

Public Kismet_Array() As String
Public Wiscan_Array() As String
Public OziWaypoint_Array() As String
Public Sniffi_Array() As String

Public Kismet_Header As Kismet_Header
Public fNameKismet As String
Public fNameWiscan As String
Public fNameOZI As String
Public fNameSniffi As String

Public PrevCount As Long

Sub Read_Kismet(fname As String, isBatch As Boolean)
    Dim i As Long
    Dim icnt As Long
    Dim tmparry() As String
    Dim strTextLine As String
    Dim bGPS    As Boolean
    'Dim fNum As Long
    fNum = FreeFile
    Dim arryKismet() As String
    fNameKismet = fname
    On Error GoTo ErrHandler
    Open fname For Input As fNum
    Line Input #fNum, strTextLine ' Read line into variable.
    Kismet_Header.strFields() = Split(strTextLine, ";")
    frmMain.bLoading = True
    If Dir(Left(fname, Len(fname) - 3) & "gps") <> "" Then bGPS = True
    
    
    Do While Not EOF(fNum)     ' Loop until end of file.
         ReDim Preserve Kismet_record(i)
         Line Input #fNum, strTextLine ' Read line into variable.
         arryKismet() = Split(strTextLine, ";")
        If UBound(arryKismet) <> -1 Then
         ReDim Preserve arryKismet(38)
            frmMain.StatusBar1.Panels(1).Text = "Reading record " & arryKismet(0)
            Kismet_record(i).Network = arryKismet(0)
            Kismet_record(i).NetType = arryKismet(1)
            Kismet_record(i).ESSID = arryKismet(2)
            If IsValidMAC(arryKismet(3)) Then Kismet_record(i).BSSID = arryKismet(3)
            Kismet_record(i).Info = arryKismet(4)
            Kismet_record(i).Channel = arryKismet(5)
            Kismet_record(i).Cloaked = arryKismet(6)
            Kismet_record(i).Encryption = arryKismet(7)
            Kismet_record(i).Decrypted = arryKismet(8)
            Kismet_record(i).MaxRate = arryKismet(9)
            Kismet_record(i).MaxSeenRate = arryKismet(10)
            Kismet_record(i).Beacon = CLng(arryKismet(11)) / 255
            If arryKismet(12) <> "" Then Kismet_record(i).LLC = arryKismet(12)
            If arryKismet(13) <> "" Then Kismet_record(i).Data = CLng(arryKismet(13))
            If arryKismet(14) <> "" Then Kismet_record(i).Crypt = CLng(arryKismet(14))
            If arryKismet(15) <> "" Then Kismet_record(i).Weak = CLng(arryKismet(15))
            If arryKismet(16) <> "" Then Kismet_record(i).Total = CLng(arryKismet(16))
            If arryKismet(17) <> "" Then Kismet_record(i).Carrier = arryKismet(17)
            If arryKismet(18) <> "" Then Kismet_record(i).Encoding = arryKismet(18)
            If arryKismet(19) <> "" Then
                tmparry = Split(arryKismet(19), " ")
                If tmparry(2) = "" Then
                'Thu Jan  2 07:58:37 2005 extra space for month
                    Kismet_record(i).FirstTime = DateSerial(tmparry(5), DateText(tmparry(1)), tmparry(3)) & " " & TimeValue(tmparry(4))
                Else
                'Thu Jan 20 07:58:37 2005
                    Kismet_record(i).FirstTime = DateSerial(tmparry(4), DateText(tmparry(1)), tmparry(2)) & " " & TimeValue(tmparry(3))
                End If
                
                tmparry = Split(arryKismet(20), " ")
                If tmparry(2) = "" Then
                'Thu Jan  2 07:58:37 2005 extra space for month
                    Kismet_record(i).LastTime = DateSerial(tmparry(5), DateText(tmparry(1)), tmparry(3)) & " " & TimeValue(tmparry(4))
                Else
                'Thu Jan 20 07:58:37 2005
                    Kismet_record(i).LastTime = DateSerial(tmparry(4), DateText(tmparry(1)), tmparry(2)) & " " & TimeValue(tmparry(3))
                End If
            End If
            If arryKismet(21) <> "" Then Kismet_record(i).BestQuality = CLng(arryKismet(21))
            If arryKismet(22) <> "" Then Kismet_record(i).BestSignal = CLng(arryKismet(22))
            If arryKismet(23) <> "" Then Kismet_record(i).BestNoise = CLng(arryKismet(23))
            If arryKismet(24) <> "" Then Kismet_record(i).GPSMinLat = CDbl(arryKismet(24))
            If arryKismet(25) <> "" Then Kismet_record(i).GPSMinLon = CDbl(arryKismet(25))
            If arryKismet(26) <> "" Then Kismet_record(i).GPSMinAlt = CDbl(arryKismet(26))
            If arryKismet(27) <> "" Then Kismet_record(i).GPSMinSpd = CDbl(arryKismet(27))
            If arryKismet(28) <> "" Then Kismet_record(i).GPSMaxLat = CDbl(arryKismet(28))
            If arryKismet(29) <> "" Then Kismet_record(i).GPSMaxLon = CDbl(arryKismet(29))
            If arryKismet(30) <> "" Then Kismet_record(i).GPSMaxAlt = CDbl(arryKismet(30))
            If arryKismet(31) <> "" Then Kismet_record(i).GPSMaxSpd = CDbl(arryKismet(31))
            If arryKismet(32) <> "" Then
                If CDbl(arryKismet(32)) <> 0 Then
                    Kismet_record(i).GPSBestLat = CDbl(arryKismet(32))
                Else
                    Kismet_record(i).GPSBestLat = Kismet_record(i).GPSMinLat + (Kismet_record(i).GPSMaxLat - Kismet_record(i).GPSMinLat) / 2
                End If
            End If
            If arryKismet(33) <> "" Then
                If CDbl(arryKismet(33)) <> 0 Then
                    Kismet_record(i).GPSBestLon = CDbl(arryKismet(33))
                Else
                    Kismet_record(i).GPSBestLon = Kismet_record(i).GPSMinLon + (Kismet_record(i).GPSMaxLon - Kismet_record(i).GPSMinLon) / 2
                End If
            End If
            If arryKismet(34) <> "" Then
                If CDbl(arryKismet(34)) <> 0 Then
                    Kismet_record(i).GPSBestAlt = CDbl(arryKismet(34))
                Else
                    Kismet_record(i).GPSBestAlt = Kismet_record(i).GPSMinAlt + (Kismet_record(i).GPSMaxAlt - Kismet_record(i).GPSMinAlt) / 2
                End If
            End If
            If arryKismet(35) <> "" Then Kismet_record(i).DataSize = CLng(arryKismet(35))
            If arryKismet(36) <> "" Then Kismet_record(i).IPType = arryKismet(36)
            If arryKismet(37) <> "" Then
                'Validate IP type
                
                If IsValidIP(arryKismet(37)) Then
                    Kismet_record(i).IP = arryKismet(37)
                Else
                    Kismet_record(i).IP = "0.0.0.0"
                End If
            End If
            
            i = i + 1
            DoEvents
            If frmMain.bAbort Then GoTo Cleanup
    End If
    Loop
    'frmMain.StatusBar1.Panels(1).Text = "Read " & i & " Records"
Cleanup:
    On Error Resume Next
    frmMain.bAbort = False
    frmMain.bLoading = False
    Close fNum
    If frmMain.BatchIndex = 0 Then frmMain.lstView.ListItems.Clear
 '   RefillListBox 0, True, 2
    KismettoNs1 isBatch
    Exit Sub
ErrHandler:
    Dim free As Long
    Dim lcntr
    free = FreeFile
'    On Error Resume Next
    Open App.Path & "\Errorlog.txt" For Output As #free
    Print #free, "----------------------------------------"
    Print #free, "File: " & fname
    Print #free, strTextLine
    Print #free, UBound(tmparry)
    For lcntr = LBound(tmparry) To UBound(tmparry)
         Print #free, lcntr & " --" & tmparry(lcntr)
    Next
    Print #free, "i" & i
    Print #free, "+++++++++++++++++++++++++++++++++++++++++"
    Close #free
    Dim ret As Long
    ret = MsgBox("Error Number: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Please email author Kismet file and the file: " & App.Path & "\Errorlog.txt", vbCritical + vbRetryCancel, "Error")
    If ret = vbRetry Then
     Resume Next
    End If
    Resume Cleanup
End Sub

Sub Read_WiScan(fname As String, isBatch As Boolean)
    Dim i As Long
    Dim icnt As Long
    Dim tmparry() As String
    Dim Header() As String
    Dim LastDate As String
    Dim strTextLine As String
'    Dim fNum As Long
    fNum = FreeFile
    Dim arryWiScan() As String
    fNameWiscan = fname
    On Error GoTo ErrHandler
    frmMain.bLoading = True
    Open fname For Input As fNum
    Do While Not EOF(fNum)     ' Loop until end of file.
         
         strTextLine = ""
         
         While Right$(strTextLine, 1) <> Chr(10)
            strTextLine = strTextLine & Input(1, #fNum)
           ' Debug.Print Asc(Right$(strTextLine, 1))
         Wend
         strTextLine = Left$(strTextLine, Len(strTextLine) - 1)
         strTextLine = Replace(Replace(strTextLine, Chr(10), ""), Chr(13), "")
         'Line Input #fNum, strTextLine ' Read line into variable.
    If Left$(strTextLine, 11) = "# $DateGMT:" Then
    LastDate = Replace(Replace(Replace(strTextLine, "# $DateGMT: ", ""), Chr(10), ""), Chr(13), "")
    If Not IsValidDate(LastDate) Then LastDate = Format(Now, "MM/DD/YY")
    End If
    
    If Left$(strTextLine, 1) <> "#" Then
         arryWiScan() = Split(strTextLine, vbTab)
         ReDim Preserve Wiscan_record(i)
            frmMain.StatusBar1.Panels(1).Text = "Reading " & arryWiScan(2)
            With Wiscan_record(i)
                .Latitude = Replace(Replace(arryWiScan(0), "N", ""), "S", "-")
                .Longitude = Replace(Replace(arryWiScan(1), "E", ""), "W", "-")
                .ESSID = Trim(Mid(arryWiScan(2), 2, Len(arryWiScan(2)) - 2))
                .NetType = arryWiScan(3)
                .BSSID = Trim(Mid(arryWiScan(4), 2, Len(arryWiScan(4)) - 2))
                .Time = CDate(LastDate & " " & Replace(arryWiScan(5), "(GMT)", ""))
                .SNR = Trim(Replace(Replace(arryWiScan(6), "[", ""), "]", ""))
                .MaxSignal = Mid(.SNR, InStr(1, .SNR, " "), InStrRev(.SNR, " ") - InStr(1, .SNR, " "))
                .MaxSNR = Left(.SNR, InStr(1, .SNR, " "))
                .MaxNoise = Right$(.SNR, Len(.SNR) - InStrRev(.SNR, " "))
                
                'Trim(Mid(.SNR, InStr(1, .SNR, " "), InStrRev(.SNR, " ") - InStr(1, .SNR, " ")))
            End With
            '"[ " & .MaxSNR & " " & .MaxSignal + 149 & " " & (.MaxSignal + 149) - .MaxSNR & " ]"
            DoEvents
            If frmMain.bAbort Then GoTo Cleanup
            i = i + 1
    End If
    Loop
   ' frmMain.StatusBar1.Panels(1).Text = "Read " & i & " Records"
Cleanup:
    On Error Resume Next
    frmMain.bAbort = False
    frmMain.bLoading = False
    Close fNum
    If frmMain.BatchIndex = 0 Then frmMain.lstView.ListItems.Clear
    WiscantoNs1 isBatch
    Exit Sub
ErrHandler:
     On Error Resume Next
    Dim free As Long
    Dim lcntr
    Close #free
    free = FreeFile
'    On Error Resume Next
'    Open App.Path & "\Errorlog.txt" For Output As #free
'    Print #free, "----------------------------------------"
'    Print #free, "File: " & fname
'    Print #free, strTextLine
'    Print #free, UBound(tmparry)
'    For lcntr = LBound(tmparry) To UBound(tmparry)
'         Print #free, lcntr & " --" & tmparry(lcntr)
'    Next
'    Print #free, "i" & i
'    Print #free, "+++++++++++++++++++++++++++++++++++++++++"
'    Close #free
    Dim ret As Long
    ret = MsgBox("Error Number: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Please email author WiScan file and the file: " & App.Path & "\Errorlog.txt", vbCritical + vbRetryCancel, "Error")
    If ret = vbRetry Then
     Resume Next
    End If
    Resume Cleanup
End Sub

Sub Read_OZI(fname As String, isBatch As Boolean)
    Dim i As Long
    Dim icnt As Long
    Dim tmparry() As String
    Dim Header() As String
    Dim LastDate As String
    Dim strTextLine As String
'    Dim fNum As Long
    fNum = FreeFile
    Dim arryOZI() As String
    fNameOZI = fname
    On Error GoTo ErrHandler
    frmMain.bLoading = True
    Open fname For Input As fNum
    Do While Not EOF(fNum)     ' Loop until end of file.
         
         strTextLine = ""
         
         While Right$(strTextLine, 1) <> Chr(10)
            strTextLine = strTextLine & Input(1, #fNum)
           ' Debug.Print Asc(Right$(strTextLine, 1))
         Wend
         strTextLine = Left$(strTextLine, Len(strTextLine) - 1)
         strTextLine = Replace(Replace(strTextLine, Chr(10), ""), Chr(13), "")
         'Line Input #fNum, strTextLine ' Read line into variable.

    LastDate = Format(Now, "MM/DD/YY")
    End If
    
    If Left$(strTextLine, 2) = "-1" Then
         arryOZI() = Split(strTextLine, ",")
         ReDim Preserve OZI_record(i)
            frmMain.StatusBar1.Panels(1).Text = "Reading " & arryOZI(2)
            With OZIWaypoint_Record(i)
                
                .Latitude = arryOZI(2)
                .Longitude = arryOZI(3)
                .ESSID = arryOZI(1)
                .BSSID = arryOZI(10)
                .Time = CDate(LastDate)
            End With
            DoEvents
            If frmMain.bAbort Then GoTo Cleanup
            i = i + 1
    End If
    Loop
   ' frmMain.StatusBar1.Panels(1).Text = "Read " & i & " Records"
Cleanup:
    On Error Resume Next
    frmMain.bAbort = False
    frmMain.bLoading = False
    Close fNum
    If frmMain.BatchIndex = 0 Then frmMain.lstView.ListItems.Clear
    OZItoNs1 isBatch
    Exit Sub
ErrHandler:
     On Error Resume Next
    Dim free As Long
    Dim lcntr
    Close #free
    free = FreeFile
'    On Error Resume Next
'    Open App.Path & "\Errorlog.txt" For Output As #free
'    Print #free, "----------------------------------------"
'    Print #free, "File: " & fname
'    Print #free, strTextLine
'    Print #free, UBound(tmparry)
'    For lcntr = LBound(tmparry) To UBound(tmparry)
'         Print #free, lcntr & " --" & tmparry(lcntr)
'    Next
'    Print #free, "i" & i
'    Print #free, "+++++++++++++++++++++++++++++++++++++++++"
'    Close #free
    Dim ret As Long
    ret = MsgBox("Error Number: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Please email author OZI file and the file: " & App.Path & "\Errorlog.txt", vbCritical + vbRetryCancel, "Error")
    If ret = vbRetry Then
     Resume Next
    End If
    Resume Cleanup
End Sub
Sub Read_Sniffi(fname As String, isBatch As Boolean)
    Dim i As Long
    Dim icnt As Long
    Dim tmparry() As String
    Dim Header() As String
    Dim LastDate As String
    Dim strTextLine As String
'    Dim fNum As Long
    fNum = FreeFile
    Dim arrySniffi() As String
    fNameSniffi = fname
    On Error GoTo ErrHandler
    frmMain.bLoading = True
    Open fname For Input As fNum
    Do While Not EOF(fNum)     ' Loop until end of file.
         
         strTextLine = ""
         
         While Right$(strTextLine, 1) <> Chr(10)
            strTextLine = strTextLine & Input(1, #fNum)
           ' Debug.Print Asc(Right$(strTextLine, 1))
         Wend
         strTextLine = Left$(strTextLine, Len(strTextLine) - 1)
         strTextLine = Replace(Replace(strTextLine, Chr(10), ""), Chr(13), "")
         'Line Input #fNum, strTextLine ' Read line into variable.

    LastDate = Format(Now, "MM/DD/YY")
  '  End If
    
   ' If Left$(strTextLine, 2) = "-1" Then
         arrySniffi() = Split(strTextLine, ",")
         ReDim Preserve Sniffi_record(i)
            frmMain.StatusBar1.Panels(1).Text = "Reading " & arrySniffi(0)
            With Sniffi_record(i)
                .ESSID = arrySniffi(0)
                .BSSID = arrySniffi(1)
                .Crypt = arrySniffi(2)
                .BestSignal = arrySniffi(3)
                .Latitude = arrySniffi(4)
                .Longitude = arrySniffi(5)
                .Channels = Replace(Trim(arrySniffi(7)), " ", ",")
                .LastChannel = arrySniffi(8)
                .IP = arrySniffi(11)
            End With
            DoEvents
            If frmMain.bAbort Then GoTo Cleanup
            i = i + 1
  '  End If
    Loop
   ' frmMain.StatusBar1.Panels(1).Text = "Read " & i & " Records"
Cleanup:
    On Error Resume Next
    frmMain.bAbort = False
    frmMain.bLoading = False
    Close fNum
    If frmMain.BatchIndex = 0 Then frmMain.lstView.ListItems.Clear
    SniffitoNs1 isBatch
    Exit Sub
ErrHandler:
     On Error Resume Next
    Dim free As Long
    Dim lcntr
    Close #free
    free = FreeFile
'    On Error Resume Next
'    Open App.Path & "\Errorlog.txt" For Output As #free
'    Print #free, "----------------------------------------"
'    Print #free, "File: " & fname
'    Print #free, strTextLine
'    Print #free, UBound(tmparry)
'    For lcntr = LBound(tmparry) To UBound(tmparry)
'         Print #free, lcntr & " --" & tmparry(lcntr)
'    Next
'    Print #free, "i" & i
'    Print #free, "+++++++++++++++++++++++++++++++++++++++++"
'    Close #free
    Dim ret As Long
    ret = MsgBox("Error Number: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Please email author Sniffi file and the file: " & App.Path & "\Errorlog.txt", vbCritical + vbRetryCancel, "Error")
    If ret = vbRetry Then
     Resume Next
    End If
    Resume Cleanup
End Sub
'       Flag                Meaning             Defined in
'Decimal    Hexadecimal
'1          1               ESS                 802.11
'2          2               IBSS                802.11
'4          4               CF Pollable         802.11
'8          8               CF-Poll Request     802.11
'16         10              Privacy (WEP)       802.11
'32         20              Short Preamble      802.11b
'64         40              PBCC                802.11b
'128        80              Channel Agility     802.11b
'1024       400             Short Slot Time     802.11g
'8192       2000            DSSS-OFDM           802.11g
'256, 512, 2048, 4096, 16384, 32768 100, 200, 800, 1000, 4000, 8000 Reserved for future use.
'1    ESS ("Infrastructure")
'2    IBSS ("Ad-Hoc")
'4    cf -Pollable
'0008 CF-Poll Request
'10   Privacy ("WEP")
'20   Short Preamble
'40   PBCC
'80   Channel Agility
'FF00 Reserved
'IEEE 802.11 wireless LAN management frame

'      bssid     BSSID (MAC address) of the network
'      channel   Last-advertised channel for network
'      clients   Number of clients (unique MACs) seen on network
'      crypt     Number of encrypted packets
'      data      Number of data packets
'      decay Displays     '!' or '.' based on network activity
'      dupeiv    Number of packets with duplicate IVs seen
'      flags     Network status flags (Address size, decrypted, etc)
'      info      Extra AP info included by some manufacturers
'      ip        Detected/guessed IP of the network
'      llc       Number of LLC packets
'      manuf     Manufacturer, if matched
'      maxrate   Maximum supported rate as advertised by AP
'      name      Name of the network or group
'      noise     Last seen noise level
'      packets   Total number of packets
'      shortname Shortened name of the network or group for small displays
'      shortssid Shortened SSID for small displays
'      signal    Last seen signal level
'      signalbar Graphical representation of signal level
'      size      Amount of data transfered on network
'      ssid      SSID/ESSID of the network or group
'      type      Network type (Probe, Adhoc, Infra, etc)
'      weak      Number of packets which appear to have weak IVs
'      wep       WEP status (does network indicate it uses WEP)

Public Sub KismettoNs1(Optional isBatch As Boolean = False)
Dim recordno As Long
Dim blankns1 As ns1
Dim tmplng As Long
Dim step As Long
Dim apcount As Long
Dim apdone As Boolean
Dim Offset As Long
Dim ret As Long
Dim alng() As Long
Dim bytes() As Byte
Dim strtmp As String
Dim dbltmp As Double
On Error GoTo ErrHandler

Dim i As Long
If (Not isBatch) Or (frmMain.BatchIndex = 0 And isBatch) Then
    ns1 = blankns1
    PrevCount = 0
    frmMain.treView.Nodes.Clear
    Set RootNode = frmMain.treView.Nodes.Add(, , "Root", "SSID", 2, 2)
    frmMain.lstView.ListItems.Clear
End If

    ns1.apcount = ns1.apcount + UBound(Kismet_record) + 1
    PrevCount = PrevCount + 1
    ns1.dwFileVer = 12
    ns1.dwSignature = "NetS"
    ReDim Preserve ns1.apinfo(ns1.apcount)
    frmMain.lblHeaderValue(0).Caption = ns1.dwSignature
    frmMain.lblHeaderValue(1).Caption = ns1.dwFileVer
    frmMain.lblHeaderValue(2).Caption = ns1.apcount

    For recordno = LBound(Kismet_record) To UBound(Kismet_record)
        ns1.apinfo(recordno + PrevCount).SSIDLength = Len(Kismet_record(recordno).ESSID)
        ns1.apinfo(recordno + PrevCount).SSID = Kismet_record(recordno).ESSID
        ns1.apinfo(recordno + PrevCount).BSSID = Kismet_record(recordno).BSSID
        ns1.apinfo(recordno + PrevCount).MaxSignal = Kismet_record(recordno).BestSignal
        ns1.apinfo(recordno + PrevCount).MinNoise = Kismet_record(recordno).BestNoise
        ns1.apinfo(recordno + PrevCount).MaxSNR = Kismet_record(recordno).BestQuality
        tmplng = 0
        If Kismet_record(recordno).NetType = "infrastructure" Then tmplng = tmplng Or &H1
        If Kismet_record(recordno).NetType = "ad-hoc" Or Kismet_record(recordno).NetType = "probe" Then tmplng = tmplng Or &H2
        If UCase(Left(Kismet_record(recordno).Encryption, 2)) <> "NO" Then tmplng = tmplng Or &H10
        ns1.apinfo(recordno + PrevCount).flags = CStr(tmplng)
        ns1.apinfo(recordno + PrevCount).BeaconInterval = Kismet_record(recordno).Beacon
        ns1.apinfo(recordno + PrevCount).firstseen.Time = Kismet_record(recordno).FirstTime
        ns1.apinfo(recordno + PrevCount).firstseen.File_Time = UtcFromLocalFileTime(FileTimeFromDate(Format(Kismet_record(recordno).FirstTime, "MM/DD/YYYY h:mm:ss AM/PM")))
        ns1.apinfo(recordno + PrevCount).lastseen.Time = Kismet_record(recordno).LastTime
        ns1.apinfo(recordno + PrevCount).lastseen.File_Time = UtcFromLocalFileTime(FileTimeFromDate(Format(Kismet_record(recordno).LastTime, "MM/DD/YYYY h:mm:ss AM/PM")))
        ns1.apinfo(recordno + PrevCount).BestLat.dbl = Kismet_record(recordno).GPSBestLat
        ns1.apinfo(recordno + PrevCount).BestLat.bytes = DoubleToByteArray(ns1.apinfo(recordno).BestLat.dbl)
        ns1.apinfo(recordno + PrevCount).BestLong.dbl = Kismet_record(recordno).GPSBestLon
        ns1.apinfo(recordno + PrevCount).BestLong.bytes = DoubleToByteArray(ns1.apinfo(recordno).BestLong.dbl)
        ns1.apinfo(recordno + PrevCount).DataCount = 0
        ns1.apinfo(recordno + PrevCount).NameLength = Len(Kismet_record(recordno).Info)
        ns1.apinfo(recordno + PrevCount).Name = Kismet_record(recordno).Info
        ns1.apinfo(recordno + PrevCount).Channels.str = Kismet_record(recordno).Channel
        If CLng(Kismet_record(recordno).Channel) <= 14 Then bytes() = LongToByteArray(2 ^ CLng(Kismet_record(recordno).Channel))
        ns1.apinfo(recordno + PrevCount).Channels.bytes = bytes
        ReDim Preserve ns1.apinfo(recordno + PrevCount).Channels.bytes(7)
        ns1.apinfo(recordno + PrevCount).LastChannel = Kismet_record(recordno).Channel
        ns1.apinfo(recordno + PrevCount).IPAddress = Kismet_record(recordno).IP
        ns1.apinfo(recordno + PrevCount).MinSignal = Kismet_record(recordno).BestSignal
        ns1.apinfo(recordno + PrevCount).MaxNoise = Kismet_record(recordno).BestNoise
        ns1.apinfo(recordno + PrevCount).DataRate = Kismet_record(recordno).MaxRate * 10
        ns1.apinfo(recordno + PrevCount).IPSubnet = "0.0.0.0"
        ns1.apinfo(recordno + PrevCount).IPMask = "0.0.0.0"
        ns1.apinfo(recordno + PrevCount).ApFlags = 0
        ns1.apinfo(recordno + PrevCount).IELength = 0
        ns1.apinfo(recordno + PrevCount).InformationElements = 0
        'causes error first time?
        frmMain.FillTreeview recordno + PrevCount
    Next recordno
    PrevCount = ns1.apcount
    Exit Sub
ErrHandler:
    MsgBox Err.Description
   ' Resume
End Sub
Public Sub WiscantoNs1(Optional isBatch As Boolean = False)
Dim recordno As Long
Dim blankns1 As ns1
Dim tmplng As Long
Dim step As Long
Dim apcount As Long
Dim apdone As Boolean
Dim Offset As Long
Dim ret As Long
Dim alng() As Long
Dim bytes() As Byte
Dim strtmp As String
Dim dbltmp As Double
On Error GoTo ErrHandler

Dim i As Long
ns1 = blankns1
If (Not isBatch) Or (frmMain.BatchIndex = 0 And isBatch) Then
    ns1 = blankns1
    PrevCount = 0
    frmMain.treView.Nodes.Clear
    Set RootNode = frmMain.treView.Nodes.Add(, , "Root", "SSID", 2, 2)
    frmMain.lstView.ListItems.Clear
End If

    ns1.apcount = ns1.apcount + UBound(Wiscan_record) + 1
    ns1.dwFileVer = 12
    ns1.dwSignature = "NetS"
'    PrevCount = PrevCount + 1
    ReDim Preserve ns1.apinfo(ns1.apcount)
    frmMain.lblHeaderValue(0).Caption = ns1.dwSignature
    frmMain.lblHeaderValue(1).Caption = ns1.dwFileVer
    frmMain.lblHeaderValue(2).Caption = ns1.apcount

    For recordno = LBound(Wiscan_record) To UBound(Wiscan_record)
        ns1.apinfo(recordno + PrevCount).SSIDLength = Len(Wiscan_record(recordno).ESSID)
        ns1.apinfo(recordno + PrevCount).SSID = Wiscan_record(recordno).ESSID
        ns1.apinfo(recordno + PrevCount).BSSID = Wiscan_record(recordno).BSSID
        ns1.apinfo(recordno + PrevCount).MaxSignal = Wiscan_record(recordno).MaxSignal
        ns1.apinfo(recordno + PrevCount).MinNoise = Wiscan_record(recordno).MaxNoise
        ns1.apinfo(recordno + PrevCount).MaxSNR = Wiscan_record(recordno).MaxSNR
        tmplng = 0
        If Wiscan_record(recordno + PrevCount).NetType = "infrastructure" Then tmplng = tmplng Or &H1
        If Wiscan_record(recordno + PrevCount).NetType = "ad-hoc" Then tmplng = tmplng Or &H2
        ns1.apinfo(recordno + PrevCount).flags = CStr(&H10)
        ns1.apinfo(recordno + PrevCount).BeaconInterval = 0
        ns1.apinfo(recordno + PrevCount).firstseen.Time = Wiscan_record(recordno).Time
        ns1.apinfo(recordno + PrevCount).firstseen.File_Time = UtcFromLocalFileTime(FileTimeFromDate(Format(Wiscan_record(recordno).Time, "MM/DD/YYYY h:mm:ss AM/PM")))
        ns1.apinfo(recordno + PrevCount).lastseen.Time = Wiscan_record(recordno).Time
        ns1.apinfo(recordno + PrevCount).lastseen.File_Time = UtcFromLocalFileTime(FileTimeFromDate(Format(Wiscan_record(recordno).Time, "MM/DD/YYYY h:mm:ss AM/PM")))
        ns1.apinfo(recordno + PrevCount).BestLat.dbl = Wiscan_record(recordno).Latitude
        ns1.apinfo(recordno + PrevCount).BestLat.bytes = DoubleToByteArray(ns1.apinfo(recordno).BestLat.dbl)
        ns1.apinfo(recordno + PrevCount).BestLong.dbl = Wiscan_record(recordno).Longitude
        ns1.apinfo(recordno + PrevCount).BestLong.bytes = DoubleToByteArray(ns1.apinfo(recordno).BestLong.dbl)
        ns1.apinfo(recordno + PrevCount).DataCount = 0
        ns1.apinfo(recordno + PrevCount).NameLength = 0
        ns1.apinfo(recordno + PrevCount).Name = ""
        ns1.apinfo(recordno + PrevCount).Channels.str = 0
        If CLng(ns1.apinfo(recordno + PrevCount).Channels.str) <= 14 Then bytes() = LongToByteArray(CLng(2 ^ CLng(ns1.apinfo(recordno + PrevCount).Channels.str)))
        ns1.apinfo(recordno + PrevCount).Channels.bytes = bytes
        ReDim Preserve ns1.apinfo(recordno + PrevCount).Channels.bytes(7)
        ns1.apinfo(recordno + PrevCount).LastChannel = ns1.apinfo(recordno + PrevCount).Channels.str
        ns1.apinfo(recordno + PrevCount).IPAddress = "0.0.0.0"
        ns1.apinfo(recordno + PrevCount).MinSignal = Wiscan_record(recordno + PrevCount).MaxSignal
        ns1.apinfo(recordno + PrevCount).MaxNoise = Wiscan_record(recordno + PrevCount).MaxSNR
        ns1.apinfo(recordno + PrevCount).DataRate = 10
        ns1.apinfo(recordno + PrevCount).IPSubnet = "0.0.0.0"
        ns1.apinfo(recordno + PrevCount).IPMask = "0.0.0.0"
        ns1.apinfo(recordno + PrevCount).ApFlags = 0
        ns1.apinfo(recordno + PrevCount).IELength = 0
        ns1.apinfo(recordno + PrevCount).InformationElements = 0
        frmMain.FillTreeview recordno + PrevCount
    Next recordno
    PrevCount = ns1.apcount
    Exit Sub
ErrHandler:
    MsgBox Err.Description
    Resume
End Sub


Public Sub SniffitoNs1(Optional isBatch As Boolean = False)
Dim recordno As Long
Dim blankns1 As ns1
Dim tmplng As Long
Dim step As Long
Dim apcount As Long
Dim apdone As Boolean
Dim Offset As Long
Dim ret As Long
Dim alng() As Long
Dim bytes() As Byte
Dim strtmp As String
Dim dbltmp As Double
On Error GoTo ErrHandler

Dim i As Long
ns1 = blankns1
If (Not isBatch) Or (frmMain.BatchIndex = 0 And isBatch) Then
    ns1 = blankns1
    PrevCount = 0
    frmMain.treView.Nodes.Clear
    Set RootNode = frmMain.treView.Nodes.Add(, , "Root", "SSID", 2, 2)
    frmMain.lstView.ListItems.Clear
End If

    ns1.apcount = ns1.apcount + UBound(Sniffi_record) + 1
    ns1.dwFileVer = 12
    ns1.dwSignature = "NetS"
'    PrevCount = PrevCount + 1
    ReDim Preserve ns1.apinfo(ns1.apcount)
    frmMain.lblHeaderValue(0).Caption = ns1.dwSignature
    frmMain.lblHeaderValue(1).Caption = ns1.dwFileVer
    frmMain.lblHeaderValue(2).Caption = ns1.apcount

    For recordno = LBound(Sniffi_record) To UBound(Sniffi_record)
        ns1.apinfo(recordno + PrevCount).SSIDLength = Len(Sniffi_record(recordno).ESSID)
        ns1.apinfo(recordno + PrevCount).SSID = Sniffi_record(recordno).ESSID
        ns1.apinfo(recordno + PrevCount).BSSID = Sniffi_record(recordno).BSSID
        ns1.apinfo(recordno + PrevCount).MaxSignal = Sniffi_record(recordno).BestSignal
        ns1.apinfo(recordno + PrevCount).MinNoise = 0
        ns1.apinfo(recordno + PrevCount).MaxSNR = 0
        tmplng = 0
        If Sniffi_record(recordno + PrevCount).NetType = "AP" Then tmplng = tmplng Or &H1
        If Sniffi_record(recordno + PrevCount).NetType = "AdHoc" Then tmplng = tmplng Or &H2
        ns1.apinfo(recordno + PrevCount).flags = CStr(&H10)
        ns1.apinfo(recordno + PrevCount).BeaconInterval = 0
        ns1.apinfo(recordno + PrevCount).firstseen.Time = Now
        ns1.apinfo(recordno + PrevCount).firstseen.File_Time = UtcFromLocalFileTime(FileTimeFromDate(Format(Now, "MM/DD/YYYY h:mm:ss AM/PM")))
        ns1.apinfo(recordno + PrevCount).lastseen.Time = Now
        ns1.apinfo(recordno + PrevCount).lastseen.File_Time = UtcFromLocalFileTime(FileTimeFromDate(Format(Now, "MM/DD/YYYY h:mm:ss AM/PM")))
        ns1.apinfo(recordno + PrevCount).BestLat.dbl = Sniffi_record(recordno).Latitude
        ns1.apinfo(recordno + PrevCount).BestLat.bytes = DoubleToByteArray(ns1.apinfo(recordno).BestLat.dbl)
        ns1.apinfo(recordno + PrevCount).BestLong.dbl = Sniffi_record(recordno).Longitude
        ns1.apinfo(recordno + PrevCount).BestLong.bytes = DoubleToByteArray(ns1.apinfo(recordno).BestLong.dbl)
        ns1.apinfo(recordno + PrevCount).DataCount = 0
        ns1.apinfo(recordno + PrevCount).NameLength = 0
        ns1.apinfo(recordno + PrevCount).Name = ""
        ns1.apinfo(recordno + PrevCount).Channels.str = Sniffi_record(recordno).Channels
        If Sniffi_record(recordno).Channels <> "" Then If CLng(Sniffi_record(recordno).Channels) <= 14 Then bytes() = LongToByteArray(2 ^ CLng(Sniffi_record(recordno).Channels))
        ns1.apinfo(recordno + PrevCount).Channels.bytes = bytes
        ReDim Preserve ns1.apinfo(recordno + PrevCount).Channels.bytes(7)
        ns1.apinfo(recordno + PrevCount).LastChannel = Sniffi_record(recordno).LastChannel
   
        ns1.apinfo(recordno + PrevCount).IPAddress = Sniffi_record(recordno).IP
        ns1.apinfo(recordno + PrevCount).MinSignal = Sniffi_record(recordno + PrevCount).BestSignal
        ns1.apinfo(recordno + PrevCount).MaxNoise = 0
        ns1.apinfo(recordno + PrevCount).DataRate = 10
        ns1.apinfo(recordno + PrevCount).IPSubnet = "0.0.0.0"
        ns1.apinfo(recordno + PrevCount).IPMask = "0.0.0.0"
        ns1.apinfo(recordno + PrevCount).ApFlags = 0
        ns1.apinfo(recordno + PrevCount).IELength = 0
        ns1.apinfo(recordno + PrevCount).InformationElements = 0
        frmMain.FillTreeview recordno + PrevCount
    Next recordno
    PrevCount = ns1.apcount
    Exit Sub
ErrHandler:
    MsgBox Err.Description
    Resume Next
End Sub

