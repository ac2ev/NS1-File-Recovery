Attribute VB_Name = "basExportKML"
Option Explicit
    Public Descriptor As Integer
    Dim Header1 As String
    Dim Header2 As String
    Const Folder1 As String = "<Folder>"
    Const Name1 As String = "<name>"
    Const Name2 As String = "</name>"
    Const Open1 As String = "<open>"
    Const Open2 As String = "</open>"
    Const Placemark1 As String = "<Placemark>"
    Const Description1 As String = "<description>"
    Const Description2 As String = "</description>"
    Const Lookat1 As String = "<LookAt>"
    Const Longitude1 As String = "<longitude>"
    Const Longitude2 As String = "</longitude>"
    Const Latitude1 As String = "<latitude>"
    Const Latitude2 As String = "</latitude>"
    Const Range1 As String = "<range>"
    Const Range2 As String = "</range>"
    Const Tilt1 As String = "<tilt>"
    Const Tilt2 As String = "</tilt>"
    Const Heading1 As String = "<heading>"
    Const Heading2 As String = "</heading>"
    Const Lookat2 As String = "</LookAt>"
    Const Point1 As String = "<Point>"
    Const Coords1 As String = "<coordinates>"
    Const Coords2 As String = "</coordinates>"
    Const Point2 As String = "</Point>"
    Const Placemark2 As String = "</Placemark>"
    Const Folder2 As String = "</Folder>"
    Const Kml2 As String = "</kml>"
    Public ExportItem As Variant
    Public picAdHocNW As String
    Public picAdHocW As String
    Public picAPNW As String
    Public picAPW As String
    Public GroupBy As Integer
    Public Use3D As Integer
        
    Public Group_Peer() As apinfo
    Public Group_AP() As apinfo
    Public Group_Ch0() As apinfo
    Public Group_Ch1() As apinfo
    Public Group_Ch2() As apinfo
    Public Group_Ch3() As apinfo
    Public Group_Ch4() As apinfo
    Public Group_Ch5() As apinfo
    Public Group_Ch6() As apinfo
    Public Group_Ch7() As apinfo
    Public Group_Ch8() As apinfo
    Public Group_Ch9() As apinfo
    Public Group_Ch10() As apinfo
    Public Group_Ch11() As apinfo
    Public Group_Ch12() As apinfo
    Public Group_Ch13() As apinfo
    Public Group_Ch14() As apinfo
    Public Group_WEP() As apinfo
    Public Group_NoWEP() As apinfo
    
Public Function ExportToKML(locfname As String) As Boolean
    Dim Quote As String
    Dim Temp() As Long
    Quote = Chr(34)
    Header1 = "<?xml version=" & Quote & "1.0" & Chr(34) & " encoding=" & Quote & "UTF-8" & Quote & "?>"
    Header2 = "<kml xmlns=" & Quote & "http://earth.google.com/kml/2.0" & Quote & ">"
    Dim fso As FileSystemObject
    Dim ts As TextStream
    Dim i As Long
    EraseArrays
    
    On Error GoTo HANDLE_ERROR
    Set fso = New FileSystemObject
    Set ts = fso.CreateTextFile(locfname, True)
    
    ts.WriteLine Header1 '<?xml version="1.0" encoding="UTF-8"?>
    ts.WriteLine Header2 '<kml xmlns="http://earth.google.com/kml/2.0">
'foo poly
'3D visualized view
ts.WriteLine Folder1 & Mid(fname, InStrRev(fname, "\") + 1, InStrRev(fname, ".") - InStrRev(fname, "\") - 1) & " (" & UBound(ns1.apinfo) & ")"

Select Case GroupBy
    Case 0 'NONE
        WriteData ts, ns1.apinfo, "Log"
    Case 1 'SSID
        WriteData ts, ns1.apinfo, "Log"
    Case 2 'CHANNEL
        ReDim Temp(14)
        For i = 1 To UBound(ns1.apinfo)
            Select Case ns1.apinfo(i).LastChannel
                Case 1
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch1(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch1(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)

                Case 2
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch2(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch2(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 3
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch3(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch3(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 4
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch4(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch4(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 5
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch5(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch5(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 6
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch6(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch6(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 7
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch7(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch7(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 8
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch8(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch8(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 9
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch9(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch9(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 10
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch10(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch10(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 11
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch11(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch11(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 12
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch12(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch12(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 13
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch13(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch13(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case 14
                    Temp(ns1.apinfo(i).LastChannel) = Temp(ns1.apinfo(i).LastChannel) + 1
                    ReDim Preserve Group_Ch14(Temp(ns1.apinfo(i).LastChannel))
                    Group_Ch14(Temp(ns1.apinfo(i).LastChannel)) = ns1.apinfo(i)
                Case Else
                    Temp(0) = Temp(0) + 1
                    ReDim Preserve Group_Ch0(Temp(0))
                    Group_Ch0(Temp(0)) = ns1.apinfo(i)
                    
            End Select
            
        Next i
        ts.WriteLine Folder1
        ts.WriteLine Name1 & Mid(fname, InStrRev(fname, "\") + 1, InStrRev(fname, ".") - InStrRev(fname, "\") - 1) & " (" & UBound(ns1.apinfo) & ")" & Name2
        If IsArrayInit(Group_Ch0) Then WriteData ts, Group_Ch0(), "Channel ? (" & UBound(Group_Ch0()) & ")"
        If IsArrayInit(Group_Ch1) Then WriteData ts, Group_Ch1(), "Channel 1 (" & UBound(Group_Ch1()) & ")"
        If IsArrayInit(Group_Ch2) Then WriteData ts, Group_Ch2(), "Channel 2 (" & UBound(Group_Ch2()) & ")"
        If IsArrayInit(Group_Ch3) Then WriteData ts, Group_Ch3(), "Channel 3 (" & UBound(Group_Ch3()) & ")"
        If IsArrayInit(Group_Ch4) Then WriteData ts, Group_Ch4(), "Channel 4 (" & UBound(Group_Ch4()) & ")"
        If IsArrayInit(Group_Ch5) Then WriteData ts, Group_Ch5(), "Channel 5 (" & UBound(Group_Ch5()) & ")"
        If IsArrayInit(Group_Ch6) Then WriteData ts, Group_Ch6(), "Channel 6 (" & UBound(Group_Ch6()) & ")"
        If IsArrayInit(Group_Ch7) Then WriteData ts, Group_Ch7(), "Channel 7 (" & UBound(Group_Ch7()) & ")"
        If IsArrayInit(Group_Ch8) Then WriteData ts, Group_Ch8(), "Channel 8 (" & UBound(Group_Ch8()) & ")"
        If IsArrayInit(Group_Ch9) Then WriteData ts, Group_Ch9(), "Channel 9 (" & UBound(Group_Ch9()) & ")"
        If IsArrayInit(Group_Ch10) Then WriteData ts, Group_Ch10(), "Channel 10 (" & UBound(Group_Ch10()) & ")"
        If IsArrayInit(Group_Ch11) Then WriteData ts, Group_Ch11(), "Channel 11 (" & UBound(Group_Ch11()) & ")"
        If IsArrayInit(Group_Ch12) Then WriteData ts, Group_Ch12(), "Channel 12 (" & UBound(Group_Ch12()) & ")"
        If IsArrayInit(Group_Ch13) Then WriteData ts, Group_Ch13(), "Channel 13 (" & UBound(Group_Ch13()) & ")"
        If IsArrayInit(Group_Ch14) Then WriteData ts, Group_Ch14(), "Channel 14 (" & UBound(Group_Ch14()) & ")"
        
        ts.WriteLine Folder2
    Case 3 'ENCRYPTION
        ReDim Temp(2)
        For i = 1 To UBound(ns1.apinfo)
            If Mid(Format(CDToH(ns1.apinfo(i).flags), "00##"), 3, 1) = "1" Then
            'WEP
                Temp(1) = Temp(1) + 1
                ReDim Preserve Group_WEP(Temp(1))
                Group_WEP(Temp(1)) = ns1.apinfo(i)
            Else
            'No WEP
                Temp(2) = Temp(2) + 1
                ReDim Preserve Group_NoWEP(Temp(2))
                Group_NoWEP(Temp(2)) = ns1.apinfo(i)
            End If
        Next i
        ts.WriteLine Folder1
        ts.WriteLine Name1 & Mid(fname, InStrRev(fname, "\") + 1, InStrRev(fname, ".") - InStrRev(fname, "\") - 1) & " (" & UBound(ns1.apinfo) & ")" & Name2
        If IsArrayInit(Group_WEP) Then WriteData ts, Group_WEP(), "Encryption (" & UBound(Group_WEP) & ")"
        If IsArrayInit(Group_NoWEP) Then WriteData ts, Group_NoWEP(), "No Encryption (" & UBound(Group_NoWEP) & ")"
        ts.WriteLine Folder2
    Case 4 'MODE
        ReDim Temp(2)
        For i = 1 To UBound(ns1.apinfo)
            If Right(ns1.apinfo(i).flags, 1) = 2 Then
            'ad-hoc
                Temp(1) = Temp(1) + 1
                ReDim Preserve Group_Peer(Temp(1))
                Group_Peer(Temp(1)) = ns1.apinfo(i)
            Else
            'BSS
                Temp(2) = Temp(2) + 1
                ReDim Preserve Group_AP(Temp(2))
                Group_AP(Temp(2)) = ns1.apinfo(i)
            End If
        Next i
        ts.WriteLine Folder1
        ts.WriteLine Name1 & Mid(fname, InStrRev(fname, "\") + 1, InStrRev(fname, ".") - InStrRev(fname, "\") - 1) & " (" & UBound(ns1.apinfo) & ")" & Name2
        If IsArrayInit(Group_Peer) Then WriteData ts, Group_Peer(), "Peer (" & UBound(Group_Peer) & ")"
        If IsArrayInit(Group_AP) Then WriteData ts, Group_AP(), "AP (" & UBound(Group_AP) & ")"
        ts.WriteLine Folder2
    Case Else
        'GroupBy_None ts
End Select

Create3DView ts, ns1.apinfo


'//fly to first network
'//range: zoom level, tilt: view angle
'ts.WriteLine "<LookAt>"
'ts.WriteLine vbTab & "<longitude>" & ns1.apinfo(1).BestLat.dbl & " </longitude>"
'ts.WriteLine vbTab & "<latitude>" & ns1.apinfo(1).BestLong.dbl & "</latitude>"
'ts.WriteLine "<range>1000</range><tilt>54</tilt><heading>-35</heading></LookAt>"

'// output the WarDrive Path in a separate placemark
ts.WriteLine "<Placemark>"
ts.WriteLine "<name>WarSession Path</name>"
ts.WriteLine "<description></description>"
ts.WriteLine "<visibility>0</visibility>"
ts.WriteLine vbTab & "<Style>"
ts.WriteLine vbTab & "<geomColor>BB0000FF</geomColor>"
ts.WriteLine vbTab & "<geomScale>3</geomScale>"
ts.WriteLine vbTab & "</Style>"
ts.WriteLine vbTab & "<LineString>"

ts.WriteLine vbTab & "<tessellate>1</tessellate>"
ts.WriteLine vbTab & "<coordinates>$line</coordinates>"
ts.WriteLine vbTab & "</LineString>"
ts.WriteLine vbTab & "</Placemark>"

'//tesselate adjusts to terrain

ts.WriteLine Folder2
ts.WriteLine Kml2 '</kml>
ts.Close
Set ts = Nothing
Set fso = Nothing
frmMain.StatusBar1.Panels(1).Text = "Finished Export"
ExportToKML = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export to KML failed. Encountered thej following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To KML"
Resume
Set fso = Nothing
Set ts = Nothing
GoTo ExitFunction
End Function

Private Sub EraseArrays()
On Error Resume Next
    Erase Group_Peer()
    Erase Group_AP()
    Erase Group_Ch1()
    Erase Group_Ch2()
    Erase Group_Ch3()
    Erase Group_Ch4()
    Erase Group_Ch5()
    Erase Group_Ch6()
    Erase Group_Ch7()
    Erase Group_Ch8()
    Erase Group_Ch9()
    Erase Group_Ch10()
    Erase Group_Ch11()
    Erase Group_Ch12()
    Erase Group_Ch13()
    Erase Group_Ch14()
    Erase Group_WEP()
    Erase Group_NoWEP()
End Sub
Private Function IsApdataInit(arTest() As APData) As Boolean
    'Check if array is initialized.
    
    On Error GoTo ErrHandler
    Dim intMax As Integer
    
    intMax = UBound(arTest)
    
    IsApdataInit = True
    
exitHandler:
    Exit Function
    
ErrHandler:
    IsApdataInit = False
    Resume exitHandler
End Function

Public Sub Create3DView(ByRef ts As TextStream, ns1Array() As apinfo, Optional SFolderName As String = "3D visualized view")
    On Error GoTo ErrHandler
    Dim lngResults As Long
    Dim i As Long
    Dim x As Long
    Dim intCounter As Long
    Dim intStartRow As Long
    
    Dim AP As apinfo
    Screen.MousePointer = vbHourglass
    Dim Lat As Double
    Dim Lon As Double
   ts.WriteLine Folder1
   ts.WriteLine Name1 & SFolderName & Name2
   ts.WriteLine Open1 & "0" & Open2   '      <open>1</open>
    
    For i = 1 To UBound(ns1Array)
      
        ts.WriteLine vbTab & Placemark1 '    <Placemark>
        With ns1Array(i)
         If Descriptor = 0 Then
             ts.WriteLine vbTab & vbTab & Name1 & "<![CDATA[" & .BSSID & "]]>" & Name2
         ElseIf Descriptor = 1 Then
             ts.WriteLine vbTab & vbTab & Name1 & "<![CDATA[" & .SSID & "]]>" & Name2
         Else
             ts.WriteLine vbTab & vbTab & Name1 & "<![CDATA[]]>" & Name2 '         <name> TEDSROUTER</name>
         End If
         ts.WriteLine vbTab & vbTab & "<visibility>0</visibility>"
         ts.WriteLine vbTab & vbTab & "<open>0</open>"
         ts.WriteLine vbTab & vbTab & "<Style>"
         ts.WriteLine vbTab & vbTab & vbTab & "<LineStyle>"
         ts.WriteLine vbTab & vbTab & vbTab & vbTab & "<width>1.5</width></LineStyle>"
         ts.WriteLine vbTab & vbTab & vbTab & vbTab & "<PolyStyle><color>8f00ff00</color>"
         ts.WriteLine vbTab & vbTab & vbTab & "</PolyStyle>"
         ts.WriteLine vbTab & vbTab & "</Style>"
         ts.WriteLine vbTab & vbTab & "<Polygon>"
         ts.WriteLine vbTab & vbTab & vbTab & "<extrude>1</extrude>"
         ts.WriteLine vbTab & vbTab & vbTab & "<tessellate>0</tessellate>"
         ts.WriteLine vbTab & vbTab & vbTab & "<altitudeMode>relativeToGround</altitudeMode>"
         ts.WriteLine vbTab & vbTab & vbTab & "<outerBoundaryIs>"
         ts.WriteLine vbTab & vbTab & vbTab & "<LinearRing>"
         ts.WriteLine vbTab & vbTab & vbTab & "<extrude>0</extrude><tessellate>0</tessellate>"
         ts.WriteLine vbTab & vbTab & vbTab & "<altitudeMode>clampToGround</altitudeMode>"
         If Not IsApdataInit(.APData) Then
             ts.WriteLine vbTab & vbTab & vbTab & Coords1 & .BestLong.dbl & "," & .BestLat.dbl & ",6" & Coords2
         Else
             ts.Write vbTab & vbTab & vbTab & Coords1
             For x = 0 To UBound(.APData)
              If .APData(x).GPSDATA.Longitude.dbl <> 0 And .APData(x).GPSDATA.Longitude.dbl <> 90# And .APData(x).GPSDATA.Latitude.dbl <> 0 And .APData(x).GPSDATA.Latitude.dbl <> 180# Then
                ts.Write .APData(x).GPSDATA.Longitude.dbl & "," & .APData(x).GPSDATA.Latitude.dbl & "," & IIf(.APData(x).GPSDATA.Altitude.dbl <= 0, 50, .APData(x).GPSDATA.Altitude.dbl) & " "
              End If
             Next x
             ts.Write Coords2 & vbCrLf
         End If
         ts.WriteLine vbTab & vbTab & "</LinearRing>"
         ts.WriteLine vbTab & vbTab & "</outerBoundaryIs>"
         ts.WriteLine vbTab & vbTab & "</Polygon>"
         ts.WriteLine vbTab & Placemark2 '      </Placemark>
         End With
    Next i
ts.WriteLine Folder2
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Resume
End Sub

Public Sub WriteData(ByRef ts As TextStream, ns1Array() As apinfo, SFolderName As String)
    On Error GoTo ErrHandler
    Dim lngResults As Long
    Dim i As Long
    Dim x As Long
    Dim intCounter As Long
    Dim intStartRow As Long
    Dim AP As apinfo
    Screen.MousePointer = vbHourglass
    Dim Lat As Double
    Dim Lon As Double
   ts.WriteLine Folder1
   ts.WriteLine Name1 & SFolderName & Name2
   ts.WriteLine Open1 & "0" & Open2   '      <open>1</open>
    For i = 1 To UBound(ns1Array)
       ts.WriteLine vbTab & Placemark1 '    <Placemark>
       ts.WriteLine vbTab & vbTab & Description1
       With ns1Array(i)
            
            ts.WriteLine vbTab & vbTab & vbTab & "<![CDATA["
    
            If ExportItem(0) Then ts.WriteLine vbTab & vbTab & vbTab & "SSID: " & ReplaceIllegals(.SSID) & " <BR>"
            If .Name <> "" Then If ExportItem(8) Then ts.WriteLine vbTab & vbTab & vbTab & "Name: " & .Name & " <BR>"
            If ExportItem(5) Then ts.WriteLine vbTab & vbTab & vbTab & "FirstSeen: " & .firstseen.Time & " <BR>"
            If ExportItem(6) Then ts.WriteLine vbTab & vbTab & vbTab & "LastSeen: " & .lastseen.Time & " <BR><hr>"
            '_____________________________________
            If ExportItem(1) Then ts.WriteLine vbTab & vbTab & vbTab & "BSSID: " & .BSSID & " <BR>"
            If ExportItem(13) Then ts.WriteLine vbTab & vbTab & vbTab & "Channel: " & .LastChannel & " <BR>"
            If ExportItem(10) Then ts.WriteLine vbTab & vbTab & vbTab & "Channels: " & .Channels.str & " <BR>"
            If ExportItem(9) And .flags <> "" Then
                If Mid(Format(CDToH(.flags), "00##"), 3, 1) = "1" Then
                    ts.WriteLine vbTab & vbTab & vbTab & "<font color=""green"">encryption: WEP</font><br>"
                Else
                 ts.WriteLine vbTab & vbTab & vbTab & "<font color=""red"">no encryption</font><br>"
                End If
            End If
            '_____________________________________
            
            If ExportItem(2) Then ts.WriteLine vbTab & vbTab & vbTab & "Type: " & IIf(Right(.flags, 1) = 2, "ad-hoc", "BSS") & " <BR>"
            
            If ExportItem(3) Or ExportItem(4) Then ts.WriteLine vbTab & vbTab & "<b>GPS coordinates</b><br>"
            If ExportItem(3) Then ts.WriteLine vbTab & vbTab & vbTab & "Latitude: " & ValToDms(.BestLat.dbl, True, True) & " <BR>"
            If ExportItem(4) Then ts.WriteLine vbTab & vbTab & vbTab & "Longitude: " & ValToDms(.BestLong.dbl, False, True) & " <BR><hr>"

            If ExportItem(7) Then ts.WriteLine vbTab & vbTab & vbTab & "SNR: " & "[ " & .MaxSNR & " " & .MaxSignal + 149 & " " & (.MaxSignal + 149) - .MaxSNR & " ]" & " <BR>"
            If ExportItem(11) Then ts.WriteLine vbTab & vbTab & vbTab & "Beacon: " & .BeaconInterval & " <BR>"
            If ExportItem(12) Then ts.WriteLine vbTab & vbTab & vbTab & "DataRate: " & .DataRate & " <BR><hr>"
            
            If ExportItem(14) Then ts.WriteLine vbTab & vbTab & vbTab & "IP Address: " & .IPAddress & " <BR>"
            If ExportItem(15) Then ts.WriteLine vbTab & vbTab & vbTab & "IP Mask: " & .IPMask & " <BR>"
            If ExportItem(16) Then ts.WriteLine vbTab & vbTab & vbTab & "IP SubNet: " & .IPSubnet & " <BR>"
            ts.WriteLine vbTab & vbTab & vbTab & "]]>"
            ts.WriteLine vbTab & vbTab & Description2  '         <description> TEDSROUTER</description>
        If Descriptor = 0 Then
            ts.WriteLine vbTab & vbTab & Name1 & "<![CDATA[" & .BSSID & "]]>" & Name2
        ElseIf Descriptor = 1 Then
            ts.WriteLine vbTab & vbTab & Name1 & "<![CDATA[" & .SSID & "]]>" & Name2
        Else
            ts.WriteLine vbTab & vbTab & Name1 & "<![CDATA[]]>" & Name2 '         <name> TEDSROUTER</name>
        End If
        ts.WriteLine vbTab & vbTab & Lookat1 '         <LookAt>
        ts.WriteLine vbTab & vbTab & vbTab & Longitude1 & .BestLong.dbl & Longitude2  '           <longitude>-95.5211983</longitude>
        ts.WriteLine vbTab & vbTab & vbTab & Latitude1 & .BestLat.dbl & Latitude2   '           <latitude>29.9926667</latitude>
        ts.WriteLine vbTab & vbTab & vbTab & Range1 & "20000" & Range2 '           <range>20000</range>
        ts.WriteLine vbTab & vbTab & vbTab & Tilt1 & "0" & Tilt2 '           <tilt>0</tilt>
        ts.WriteLine vbTab & vbTab & vbTab & Heading1 & "0" & Heading2 '           <heading>0</heading>
        ts.WriteLine vbTab & vbTab & Lookat2 '         </LookAt>
        ts.WriteLine vbTab & vbTab & "<Style>"
        ts.WriteLine vbTab & vbTab & "<IconStyle>"
        ts.WriteLine vbTab & vbTab & "<Icon>"
        
         IconPath = creg.GetRegistryValue("IconPath", App.Path & "\icons")
         If IconPath = App.Path & "\icons" And Not FileExists(App.Path & "\icons", vbDirectory) Then
            MkDir App.Path & "\icons"
            If Not FileExists(App.Path & "\icons\100.PNG") Then ExtractRes App.Path & "\icons", 100, "PNG"
            If Not FileExists(App.Path & "\icons\101.PNG") Then ExtractRes App.Path & "\icons", 101, "PNG"
            If Not FileExists(App.Path & "\icons\102.PNG") Then ExtractRes App.Path & "\icons", 102, "PNG"
            If Not FileExists(App.Path & "\icons\103.PNG") Then ExtractRes App.Path & "\icons", 103, "PNG"
            If Not FileExists(App.Path & "\icons\104.PNG") Then ExtractRes App.Path & "\icons", 104, "PNG"
            If Not FileExists(App.Path & "\icons\105.PNG") Then ExtractRes App.Path & "\icons", 105, "PNG"
         End If
            If GroupBy <> 0 Then
                picAdHocNW = creg.GetRegistryValue("picAdHocNW", App.Path & "\icons\103.png", , , , False)
                picAdHocW = creg.GetRegistryValue("picAdHocW", App.Path & "\icons\105.png", , , , False)
                picAPNW = creg.GetRegistryValue("picAPNW", App.Path & "\icons\101.png", , , , False)
                picAPW = creg.GetRegistryValue("picAPW", App.Path & "\icons\100.png", , , , False)
            Else
                picAdHocNW = App.Path & "\icons\100.png"
                picAdHocW = App.Path & "\icons\101.png"
                picAPNW = App.Path & "\icons\100.png"
                picAPW = App.Path & "\icons\101.png"
            End If
        If Right(.flags, 1) = 2 Then
        '"ad-hoc"
          If Mid(Format(CDToH(.flags), "00##"), 3, 1) = "1" Then
            ts.WriteLine vbTab & vbTab & vbTab & "<href>" & picAdHocW & "</href>"
          Else
            ts.WriteLine vbTab & vbTab & vbTab & "<href>" & picAdHocNW & "</href>"
          End If
        Else
        '"BSS"
         If .flags <> "" Then
         If Mid(Format(CDToH(.flags), "00##"), 3, 1) = "1" Then
            ts.WriteLine vbTab & vbTab & vbTab & "<href>" & picAPW & "</href>"
         Else
            ts.WriteLine vbTab & vbTab & vbTab & "<href>" & picAPNW & "</href>"
         End If
         Else
            ts.WriteLine vbTab & vbTab & vbTab & "<href>" & picAPNW & "</href>"
         End If
        End If
        ts.WriteLine vbTab & vbTab & vbTab & "<w>24</w>"
        ts.WriteLine vbTab & vbTab & vbTab & "<h>24</h>"
        ts.WriteLine vbTab & vbTab & vbTab & "</Icon>"
        ts.WriteLine vbTab & vbTab & vbTab & "<scale>0.5</scale>"
        ts.WriteLine vbTab & vbTab & vbTab & "</IconStyle>"
        ts.WriteLine vbTab & vbTab & vbTab & "<LabelStyle>"
        ts.WriteLine vbTab & vbTab & vbTab & "<scale>0.9</scale>"
        ts.WriteLine vbTab & vbTab & vbTab & "</LabelStyle>"
        ts.WriteLine vbTab & vbTab & vbTab & "</Style>"
        ts.WriteLine vbTab & vbTab & vbTab & Point1 '         <Point>
        ts.WriteLine vbTab & vbTab & vbTab & Coords1 & .BestLong.dbl & "," & _
                             .BestLat.dbl & ",0" & _
                            Coords2 '           <coordinates>-95.5211983,29.9926667,0</coordinates>
        ts.WriteLine vbTab & vbTab & vbTab & Point2 '         </Point>
        ts.WriteLine vbTab & Placemark2  '      </Placemark>
        End With
    Next i
ts.WriteLine Folder2
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Resume
End Sub
Public Function ReplaceIllegals(ByRef sSource As String) As String
 Dim i As Integer
 For i = 1 To Len(sSource)
    If Asc(Mid(sSource, i, 1)) > 127 Then
         sSource = Replace(sSource, Mid(sSource, i, 1), " ")
    End If
 Next
    ReplaceIllegals = sSource
End Function


