Attribute VB_Name = "modMain"
Option Explicit
Public fname As String
Public dontcheckreg As Boolean
Public BatchJob As Boolean
Public BatchNode As MSComctlLib.Node
Public RootNode As MSComctlLib.Node
Public ParentNode As MSComctlLib.Node
Public ChildNode As MSComctlLib.Node
Public LastNode As MSComctlLib.Node

Public Type ItemOffset
    BeginOffset As Long
    EndOffset As Long
End Type
    

Public LastGoodOffset As ItemOffset
Global gDebugMode As Boolean
Global LiveUpdate As Boolean



'API Declarations
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

' get window word constants
 Const GWW_HWNDPARENT = (-8)
'
'------------------------------------------------------------------------------------
'  ROUTINE: InIDE:BOOL,    Params( inhwnd InputOnly )
'  Purpose: to determine if the program is running in the Integrated Development Environment '
'  Description: Uses the class of the hidden parent window to determine if the program
'     is running in the IDE or is compiled into an EXE.
'
'  INPUT: inhwnd -- the window handle of the calling window
'  OUTPUT: return code is True if the program is running in the IDE, False
'  if it is an EXE '
'
'------------------------------------------------------------------------------------
Function InIDE(ByVal hwnd As Long) As Boolean

    Dim parent As Long
    Dim pclass As String
    Dim nlen As Long

    parent = GetWindowLong(hwnd, GWW_HWNDPARENT)
    pclass = Space$(32)
    nlen = GetClassName(parent, pclass, 31)
    pclass = Left$(pclass, nlen)
    If InStr(pclass, "RT") Then
        InIDE = False
    Else
        InIDE = True
    End If

End Function

Public Function SaveLW(lW As ListView, fname As String)
    
    Dim FileId As Integer
    Dim x As Integer
    Dim sIdx As Integer
    Dim i As Integer
    Dim sTextLine As String
    
    sIdx = lW.ColumnHeaders.count - 1
    FileId = FreeFile
    On Error Resume Next
    Open fname For Output As #FileId
    For i = 1 To lW.ColumnHeaders.count
        If i = 1 Then
            sTextLine = lW.ColumnHeaders(i).Text
        Else
            sTextLine = sTextLine & ";" & lW.ColumnHeaders(i).Text
        End If
    Next
        Print #FileId, sTextLine
    
    For i = 1 To lW.ListItems.count
        sTextLine = lW.ListItems.item(i).Text
        For x = 1 To sIdx
            sTextLine = sTextLine & ";" & lW.ListItems.item(i).SubItems(x)
        Next
        Print #FileId, sTextLine
    Next
    Close #FileId
    'Lw.ListItems.Clear

End Function

' Function to Convert Hexadecimal to Decimal
' *************************************************************

Public Function CHToD(BinVal As String) As String
Dim iVal#, Temp#, i%, Length%

Length = Len(BinVal)
For i = 0 To Length - 1
    Temp = HexToNo(Mid(BinVal, Length - i, 1))
    iVal = iVal + (Temp * (16 ^ i))
Next i
CHToD = iVal
End Function

Private Function HexToNo(i As String) As Long
Select Case i
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
        HexToNo = CInt(i)
    Case "A", "a":
        HexToNo = 10
    Case "B", "b":
        HexToNo = 11
    Case "C", "c":
        HexToNo = 12
    Case "D", "d":
        HexToNo = 13
    Case "E", "e":
        HexToNo = 14
    Case "F", "f":
        HexToNo = 15
End Select
End Function

' Function to Convert Decimal to Hexadecimal
' *************************************************************

Public Function CDToH(Value As Variant) As String
Dim iVal#, Temp#, ret%, i%, str$
Dim BinVal$()

iVal = Value
Do
    Temp = iVal / 16
    ret = InStr(Temp, ".")
    If ret > 0 Then
        Temp = Left(Temp, ret - 1)
    End If
    ret = iVal Mod 16
    ReDim Preserve BinVal(i)
    BinVal(i) = NoToHex(ret)
    i = i + 1
    iVal = Temp
Loop While Temp > 0
For i = UBound(BinVal) To 0 Step -1
    str = str + CStr(BinVal(i))
Next
If Len(str) = 1 Then
    str = "0" & str
End If

CDToH = str

End Function


Private Function NoToHex(i As Integer) As String
Select Case i
    Case 0 To 9
        NoToHex = CStr(i)
    Case 10:
        NoToHex = "A"
    Case 11:
        NoToHex = "B"
    Case 12:
        NoToHex = "C"
    Case 13:
        NoToHex = "D"
    Case 14:
        NoToHex = "E"
    Case 15:
        NoToHex = "F"
End Select
End Function



