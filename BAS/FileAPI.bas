Attribute VB_Name = "modFileAPI"
Option Explicit

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const CREATE_ALWAYS = 2
Private Const OPEN_ALWAYS = 4
Private Const INVALID_HANDLE_VALUE = -1


Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, _
   lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
   lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
  ByVal hObject As Long) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
  ByVal hFile As Long, lpBuffer As Any, _
  ByVal nNumberOfBytesToWrite As Long, _
  lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function CreateFile Lib "kernel32" _
  Alias "CreateFileA" (ByVal lpFileName As String, _
  ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
  ByVal lpSecurityAttributes As Long, _
  ByVal dwCreationDisposition As Long, _
  ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) _
  As Long

Private Declare Function FlushFileBuffers Lib "kernel32" ( _
  ByVal hFile As Long) As Long
Public Function WriteStringToFile(FileName As String, ByVal TheData As String, _
    Optional NoOverwrite As Boolean = False) As Boolean
'***************************************************************
'PURPOSE:  WRITES STRING DATA TO FILE USING WRITEFILE API

'PARAMETERS:    FileName: Name Of File
'               TheData: String Data to write to file
'               NoOverwrite (Optional): If set to true, will exit and
'                   return false if file exists.  This function does not
'                   work for appending to a file

'RETURNS:       True If Successful, false otherwise

'EXAMPLE:       WriteStringToFile "C:\MyFile.txt", "Hello World"
'**************************************************************
        

Dim lHandle As Long
Dim lSuccess As Long
Dim lBytesWritten As Long, lBytesToWrite As Long
If NoOverwrite = True And Dir(FileName) <> "" Then Exit Function
lBytesToWrite = Len(TheData)
lHandle = CreateFile(FileName, GENERIC_WRITE Or GENERIC_READ, _
                     0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)

If lHandle <> INVALID_HANDLE_VALUE Then
   lSuccess = WriteFile(lHandle, ByVal TheData, _
                        lBytesToWrite, lBytesWritten, 0) <> 0
   If lSuccess <> 0 Then
      'Flush the file buffers (not sure if this is necessary)
      lSuccess = FlushFileBuffers(lHandle)
      'Close the file.
      lSuccess = CloseHandle(lHandle)
   End If
End If
ErrorHandler:
WriteStringToFile = lSuccess <> 0
End Function
Public Function ReadStringFromFile(FileName As String) As String
'***************************************************************
'PURPOSE:  READS STRING DATA FROM FILE USING READFILE API

'PARAMETERS:    FileName: Name Of File

'RETURNS:       Contents of file as a string

'EXAMPLE:       dim sAns as String
'               sAns = ReadStringFromFile("C:\MyFile.txt")
'**************************************************************
On Error GoTo ErrorHandler
Dim lHandle As Long
Dim lSuccess As Long
Dim lBytesRead As Long
Dim lBytesToRead As Long
Dim bytArr() As Byte
Dim sAns As String

lBytesToRead = FileLen(FileName)
ReDim bytArr(lBytesToRead) As Byte
'Get a handle to file
lHandle = CreateFile(FileName, GENERIC_WRITE Or GENERIC_READ, _
                     0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
If lHandle <> INVALID_HANDLE_VALUE Then
    'read file contents into a bytearray and convert to string
   lSuccess = ReadFile(lHandle, bytArr(0), _
                       lBytesToRead, lBytesRead, 0)
   
   sAns = ByteArrayToString(bytArr)
   ReadStringFromFile = sAns
   CloseHandle lHandle
   
End If
ErrorHandler:
End Function
Public Function ReadBytesFromFile(FileName As String) As Byte()
'***************************************************************
'PURPOSE:  READS BINARY DATA FROM FILE USING READFILE API

'PARAMETERS:    FileName: Name Of File

'RETURNS:       ByteArray Containing file data

'EXAMPLE:   Opens a Binary Document and Copies to a different file
           'dim bytBinaryDocument() as Byte
           'bytBinaryDocument = ReadBytesFromFile("C:\MyDoc.doc")
           'WriteBytesToFile "C:\MyNewDoc.doc", bytBinaryDocument
'**************************************************************
On Error GoTo ErrorHandler
Dim lHandle As Long
Dim lSuccess As Long
Dim lBytesRead As Long
Dim lBytesToRead As Long
Dim bytArr() As Byte


lBytesToRead = FileLen(FileName)
ReDim bytArr(lBytesToRead) As Byte
'Get a handle to file
lHandle = CreateFile(FileName, GENERIC_WRITE Or GENERIC_READ, _
                     0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
If lHandle <> INVALID_HANDLE_VALUE Then
    'read file contents into a bytearray and convert to string
   lSuccess = ReadFile(lHandle, bytArr(0), _
                       lBytesToRead, lBytesRead, 0)
   
  
   ReadBytesFromFile = bytArr
   CloseHandle lHandle
   
End If
ErrorHandler:
End Function

Public Function WriteBytesToFile(FileName As String, TheData() As Byte, _
  Optional NoOverwrite As Boolean = False) As Boolean
'***************************************************************
'PURPOSE:  WRITES BINARY DATA TO FILE USING WRITEFILE API

'PARAMETERS:    FileName: Name Of File
'               TheData: ByteArray Containing Binary Data
'               NoOverwrite (Optional): If set to true, will exit and
'                   return false if file exists.  This function does not
'                   work for appending to a file
              
'RETURNS:       True If Successful, false otherwise

'EXAMPLE: See Example for ReadBytesToFile
'*****************************************************
        

Dim lHandle As Long
Dim lSuccess As Long
Dim lBytesWritten As Long, lBytesToWrite As Long
If NoOverwrite = True And Dir(FileName) <> "" Then Exit Function
lBytesToWrite = UBound(TheData) - LBound(TheData)
         
lHandle = CreateFile(FileName, GENERIC_WRITE Or GENERIC_READ, _
                     0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)

'CreateFile returns INVALID_HANDLE_VALUE if it fails.
If lHandle <> INVALID_HANDLE_VALUE Then
   lSuccess = WriteFile(lHandle, TheData(0), _
                        lBytesToWrite, lBytesWritten, 0) <> 0
   If lSuccess <> 0 Then
      lSuccess = FlushFileBuffers(lHandle)
      lSuccess = CloseHandle(lHandle)
   End If
End If
ErrorHandler:
WriteBytesToFile = lSuccess <> 0
End Function
Private Function ByteArrayToString(bytArray() As Byte) As String
    Dim sAns As String
    Dim iPos As String
    
    sAns = StrConv(bytArray, vbUnicode)
    iPos = InStr(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)
    
    ByteArrayToString = sAns
 
 End Function


