Attribute VB_Name = "modDialog"
Option Explicit
Public Const MAX_PATH As Long = 260
Public Type ShortItemId
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As ShortItemId
End Type

Public Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
'Public Const MAX_PATH = 260

Global StartMenu As String
Global Mode As String

' Declare API functions.

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, _
     ByVal lParam As Any) As Long

Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
   Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" _
   (ByVal pv As Long)

Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long


