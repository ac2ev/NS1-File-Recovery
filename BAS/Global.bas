Attribute VB_Name = "Global"
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (Prop As SHELLEXECUTEINFO) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public creg As cRegistry
Public IconPath As String
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_READONLY = &H1
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Type FindStruct
    StrWhat As String
    FindType As Integer
    LastFind As Integer
    Direction As Integer
    Search As Boolean
End Type
Public FindNext As FindStruct

Public Type StructL
    ID As Long
    msg As String * 100
End Type
Public LPK As StructL

Public PathApp As String
Public SelectedLPK As String
Public Const FileFilter As String = "NetStumbler File (*.ns1)|*.ns1|WifiFoFum (*.ns1)|*.ns1|Kismet (*.csv)|*.csv|Wi-Scan (*.wis)|*.wis|Summary (*.*)|*.*|OziExplorer (*.wpt)|*.wpt|Sniffi (*.txt)|*.txt"
Public Function GetTempDir() As String
   
   GetTempDir = String$(255, Chr$(0))
   GetTempDir = Left$(GetTempDir, GetTempPath(Len(GetTempDir), GetTempDir))

End Function
Sub Main()

    PathApp = App.Path
    SelectedLPK = GetSetting("HexEdit", "General", "Language", "English.lpk")
    If Right(PathApp, 1) <> "\" Then PathApp = PathApp & "\"
    frmEditor.Show
    
    
End Sub
Public Function GetMsg(ByVal nMsg As Integer) As String

    Get #3, nMsg, LPK
    GetMsg = Trim(LPK.msg)

End Function
Public Function ShowSave(ByVal hForm As Long, ByVal FileName As String, ByVal Filter As String, ByVal Title As String, ByVal InitDir As String) As String
 
    Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.lpstrTitle = Title
    ofn.hwndOwner = hForm
    ofn.hInstance = App.hInstance
    If Right(Filter, 1) <> "|" Then Filter = Filter & "|"
    For a = 1 To Len(Filter)
        If Mid(Filter, a, 1) = "|" Then Mid(Filter, a, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = FileName & Space(255 - Len(FileName))
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
    ShowSave = IIf(GetSaveFileName(ofn), Trim(ofn.lpstrFile), "")

End Function
Public Function ShowOpen(ByVal hForm As Long, ByVal Filter As String, ByVal Title As String, ByVal InitDir As String) As String
 
    Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hForm
    ofn.hInstance = App.hInstance
    If Right(Filter, 1) <> "|" Then Filter = Filter & "|"
    For a = 1 To Len(Filter)
        If Mid(Filter, a, 1) = "|" Then Mid(Filter, a, 1) = Chr(0)
    Next
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = Title
    ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    ShowOpen = IIf(GetOpenFileName(ofn), Trim(ofn.lpstrFile), "")

End Function
Public Sub CenterForm(frm)
        
    frm.Top = (Screen.Height / 2) - (frm.Height / 2)
    frm.Left = (Screen.Width / 2) - (frm.Width / 2)

End Sub
Public Function FileExists(ByVal PathName As String, Optional FileType As VBA.VbFileAttribute = vbNormal) As Boolean

    'Returns True if the passed pathname exi
    '     st
    'Otherwise returns False

    If PathName <> "" Then
        FileExists = (Dir$(PathName, FileType) <> "")
    End If
End Function

Public Function IsValidIP(Test As String) As Boolean
    Dim SubNets() As String
    Dim i As Integer
    

    If LCase(Test) = "localhost" Then
        IsValidIP = True
        Exit Function
    End If

    

    If Len(Test) > 16 Then
        IsValidIP = False
        Exit Function
    End If

    SubNets = Split(Test, ".")
    
    

    If UBound(SubNets) > 3 Or UBound(SubNets) <= 3 Then
        IsValidIP = False
        Exit Function
    End If

    

    For i = 0 To 3


        If Not IsNumeric(SubNets(i)) Or SubNets(i) < 0 Or SubNets(i) > 255 Then
            IsValidIP = False
            Exit Function
        End If

    Next

    
    IsValidIP = True
    Exit Function
End Function

Public Function IsValidMAC(Test As String) As Boolean
    Dim IDs() As String
    Dim i As Integer
    

    If Len(Test) > 17 Then
        IsValidMAC = False
        Exit Function
    End If

    IDs = Split(Test, ":")
    
    

    If UBound(IDs) > 5 Or UBound(IDs) <= 4 Then
        IsValidMAC = False
        Exit Function
    End If

   
    IsValidMAC = True
    Exit Function
End Function

