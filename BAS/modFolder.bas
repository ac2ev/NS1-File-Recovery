Attribute VB_Name = "modFolder"
Option Explicit

Type tagOPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    strFilter As String
    strCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    strFile As String
    nMaxFile As Long
    strFileTitle As String
    nMaxFileTitle As Long
    strInitialDir As String
    strTitle As String
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    strDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Declare Function aht_apiGetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare Function aht_apiGetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Global Const ahtOFN_READONLY = &H1
Global Const ahtOFN_OVERWRITEPROMPT = &H2
Global Const ahtOFN_HIDEREADONLY = &H4
Global Const ahtOFN_NOCHANGEDIR = &H8
Global Const ahtOFN_SHOWHELP = &H10
Global Const ahtOFN_NOVALIDATE = &H100
Global Const ahtOFN_ALLOWMULTISELECT = &H200
Global Const ahtOFN_EXTENSIONDIFFERENT = &H400
Global Const ahtOFN_PATHMUSTEXIST = &H800
Global Const ahtOFN_FILEMUSTEXIST = &H1000
Global Const ahtOFN_CREATEPROMPT = &H2000
Global Const ahtOFN_SHAREAWARE = &H4000
Global Const ahtOFN_NOREADONLYRETURN = &H8000
Global Const ahtOFN_NOTESTFILECREATE = &H10000
Global Const ahtOFN_NONETWORKBUTTON = &H20000
Global Const ahtOFN_NOLONGNAMES = &H40000
Global Const ahtOFN_EXPLORER = &H80000
Global Const ahtOFN_NODEREFERENCELINKS = &H100000
Global Const ahtOFN_LONGNAMES = &H200000

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
Public Const SE_ERR_NOASSOC As Long = 31



Function GetOpenFile(Optional varDirectory As Variant, Optional varTitleForDialog As Variant, Optional vMode As Boolean, Optional vType As Boolean) As Variant
Dim strFilter As String
Dim lngFlags As Long
Dim varFileName As Variant

    lngFlags = ahtOFN_FILEMUSTEXIST Or ahtOFN_HIDEREADONLY Or ahtOFN_NOCHANGEDIR
    If IsMissing(varDirectory) Then
        varDirectory = ""
    End If
    If IsMissing(varTitleForDialog) Then
        varTitleForDialog = ""
    End If

    If vType = False Then
       strFilter = ahtAddFilterItem(strFilter, "All (*.*)", "*.*")
    Else
       strFilter = ahtAddFilterItem(strFilter, "HTML (*.html)", "*.html")
    End If
    ' Now actually call to get the file name.
    varFileName = ahtCommonFileOpenSave(OpenFile:=vMode, InitialDir:=varDirectory, Filter:=strFilter, FLAGS:=lngFlags, DialogTitle:=varTitleForDialog)
    If Not IsNull(varFileName) Then
        varFileName = TrimNull(varFileName)
    End If
    GetOpenFile = varFileName
End Function

Function ahtCommonFileOpenSave(Optional ByRef FLAGS As Variant, Optional ByVal InitialDir As Variant, Optional ByVal Filter As Variant, Optional ByVal FilterIndex As Variant, Optional ByVal DefaultExt As Variant, Optional ByVal FileName As Variant, Optional ByVal DialogTitle As Variant, Optional ByVal hwnd As Variant, Optional ByVal OpenFile As Variant) As Variant

Dim OFN As tagOPENFILENAME
Dim strFilename As String
Dim strFileTitle As String
Dim fResult As Boolean
    ' Give the dialog a caption title.
    If IsMissing(InitialDir) Then InitialDir = CurDir
    If IsMissing(Filter) Then Filter = ""
    If IsMissing(FilterIndex) Then FilterIndex = 1
    If IsMissing(FLAGS) Then FLAGS = 0&
    If IsMissing(DefaultExt) Then DefaultExt = ""
    If IsMissing(FileName) Then FileName = ""
    If IsMissing(DialogTitle) Then DialogTitle = ""
    If IsMissing(hwnd) Then hwnd = frmLiveUpdate.hwnd
    If IsMissing(OpenFile) Then OpenFile = True
    ' Allocate string space for the returned strings.
    strFilename = Left(FileName & String(256, 0), 256)
    strFileTitle = String(256, 0)
    ' Set up the data structure before you call the function
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hwnd
        .strFilter = Filter
        .nFilterIndex = FilterIndex
        .strFile = strFilename
        .nMaxFile = Len(strFilename)
        .strFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle)
        .strTitle = DialogTitle
        .FLAGS = FLAGS
        .strDefExt = DefaultExt
        .strInitialDir = InitialDir
        ' Didn't think most people would want to deal with
        ' these options.
        .hInstance = 0
        .strCustomFilter = ""
        .nMaxCustFilter = 0
        .lpfnHook = 0
        'New for NT 4.0
        .strCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
    End With
    ' This will pass the desired data structure to the
    ' Windows API, which will in turn it uses to display
    ' the Open/Save As Dialog.
    If OpenFile Then
        fResult = aht_apiGetOpenFileName(OFN)
    Else
        fResult = aht_apiGetSaveFileName(OFN)
    End If
    
    If fResult Then
        
        If Not IsMissing(FLAGS) Then FLAGS = OFN.FLAGS
        ahtCommonFileOpenSave = TrimNull(OFN.strFile)
    Else
        ahtCommonFileOpenSave = vbNullString
    End If
End Function
Function ahtAddFilterItem(strFilter As String, _
    strDescription As String, Optional varItem As Variant) As String
    If IsMissing(varItem) Then varItem = "*.*"
    ahtAddFilterItem = strFilter & strDescription & vbNullChar & varItem & vbNullChar
End Function
Private Function TrimNull(ByVal strItem As String) As String
Dim intPos As Integer
    intPos = InStr(strItem, vbNullChar)
    If intPos > 0 Then
        TrimNull = Left(strItem, intPos - 1)
    Else
        TrimNull = strItem
    End If
End Function

Public Function RetWavFiles(strPath) As String
Dim strFilter As String
Dim lngFlags As Long
Dim strDBFile As String
Dim Lechemin As String
  strFilter = ahtAddFilterItem(strFilter, "Fichier Son (*.Wav)", "*.WAV")
  
  Lechemin = ahtCommonFileOpenSave(InitialDir:=strPath, FileName:=strDBFile, Filter:=strFilter, FilterIndex:=3, FLAGS:=lngFlags, DialogTitle:="Ouvrir un fichier Wav")
  
  If Not IsNull(Lechemin) Then
    RetWavFiles = Lechemin
  End If
End Function

Public Function RetOccurence(SearchString As String, Phrase As String) As Integer
'Retourne le nombre d'occurence d'une string dans une chaîne
'By christ

Dim i As Integer
Dim ctr As Integer

If IsNull(SearchString) Then RetOccurence = 0: Exit Function
If IsNull(Phrase) Then RetOccurence = 0: Exit Function
ctr = 0


  For i = 1 To Len(Phrase)
      If Mid(Phrase, i, Len(SearchString)) = SearchString Then
         'Debug.Print Mid(Phrase, i, Len(SearchString))
         ctr = ctr + 1
      End If
  Next
  RetOccurence = ctr

End Function







