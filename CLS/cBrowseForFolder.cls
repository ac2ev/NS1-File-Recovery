VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrowseForFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

Private Const MAX_PATH = 260
Private Const S_OK = 0           ' indicates success
Private Const S_FALSE = 1&   ' special HRESULT value
Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As Long ';// Return display name of item selected.
    lpszTitle As Long ';      // text to go in the banner over the tree.
    ulFlags As Long ';       // Flags that control the return stuff
    lpfn As Long
    lParam As Long         '// extra info that's passed back in callbacks
    iImage As Long ';      // output var: where to return the Image index.
End Type
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
   (lpBI As BROWSEINFO) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
   (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As Long) As Long
Private Declare Function SHGetDesktopFolder Lib "shell32.dll" _
   (ppshf As IVBShellFolder) As Long
Private Declare Function SHGetMalloc Lib "shell32.dll" (ppMalloc As IVBMalloc) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
   Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

'BROWSEINFO.ulFlags values:
Private Const BIF_RETURNONLYFSDIRS = &H1      'Only returns file system directories
Private Const BIF_DONTGOBELOWDOMAIN = &H2     'Does not include network folders below domain level
Private Const BIF_STATUSTEXT = &H4            'Includes status area in the dialog for use with callback
Private Const BIF_RETURNFSANCESTORS = &H8     'Only returns file system ancestors.
Private Const BIF_EDITBOX = &H10              'allows user to rename selection
Private Const BIF_VALIDATE = &H20             'insist on valid editbox result (or CANCEL)
Private Const BIF_USENEWUI = &H40             'Version 5.0. Use the new user-interface. Setting
                                             'this flag provides the user with a larger dialog box
                                             'that can be resized. It has several new capabilities
                                             'including: drag and drop capability within the
                                             'dialog box, reordering, context menus, new folders,
                                             'delete, and other context menu commands. To use
                                             'this flag, you must call OleInitialize or
                                             'CoInitialize before calling SHBrowseForFolder.
Private Const BIF_BROWSEFORCOMPUTER = &H1000  'Only returns computers.
Private Const BIF_BROWSEFORPRINTER = &H2000   'Only returns printers.
Private Const BIF_BROWSEINCLUDEFILES = &H4000 'Browse for everything

Private Const WM_USER = &H400
Private Const BFFM_SETSTATUSTEXTA = (WM_USER + 100)
Private Const BFFM_ENABLEOK = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW = (WM_USER + 103)
Private Const BFFM_SETSTATUSTEXTW = (WM_USER + 104)

Public Enum efbrCSIDLConstants
   CSIDL_DESKTOP = &H0                   '(desktop)
   CSIDL_INTERNET = &H1                  'Internet Explorer (icon on desktop)
   CSIDL_PROGRAMS = &H2                  'Start Menu\Programs
   CSIDL_CONTROLS = &H3                  'My Computer\Control Panel
   CSIDL_PRINTERS = &H4                  'My Computer\Printers
   CSIDL_PERSONAL = &H5                  'My Documents
   CSIDL_FAVORITES = &H6                 '(user name)\Favorites
   CSIDL_STARTUP = &H7                   'Start Menu\Programs\Startup
   CSIDL_RECENT = &H8                    '(user name)\Recent
   CSIDL_SENDTO = &H9                    '(user name)\SendTo
   CSIDL_BITBUCKET = &HA                 '(desktop)\Recycle Bin
   CSIDL_STARTMENU = &HB                 '(user name)\Start Menu
   CSIDL_DESKTOPDIRECTORY = &H10         '(user name)\Desktop
   CSIDL_DRIVES = &H11                   'My Computer
   CSIDL_NETWORK = &H12                  'Network Neighborhood
   CSIDL_NETHOOD = &H13                  '(user name)\nethood
   CSIDL_FONTS = &H14                    'windows\fonts
   CSIDL_TEMPLATES = &H15
   CSIDL_COMMON_STARTMENU = &H16         'All Users\Start Menu
   CSIDL_COMMON_PROGRAMS = &H17          'All Users\Programs
   CSIDL_COMMON_STARTUP = &H18           'All Users\Startup
   CSIDL_COMMON_DESKTOPDIRECTORY = &H19  'All Users\Desktop
   CSIDL_APPDATA = &H1A                  '(user name)\Application Data
   CSIDL_PRINTHOOD = &H1B                '(user name)\PrintHood
   CSIDL_LOCAL_APPDATA = &H1C            '(user name)\Local Settings\Applicaiton Data (non roaming)
   CSIDL_ALTSTARTUP = &H1D               'non localized startup
   CSIDL_COMMON_ALTSTARTUP = &H1E        'non localized common startup
   CSIDL_COMMON_FAVORITES = &H1F
   CSIDL_INTERNET_CACHE = &H20
   CSIDL_COOKIES = &H21
   CSIDL_HISTORY = &H22
   CSIDL_COMMON_APPDATA = &H23           'All Users\Application Data
   CSIDL_WINDOWS = &H24                  'GetWindowsDirectory()
   CSIDL_SYSTEM = &H25                   'GetSystemDirectory()
   CSIDL_PROGRAM_FILES = &H26            'C:\Program Files
   CSIDL_MYPICTURES = &H27               'C:\Program Files\My Pictures
   CSIDL_PROFILE = &H28                  'USERPROFILE
   CSIDL_PROGRAM_FILES_COMMON = &H2B     'C:\Program Files\Common
   CSIDL_COMMON_TEMPLATES = &H2D         'All Users\Templates
   CSIDL_COMMON_DOCUMENTS = &H2E         'All Users\Documents
   CSIDL_COMMON_ADMINTOOLS = &H2F        'All Users\Start Menu\Programs\Administrative Tools
   CSIDL_ADMINTOOLS = &H30               '(user name)\Start Menu\Programs\Administrative Tools

   CSIDL_FLAG_CREATE = &H8000            'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
   CSIDL_FLAG_DONT_VERIFY = &H4000       'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
   CSIDL_FLAG_MASK = &HFF00              'mask for all possible flag values
End Enum

Private m_hWndOwner As Long
Private m_sTitle As String
Private m_sInitialDir As String
Private m_sRootDir As String
Private m_bEditBox As Boolean
Private m_bStatusText As Boolean
Private m_bFileSystemOnly As Boolean
Private m_bValidateText As Boolean
Private m_bUseNewUI As Boolean
Private m_sDisplayName As String
Private m_pidlInitial As Long
Private m_bShown As Boolean
Private m_hWNdDialog As Long
Private m_bIncludeFiles As Boolean

Public Event Initialized()
Public Event SelectionChanged(ByVal sPath As String, ByRef bAllowOk As Boolean)
Attribute SelectionChanged.VB_Description = "Raised when the user changes the selection in the dialog box."
Public Event ValidationFailed(ByVal sPath As String, ByRef bKeepOpen As Boolean)
Attribute ValidationFailed.VB_Description = "Raised when the dialog edit box contains invalid text and OK is chosen.  Requires IE4 or higher."

Private alloc As IVBMalloc

Private Property Get Allocator() As IVBMalloc
    If alloc Is Nothing Then SHGetMalloc alloc
    Set Allocator = alloc
End Property

Public Property Get SpecialFolderLocation(ByVal eFolder As efbrCSIDLConstants) As String
Attribute SpecialFolderLocation.VB_Description = "Gets the location of a special folder.  Note that some special folders can only be retrieved for systems with IE4 or higher."
Dim pidl As Long
On Error Resume Next
   ' Get pidl of special folder:
   SHGetSpecialFolderLocation m_hWndOwner, eFolder, pidl
   If Err = 0 Then
      ' Convert it to a path:
      SpecialFolderLocation = PathFromPidl(pidl)
      ' Free the pidl:
      Allocator.free ByVal pidl
      pidl = 0
   End If
End Property
Private Function GetDesktopFolder() As IVBShellFolder
Dim lR As Long
    lR = SHGetDesktopFolder(GetDesktopFolder)
End Function
Public Property Get EditBox() As Boolean
Attribute EditBox.VB_Description = "Gets/sets whether the browse for folder dialog will show an edit box (only supported for IE4.0 and above)"
   EditBox = m_bEditBox
End Property
Public Property Let EditBox(ByVal bState As Boolean)
   m_bEditBox = bState
End Property
Public Property Get StatusText() As Boolean
Attribute StatusText.VB_Description = "Gets/sets whether the dialog box will have a line for status text."
   StatusText = m_bStatusText
End Property
Public Property Let StatusText(ByVal bState As Boolean)
   m_bStatusText = bState
End Property
Public Property Get FileSystemOnly() As Boolean
Attribute FileSystemOnly.VB_Description = "Gets/sets whether only file system objects can be picked in the dialog box."
   FileSystemOnly = m_bFileSystemOnly
End Property
Public Property Let FileSystemOnly(ByVal bState As Boolean)
   m_bFileSystemOnly = bState
End Property
Public Property Get ValidateEditBox() As Boolean
Attribute ValidateEditBox.VB_Description = "Gets/sets whether dialogs with an edit box allow the user to choose OK even when the edit box includes an invalid folder.  IE4 or higher required."
   ValidateEditBox = m_bValidateText
End Property
Public Property Let ValidateEditBox(ByVal bState As Boolean)
   m_bValidateText = bState
End Property
Public Property Get UseNewUI() As Boolean
Attribute UseNewUI.VB_Description = "Gets/sets whether the folder dialog box appears in the New style (larger, resizable, drag/drop support).  Only supported for Win2000."
   UseNewUI = m_bUseNewUI
End Property
Public Property Let UseNewUI(ByVal bState As Boolean)
   m_bUseNewUI = bState
End Property
Public Property Get Title() As String
Attribute Title.VB_Description = "Gets/sets the browse for folder dialog title."
   Title = m_sTitle
End Property
Public Property Let Title(ByVal sTitle As String)
   m_sTitle = sTitle
End Property
Public Property Get IncludeFiles() As Boolean
   IncludeFiles = m_bIncludeFiles
End Property
Public Property Let IncludeFiles(ByVal bState As Boolean)
   m_bIncludeFiles = bState
End Property
Public Property Get hwndOwner() As Long
Attribute hwndOwner.VB_Description = "Sets the window which owns the dialog box."
   hwndOwner = m_hWndOwner
End Property
Public Property Let hwndOwner(ByVal lhWnd As Long)
   m_hWndOwner = lhWnd
End Property
Public Property Get InitialDir() As String
Attribute InitialDir.VB_Description = "Gets/sets the initial directory which will be navigated to when the dialog is opened.  If the directory does not exist, there will be an error on opening the dialog."
   InitialDir = m_sInitialDir
End Property
Public Property Let InitialDir(ByVal sDir As String)
   m_sInitialDir = sDir
End Property
Public Property Get RootDir() As String
Attribute RootDir.VB_Description = "Gets/sets the top level path for the folder browsing.  Set to a blank string for the default."
   RootDir = m_sRootDir
End Property
Public Property Let RootDir(ByVal sDir As String)
   m_sRootDir = sDir
End Property

Private Function plGetOptions() As Long
Dim lOpt As Long
   If m_bEditBox Then
      lOpt = BIF_EDITBOX
   End If
   If m_bStatusText Then
      lOpt = lOpt Or BIF_STATUSTEXT
   End If
   If m_bFileSystemOnly Then
      lOpt = lOpt Or BIF_RETURNONLYFSDIRS Or BIF_RETURNFSANCESTORS
   End If
   If m_bValidateText Then
      lOpt = lOpt Or BIF_VALIDATE
   End If
   If m_bUseNewUI Then
      lOpt = lOpt Or BIF_USENEWUI
   End If
   If (m_bIncludeFiles) Then
      lOpt = lOpt Or BIF_BROWSEINCLUDEFILES
   End If
   plGetOptions = lOpt
   'Private Const BIF_BROWSEFORCOMPUTER = &H1000  'Only returns computers.
   'Private Const BIF_BROWSEFORPRINTER = &H2000   'Only returns printers.
   'Private Const BIF_BROWSEINCLUDEFILES = &H4000 'Browse for everything
End Function

Public Function BrowseForFolder() As String
Attribute BrowseForFolder.VB_Description = "Shows the browse for folder dialog."
Dim tBI As BROWSEINFO
Dim sOut As String
Dim sTitle As String
Dim pidlRoot As Long
Dim pidlInitial As Long
Dim pidlOut As Long
Dim sPath As String

   tBI.hwndOwner = m_hWndOwner
   sOut = String$(MAX_PATH, 0)
   tBI.pszDisplayName = StrPtr(sOut)
   sTitle = StrConv(m_sTitle, vbFromUnicode)
   tBI.lpszTitle = StrPtr(sTitle)
   tBI.ulFlags = plGetOptions()
   tBI.iImage = 0
    
   If Len(m_sRootDir) <> 0 Then
      ' Get a PIDL for the selected path:
      pidlRoot = PathToPidl(m_sRootDir)
   End If
   tBI.pidlRoot = pidlRoot
   If Len(m_sInitialDir) <> 0 Then
      m_pidlInitial = PathToPidl(m_sInitialDir)
   End If
   tBI.lParam = ObjPtr(Me)
   tBI.lpfn = plAddressOf(AddressOf BrowseCallbackProc)
   
   m_bShown = True
   pidlOut = SHBrowseForFolder(tBI)
   m_hWNdDialog = 0
   m_bShown = False
   m_sDisplayName = PointerToString(tBI.pszDisplayName)
   BrowseForFolder = PathFromPidl(pidlOut)
   
   ' Free the pidls we create
   If pidlRoot <> 0 Then
      Allocator.free ByVal pidlRoot
      pidlRoot = 0
   End If
   If m_pidlInitial <> 0 Then
      Allocator.free ByVal m_pidlInitial
   End If
   m_pidlInitial = 0
   
End Function
Friend Property Get pidlInitial() As Long
   pidlInitial = m_pidlInitial
End Property
Friend Sub SelectionChange(ByVal hWnd As Long, ByVal sPath As String, ByVal lParam As Long)
Dim bAllowOk As Boolean
   bAllowOk = True
   RaiseEvent SelectionChanged(sPath, bAllowOk)
   If Not bAllowOk Then
      SendMessageLong hWnd, BFFM_ENABLEOK, 0, 0
   End If
End Sub
Friend Function ValidateFailed(ByVal hWnd As Long, ByVal sPath As String) As Long
Dim bKeepOpen As Boolean
   RaiseEvent ValidationFailed(sPath, bKeepOpen)
   If bKeepOpen Then
      ValidateFailed = 1
   End If
End Function
Friend Function Initialized(ByVal hWnd As Long)
Attribute Initialized.VB_Description = "Raised when the dialog is about to be shown."
   m_hWNdDialog = hWnd
   RaiseEvent Initialized
End Function
Public Property Get DisplayName()
Attribute DisplayName.VB_Description = "Returns the display name of the selected item."
   DisplayName = m_sDisplayName
End Property
Public Sub SetFolder(ByVal sPath As String)
Attribute SetFolder.VB_Description = "Whilst the dialog box is open, sets the folder to the specified path."
Dim pidl As Long
   If m_bShown Then
      pidl = PathToPidl(sPath)
      SendMessageLong m_hWNdDialog, BFFM_SETSELECTIONA, 0, pidl
      Allocator.free pidl
      SetFocusAPI m_hWNdDialog
   Else
      pError 2
   End If
End Sub
Public Sub SetStatus(ByVal sText As String)
Attribute SetStatus.VB_Description = "Whilst the dialog box is open, sets the status text to the specified string (if StatusText is specified)."
Dim lR As Long
   If m_bShown Then
      lR = SendMessageStr(m_hWNdDialog, BFFM_SETSTATUSTEXTA, 0&, sText)
   Else
      pError 2
   End If
End Sub
Friend Property Get DialoghWnd() As Long
   If m_bShown Then
      DialoghWnd = m_hWNdDialog
   Else
      pError 2
   End If
End Property
Private Sub pError(ByVal lErr As Long)
   Err.Raise 26000 + lErr, App.EXEName & ".cBrowseForFolder", "Operation invalid unless dialog is displayed."
End Sub
Private Function plAddressOf(ByVal lPtr As Long) As Long
   plAddressOf = lPtr
End Function

Private Function PathToPidl(sPath As String) As Long
Dim folder As IVBShellFolder
Dim pidlMain As Long
Dim cParsed As Long
Dim afItem As Long
Dim lFilePos As Long
Dim lR As Long
Dim sRet As String

   ' Make sure the file name is fully qualified
   sRet = String$(MAX_PATH, 0)
   lR = GetFullPathName(sPath, MAX_PATH, sRet, lFilePos)
   If lR = 0 Then
      ApiRaise Err.LastDllError
   Else
      ' debug.Assert c <= cMaxPath
      sPath = Left$(sRet, lR)
   
      ' Convert the path name into a pointer to an item ID list (pidl)
      Set folder = GetDesktopFolder
      ' Will raise an error if path cannpt be found:
      folder.ParseDisplayName 0&, 0&, sPath, cParsed, pidlMain, afItem
      PathToPidl = pidlMain
   End If
    
End Function
Private Sub ApiRaise(ByVal e As Long)
   Err.Raise vbObjectError + 29000 + e, _
              App.EXEName & ".cBrowseForFolder", ApiError(e)
End Sub
Private Function ApiError(ByVal e As Long) As String
Dim s As String, c As Long
   s = String(256, 0)
   c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                     FORMAT_MESSAGE_IGNORE_INSERTS, _
                     0&, e, 0&, s, Len(s), ByVal 0&)
   If c Then ApiError = Left$(s, c)
End Function
Private Function PathFromPidl(ByVal pidl As Long) As String
Dim sPath As String
Dim lR As Long
   sPath = String$(MAX_PATH, 0)
   lR = SHGetPathFromIDList(pidl, sPath)
   If lR <> 0 Then
      PathFromPidl = Left$(sPath, lstrlen(sPath))
   End If
End Function



Private Sub Class_Initialize()
   m_sTitle = "Choose Folder"
End Sub

