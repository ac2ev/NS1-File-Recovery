VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class contains functions to control and manage a computers (local or remote) registry."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'**************************************************************************************
'  Title    :  Registry Manipulation Class
'  Author   :  Todd Herman/Don Kiser
'  Date     :  09/02
'  Purpose  :  I have found several Registry modules and even registry classes.
'              I looked through these and used them at times but found them lacking.
'              They never did all the things I needed.  That's where this class comes
'              in.  The purpose of this class is to try and put as much registry
'              control into it as possible.  One key feature this class will offer
'              that none other do, is the ability to LOAD/UNLOAD registry hives and
'              perform basic registry tasks on these hives.  A helpfile should have
'              been included with this class.  This file fully documents this class.
'**************************************************************************************

'Error Constants
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_INVALID_FUNCTION = 1
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_INVALID_HANDLE = 6&
Private Const ERROR_INVALID_ACCESS = 12&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_NETNAME_DELETED = 64&
Private Const ERROR_NETWORK_ACCESS_DENIED = 65&
Private Const ERROR_TOO_MANY_SESS = 69&
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_BAD_PATHNAME = 161&
Private Const ERROR_LOCK_FAILED = 167&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_REGISTRY_RECOVERED = 1014&
Private Const ERROR_REGISTRY_CORRUPT = 1015&
Private Const ERROR_REGISTRY_IO_FAILED = 1016&
Private Const ERROR_NOT_REGISTRY_FILE = 1017&
Private Const ERROR_KEY_DELETED = 1018&
Private Const ERROR_NO_LOG_SPACE = 1019&
Private Const ERROR_KEY_HAS_CHILDREN = 1020&
Private Const ERROR_NOT_ALL_ASSIGNED As Long = 1300&
Private Const ERROR_NO_TOKEN = 1008&
'Registry Access Constants
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

'Token and Privilege Constants
Private Const TokenPrivileges As Long = 3
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const SE_PRIVILEGE_DISABLED As Long = 0
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"
Private Const SE_RESTORE_NAME = "SeRestorePrivilege"
Private Const TOKEN_ASSIGN_PRIMARY = &H1
Private Const TOKEN_DUPLICATE = &H2
Private Const TOKEN_IMPERSONATE = &H4
Private Const TOKEN_QUERY = &H8
Private Const TOKEN_QUERY_SOURCE = &H10
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_ADJUST_GROUPS = &H40
Private Const TOKEN_ADJUST_DEFAULT = &H80
Private Const TOKEN_ADJUST_SESSIONID = &H100
Private Const TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE = &H10 Or TOKEN_ADJUST_PRIVILEGES = &H20 Or TOKEN_ADJUST_GROUPS = &H40 Or TOKEN_ADJUST_DEFAULT = &H80 Or TOKEN_ADJUST_SESSIONID = &H100)
Private Const TOKEN_READ = (STANDARD_RIGHTS_READ Or TOKEN_QUERY)
Private Const TOKEN_WRITE = (STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
Private Const TOKEN_EXECUTE = (STANDARD_RIGHTS_WRITE)
Private Const TOKEN_MY_ACCESS = (TOKEN_READ Or TOKEN_WRITE)

'Miscellaneous Constants
Private Const REG_OPTION_VOLATILE = 1
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const LANG_NEUTRAL = &H0
Private Const SUBLANG_DEFAULT = &H1
Private Const ANYSIZE_ARRAY = 20

'Required Registry Types
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

'Security Privilege Types and Structures
Private Type luid
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As luid
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

'Enumerations
Public Enum DataType
    REG_SZ = &H1
    REG_EXPAND_SZ = &H2
    REG_BINARY = &H3
    REG_DWORD = &H4
    REG_MULTI_SZ = &H7
End Enum

Public Enum HKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

'Registry API Declarations
Private Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Private Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long

'Token and Process API Declarations
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As luid) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Long, ReturnLength As Long) As Long
Private Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long

'Miscellaneous API Declarations
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)

'Local Variables
Private mvarhkey As Long                'The handle of the KEY being accessed
Private mvarKeyPath As String           'The path to a specific key
Private mvarLastError As Long           'Error number of the last error
Private mvarErrorDescription As String  'Description of the last error
Private mvarErrorAction As String       'What action was being done when error occurred
Public Property Let KeyPath(ByVal sData As String)
Attribute KeyPath.VB_Description = "The root path.  This should be the parent path for the value stored in the SubKey property."
    mvarKeyPath = sData
End Property
Public Property Get KeyPath() As String
    KeyPath = mvarKeyPath
End Property
Public Property Let hKey(ByVal lData As HKeys)
Attribute hKey.VB_Description = "Contains the handle to the registry key."
    mvarhkey = lData
End Property
Public Property Get hKey() As HKeys
    hKey = mvarhkey
End Property
Public Property Get Error() As Long
    'Read only property containing the number of the last error that occurred
    Error = mvarLastError
End Property
Public Property Get ErrorDescription() As String
    'Read only property that stores the description of the last error
    ErrorDescription = mvarErrorDescription
End Property
Public Property Get ErrorAction() As String
    'Read only property that stores what action was being performed when
    'an error occurred
    ErrorAction = mvarErrorAction
End Property
Private Sub Class_Initialize()

    'Class initialization routine
    mvarLastError = 0
    mvarErrorDescription = ""
    mvarErrorAction = ""

End Sub
Public Function CreateKey(Optional ByVal lHkey As Variant, Optional sKeyPath As Variant) As Boolean

    'The name basically gives away what this function does.  It creates a new key.
    'If the key is already present, it opens it.. although this function will simply
    'close it immediately
    Dim lHandle As Long
    Dim lResult As Long
    Dim lDisposition As Long
    Dim udtSecurity As SECURITY_ATTRIBUTES
    
    'Clear any error values
    Call ClearError
    
    If IsMissing(lHkey) Then
        lHkey = mvarhkey
    Else
        lHkey = CLng(lHkey)
    End If
    
    If IsMissing(sKeyPath) Then
        sKeyPath = mvarKeyPath
    Else
        sKeyPath = CStr(sKeyPath)
    End If
    
    'Make sure the keypath passed does not begin with the "\" character becuase this
    'might cause problems
    If Left(sKeyPath, 1) = "\" Then sKeyPath = Mid(sKeyPath, 2)
    
    'Try and create the key
    lResult = RegCreateKeyEx(lHkey, sKeyPath, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, udtSecurity, lHandle, lDisposition)
    If lResult = ERROR_SUCCESS Then
        CreateKey = True
    Else
        'Log any errors
        mvarLastError = lResult
        mvarErrorDescription = GetErrorMessage(lResult)
        mvarErrorAction = "Attempting to create the key '" & sKeyPath & "'"
        CreateKey = False
    End If
    
    'Close the handle to the new or opened key
    If lHandle Then RegCloseKey lHandle
    
End Function
Public Function DeleteKey(Optional ByVal lHkey As Variant, Optional sKeyPath As Variant, Optional bRecursive As Variant) As Boolean

    'Again, the name sort of gives this away.  This function will delete a key.
    'This can be done on a single key or recursively.
    Dim lResult As Long
    
    'Clear any error values
    Call ClearError
    
    If IsMissing(lHkey) Then
        lHkey = mvarhkey
    Else
        lHkey = CLng(lHkey)
    End If
    
    If IsMissing(bRecursive) Then
        bRecursive = CBool(False)
    Else
        bRecursive = CBool(bRecursive)
    End If
    
    If IsMissing(sKeyPath) Then
        sKeyPath = mvarKeyPath
    Else
        sKeyPath = CStr(sKeyPath)
    End If
    
    'Make sure the keypath passed does not begin with the "\" character becuase this
    'might cause problems
    If Left(sKeyPath, 1) = "\" Then sKeyPath = Mid(sKeyPath, 2)
    
    'Check if recursion was selected
    If Not bRecursive Then
        'Try and delete the key
        lResult = RegDeleteKey(lHkey, sKeyPath)
        
        If lResult = ERROR_SUCCESS Then
            DeleteKey = True
        Else
            'Log any errors that occurred
            mvarLastError = lResult
            mvarErrorDescription = GetErrorMessage(lResult)
            mvarErrorAction = "Attempting to delete the key '" & sKeyPath & "'"
            DeleteKey = False
        End If
    Else
        'On Windows NT, you must delete all subkeys before you can delete the
        'parent key.  This option will recursively delete all subkeys.  This option
        'is not needed under Windows 95/98/ME.  Deleting the top most key will
        'automatically delete all subkeys.
        Call DeleteKeyRecursive(lHkey, sKeyPath)
        
        'Check if there was an error
        If mvarLastError = ERROR_SUCCESS Then
            DeleteKey = True
        Else
            DeleteKey = False
        End If
    End If
    
End Function
Private Sub DeleteKeyRecursive(ByVal lHkey As Long, ByVal sKeyPath As String)

    Dim lResult As Long
    Dim sCurrentKeyPath As String
    Dim colSubKeys As Collection
    Dim saKeys() As String
    Dim i As Long
    
    'Now we enum the keys then loop through that collection of keys
    'and keep calling delete keys recursively for each item found
    saKeys() = EnumRegistryKeys(lHkey, sKeyPath)
    If Not saKeys(0) = "No Data" Then
        For i = 0 To UBound(saKeys)
            'Loop through the subkeys collection
            Call DeleteKeyRecursive(lHkey, KeyPath & "\" & saKeys(i))
        Next i
    End If
    
    If mvarLastError = ERROR_SUCCESS Then
        'Delete this key
        lResult = RegDeleteKey(lHkey, sKeyPath)
        
        'Log any errors that occurred
        If Not lResult = ERROR_SUCCESS Then
            mvarLastError = lResult
            mvarErrorDescription = GetErrorMessage(lResult)
            mvarErrorAction = "Attempting to delete the key '" & sKeyPath & "'"
        End If
    End If
    
End Sub
Public Function DeleteValue(sValueName As String, Optional ByVal lHkey As Variant, Optional sKeyPath As Variant) As Boolean

    'This function will delete the specified value.
    Dim lResult As Long
    Dim lHandle As Long
    
    'Clear any error values
    Call ClearError
    
    If IsMissing(lHkey) Then
        lHkey = mvarhkey
    Else
        lHkey = CLng(lHkey)
    End If
    
    If IsMissing(sKeyPath) Then
        sKeyPath = mvarKeyPath
    Else
        sKeyPath = CStr(sKeyPath)
    End If
    
    'Make sure the keypath passed does not begin with the "\" character becuase this
    'might cause problems
    If Left(sKeyPath, 1) = "\" Then sKeyPath = Mid(sKeyPath, 2)
    
    'First, open the key that contains the value
    lResult = RegOpenKeyEx(lHkey, sKeyPath, 0, KEY_ALL_ACCESS, lHandle)
    If lResult = ERROR_SUCCESS Then
        'Try and delete the specified value
        lResult = RegDeleteValue(lHandle, sValueName)
        If lResult = ERROR_SUCCESS Then
            RegCloseKey lHandle
            DeleteValue = True
        Else
            'Log any errors that occurred
            mvarLastError = lResult
            mvarErrorDescription = GetErrorMessage(lResult)
            mvarErrorAction = "Attempting to delete the value '" & sValueName & "' from '" & sKeyPath & "'"
            DeleteValue = False
        End If
    Else
        'Log any errors that occurred
        mvarLastError = lResult
        mvarErrorDescription = GetErrorMessage(lResult)
        mvarErrorAction = "Attempting to delete the value '" & sValueName & "' from '" & sKeyPath & "'"
        DeleteValue = False
    End If
    
End Function
Public Function EnumRegistryKeys(Optional ByVal lHkey As Variant, Optional ByVal sKeyPath As Variant) As Variant

    'Enumerate all keys under a given key.  The information is returned as an array.
    Dim lResult As Long
    Dim lHandle As Long
    Dim lIndex As Long
    Dim lLength As Long
    Dim sSubKeyName As String
    Dim udtFileTime As FILETIME
    Dim saSubKeys() As String
    
    If IsMissing(lHkey) Then
        lHkey = mvarhkey
    Else
        lHkey = CLng(lHkey)
    End If
    
    If IsMissing(sKeyPath) Then
        sKeyPath = mvarKeyPath
    Else
        sKeyPath = CStr(sKeyPath)
    End If
    
    'Make sure the keypath passed does not begin with the "\" character becuase this
    'might cause problems
    If Left(sKeyPath, 1) = "\" Then sKeyPath = Mid(sKeyPath, 2)
    
    If Len(sKeyPath) Then
        'Attempt to open the key
        lResult = RegOpenKeyEx(lHkey, sKeyPath, 0, KEY_READ, lHandle)
        If lResult = ERROR_SUCCESS Then
            Do While lResult = ERROR_SUCCESS
                'The key was successfully opened, now we need to enumerate the sub keys
                lLength = 255
                sSubKeyName = Space(lLength)
                
                'Call the enumeration function for each item and store the item
                'in an array
                lResult = RegEnumKeyEx(lHandle, lIndex, sSubKeyName, lLength, 0, "", vbNull, udtFileTime)
                If lResult = ERROR_SUCCESS Or lResult = ERROR_MORE_DATA Then
                    ReDim Preserve saSubKeys(lIndex)
                    
                    saSubKeys(lIndex) = TrimNull(sSubKeyName)
                Else
                    Exit Do
                End If
                lResult = 0
                lIndex = lIndex + 1
            Loop
        Else
            'An error occurred
            mvarLastError = lResult
            mvarErrorDescription = GetErrorMessage(lResult)
            mvarErrorAction = "Attempting to enumerate all subkeys under '" & sKeyPath & "'"
        End If
    Else
        'Log any errors
        mvarLastError = ERROR_BAD_PATHNAME
        mvarErrorDescription = "No key path was supplied"
        mvarErrorAction = "Attempting to enumerate all subkeys"
    End If
    
    'We don't want to return an empty array so make sure at least one array
    'element exists and set it to 'No Data'
    If Not (mvarLastError = ERROR_SUCCESS Or mvarLastError = ERROR_NO_MORE_ITEMS) Or lIndex = 0 Then
        ReDim saSubKeys(0)
        
        saSubKeys(0) = "No Data"
    End If
    
    EnumRegistryKeys = saSubKeys
    
    'Close the registry handle
    If lHandle Then RegCloseKey lHandle
    
End Function
Public Function EnumRegistryValues(Optional ByVal lHkey As Variant, Optional ByVal sKeyPath As Variant) As Variant

    'Enumerate all keys under a given key
    Dim lResult As Long
    Dim lHandle As Long
    Dim lIndex As Long
    Dim lNameLength As Long
    Dim Length As Long
    Dim sValueName As String
    Dim sValue As String
    Dim yBuffer() As Byte
    Dim lValueType As Long
    Dim saValueNames() As String
    Dim saValueData() As String
    Dim saValues() As String
    Dim lTemp As Long
    Dim sTemp As String
    Dim i As Long
    
    'Clear any error values
    Call ClearError
    
    If IsMissing(lHkey) Then
        lHkey = mvarhkey
    Else
        lHkey = CLng(lHkey)
    End If
    
    If IsMissing(sKeyPath) Then
        sKeyPath = mvarKeyPath
    Else
        sKeyPath = CStr(sKeyPath)
    End If
    
    'Make sure the keypath passed does not begin with the "\" character becuase this
    'might cause problems
    If Left(sKeyPath, 1) = "\" Then sKeyPath = Mid(sKeyPath, 2)
    
    If Len(sKeyPath) Then
        'Attempt to open the key
        lResult = RegOpenKeyEx(lHkey, sKeyPath, 0, KEY_READ, lHandle)
        If lResult = ERROR_SUCCESS Then
            'The key was successfully opened, now we need to enumerate the values
            Do While lResult = ERROR_SUCCESS
                Length = 1
                lNameLength = 255
                sValueName = Space(lNameLength)
                ReDim yBuffer(0 To Length - 1)
                
                'Attempt to enumerate the first value
                lResult = RegEnumValue(lHandle, lIndex, sValueName, lNameLength, ByVal 0&, lValueType, yBuffer(0), Length)
                
                'Check if enough buffer space was allocated
                If lResult = ERROR_MORE_DATA Or lResult = ERROR_SUCCESS Then
                    If lResult = ERROR_MORE_DATA Then
                        ReDim yBuffer(0 To Length - 1)
                        
                        'We needed to allocate more buffer space.  We did, so now
                        'try the function again
                        lResult = RegEnumValue(hKey, lIndex, sValueName, lNameLength, ByVal 0&, lValueType, yBuffer(0), Length)
                        If Not lResult = ERROR_SUCCESS Then
                            If lResult = ERROR_NO_MORE_ITEMS Then
                                Exit Do
                            Else
                                'We ran into some problem, flag the error properties
                                mvarLastError = lResult
                                mvarErrorDescription = GetErrorMessage(lResult)
                                mvarErrorAction = "Attempting to enumerate all values under '" & sKeyPath & "'"
                                Exit Do
                            End If
                        End If
                    End If
                    
                    'If we made it here, everything is all good
                    ReDim Preserve saValueNames(lIndex)
                    ReDim Preserve saValueData(lIndex)
                    
                    'Initialize our storage arrays.  We have one array for all the names
                    'and one array for all the values.  At the end, both arrays will be
                    'combined into one two-dimensional array which will be returned.
                    If sValueName = "" And Length > 0 Then
                         saValueNames(lIndex) = "[Default]"
                    Else
                        saValueNames(lIndex) = TrimNull(sValueName)
                    End If
                    saValueData(lIndex) = Space(Length - 1)
                    
                    'Analyze the data type of the current value
                    Select Case lValueType
                        Case REG_DWORD
                            'We use the CopyMemory api to quickly copy the contents
                            'of our byte array to a long.  We then format it to the
                            'way we want it
                            CopyMemory lTemp, yBuffer(0), Length
                            saValueData(lIndex) = "REG_DWORD : " & IIf(lTemp = 0, CStr(lTemp), "0x" & CStr(lTemp))
                        Case REG_BINARY
                            'We need to loop through our byte array and perform some
                            'converting to get the string to appear as we want
                            For i = 0 To Length - 1
                                If i = 0 Then saValueData(lIndex) = "REG_BINARY : "
                                saValueData(lIndex) = saValueData(lIndex) & " " & Format(Trim(Hex(yBuffer(i))), "0#")
                            Next i
                        Case REG_SZ
                            'Use CopyMemory to copy the string information to our array
                            If Length = 0 Then
                                saValueData(lIndex) = "REG_SZ : (No Value)"
                            Else
                                CopyMemory ByVal saValueData(lIndex), yBuffer(0), Length - 1
                                saValueData(lIndex) = "REG_SZ : " & saValueData(lIndex)
                            End If
                        Case REG_EXPAND_SZ
                            'Use CopyMemory to copy the string information to our array.
                            'This data type may contain EnvironmentStrings, such as
                            '%PATH%.  We use the ExpandEnvironmentStrings function to
                            'return the full string.
                            If Length = 0 Then
                                saValueData(lIndex) = "REG_EXPAND_SZ : (No Value)"
                            Else
                                CopyMemory ByVal saValueData(lIndex), yBuffer(0), Length - 1
                                
                                'If the buffer is not large enough, the needed size is returned
                                Length = ExpandEnvironmentStrings(saValueData(lIndex), sTemp, Len(sTemp))
                                If sTemp = "" Then
                                    sTemp = Space(Length)
                                    Length = ExpandEnvironmentStrings(saValueData(lIndex), sTemp, Length)
                                End If
                                
                                saValueData(lIndex) = "REG_EXPAND_SZ : " & TrimNull(sTemp)
                            End If
                        Case REG_MULTI_SZ
                            'Use CopyMemory to copy the string information to our array.
                            'This information is returned as a vbNullChar delimited string.
                            If Length = 0 Then
                                saValueData(lIndex) = "REG_MULTI_SZ : (No Value)"
                            Else
                                CopyMemory ByVal saValueData(lIndex), yBuffer(0), Length - 2
                                saValueData(lIndex) = "REG_MULTI_SZ : " & saValueData(lIndex)
                            End If
                        Case Else
                            'Unsupported type
                            saValueData(lIndex) = "DATA TYPE UNSUPPORTED"
                    End Select
                Else
                    'An error occurred
                    mvarLastError = lResult
                    mvarErrorDescription = GetErrorMessage(lResult)
                    mvarErrorAction = "Attempting to enumerate all values under '" & sKeyPath & "'"
                    Exit Do
                End If
                lIndex = lIndex + 1
            Loop
        Else
            'An error occurred
            mvarLastError = lResult
            mvarErrorDescription = GetErrorMessage(lResult)
            mvarErrorAction = "Attempting to enumerate all values under '" & sKeyPath & "'"
        End If
    Else
        'An error occurred
        mvarLastError = ERROR_BAD_PATHNAME
        mvarErrorDescription = "No key path was supplied"
        mvarErrorAction = "Attempting to enumerate all values"
    End If
    
    'If no errors occurred, combine the two temporary arrays into 1 2-dimension array
    If (mvarLastError = ERROR_SUCCESS Or mvarLastError = ERROR_NO_MORE_ITEMS) And lIndex > 0 Then
        ReDim saValues(lIndex, 0 To 1)
        
        For lIndex = 0 To UBound(saValueNames)
            saValues(lIndex, 0) = saValueNames(lIndex)
            saValues(lIndex, 1) = saValueData(lIndex)
        Next
    Else
        'We don't want to return an empty array, so store something in it
        ReDim saValues(1, 0 To 1) As String
        saValues(0, 0) = "No Data"
        saValues(0, 1) = "No Data"
    End If
    
    EnumRegistryValues = saValues
        
    'Close the registry handle
    If lHandle Then RegCloseKey lHandle
    
End Function
Public Function GetRegistryValue(ByVal sValueName As String, Optional ByVal vDefaultValue As Variant, Optional lHkey As Variant, Optional ByVal sKeyPath As Variant, Optional ByRef ValueType As Long, Optional AppendPathToCurrentKeyPath As Boolean = True) As Variant
    Dim lResult As Long
    Dim lHandle As Long
    Dim Length As Long
    Dim sValue As String
    Dim yBuffer() As Byte
    Dim lValueType As Long
    Dim lTemp As Long
    Dim sTemp As String
    Dim i As Long
    Dim TestString As String
    'Clear any error values
    Call ClearError

    If IsMissing(vDefaultValue) Or IsEmpty(vDefaultValue) Then
        GetRegistryValue = Empty
    Else
        GetRegistryValue = vDefaultValue 'May get overwritten if value is found
    End If
    
    If IsMissing(lHkey) Then
        lHkey = mvarhkey
    Else
        lHkey = CLng(lHkey)
    End If
    
    If IsMissing(sKeyPath) Then
        sKeyPath = mvarKeyPath
    Else
        sKeyPath = CStr(sKeyPath)
    End If
    
    If AppendPathToCurrentKeyPath Then
        sKeyPath = mvarKeyPath & sKeyPath
    End If
    'Make sure the keypath passed does not begin with the "\" character becuase this
    'might cause problems
    If Left(sKeyPath, 1) = "\" Then sKeyPath = Mid(sKeyPath, 2)
    
    'Attempt to open the key
    If Len(sKeyPath) Then
        lResult = RegOpenKeyEx(lHkey, sKeyPath, 0, KEY_READ, lHandle)
        If lResult = ERROR_SUCCESS Then
            Length = 1024
            ReDim yBuffer(0 To Length - 1)
            
            'Query the value
            lResult = RegQueryValueEx(lHandle, sValueName, ByVal 0&, lValueType, yBuffer(0), Length)
            If lResult = ERROR_MORE_DATA Or lResult = ERROR_SUCCESS Then
                'We need to check if our data buffer was large enough
                If lResult = ERROR_MORE_DATA Then
                    ReDim yBuffer(0 To Length - 1)
                    'The data buffer was not large enough so we resized it and
                    'need to query the value again
                    lResult = RegQueryValueEx(lHandle, sValueName, ByVal 0&, lValueType, yBuffer(0), Length)
                    If Not lResult = ERROR_SUCCESS Then
                        If Not lResult = ERROR_NO_MORE_ITEMS Then
                            'We ran into some problem, flag the error properties
                            mvarLastError = lResult
                            mvarErrorDescription = GetErrorMessage(lResult)
                            mvarErrorAction = "Attempting to get the value '" & sValueName & "' from '" & sKeyPath & "'"
                            Exit Function
                        End If
                    End If
                End If
                
                'The following case statement will format the data we just got,
                'according to the data type
                Select Case lValueType
                    Case REG_DWORD
                        'We use the CopyMemory api to quickly copy the contents
                        'of our byte array to a long.  We then format it to the
                        'way we want it
                        CopyMemory lTemp, yBuffer(0), Length
                        GetRegistryValue = lTemp
                    Case REG_BINARY
                        'Resize resulting array to fit
                        If Length <> UBound(yBuffer) + 1 Then
                            ReDim Preserve yBuffer(0 To Length - 1) As Byte
                        End If
                        GetRegistryValue = yBuffer()
                    Case REG_SZ
                        ' copy everything but the trailing null char
                        ' Ignore Blank Strings
                        If Length <> 0 Then
                            sValue = Space$(Length - 1)
                            CopyMemory ByVal sValue, yBuffer(0), Length - 1
                            GetRegistryValue = sValue
                        End If
                    Case REG_EXPAND_SZ
                        'Use CopyMemory to copy the string information to our array.
                        'This data type may contain EnvironmentStrings, such as
                        '%PATH%.  We use the ExpandEnvironmentStrings function to
                        'return the full string.
                        ' Ignore a Blank String
                        If Length <> 0 Then
                           sValue = Space$(Length - 1)
                           CopyMemory ByVal sValue, yBuffer(0), Length - 1
                           Length = ExpandEnvironmentStrings(sValue, sValue, Len(sValue))
                           GetRegistryValue = Left$(sValue, Length)
                        End If
                    Case REG_MULTI_SZ
                        'Use CopyMemory to copy the string information to our array.
                        'This information is returned as a vbNullChar delimited string.
                        sValue = Space$(Length - 2)
                        CopyMemory ByVal sValue, yBuffer(0), Length - 2
                        'A nonexistant value for REG_MULTI_SZ will return a string of nulls
                        'with a length = 1022
                        'This is because at the beginging of the routine we define Length = 1024
                        ' resString = Space$(length -2) = 1022
                        'So If we trims all nulls and are left with an empty string then
                        'the value doesn't exist so the defualt value is returned
                        'Set resstring to a temporary variable because trimnull will truncate it
                        TestString = sValue
                        If Len(TrimNull(sValue)) > 0 Then GetRegistryValue = sValue
                    Case Else
                        'Unsupported type
                        mvarLastError = ERROR_CANTREAD
                        mvarErrorDescription = "Unsupported Type"
                        mvarErrorAction = "Attempting to get the value '" & sValueName & "'"
                End Select
            Else
                'An error occurred
                mvarLastError = lResult
                mvarErrorDescription = GetErrorMessage(lResult)
                mvarErrorAction = "Attempting to get the value '" & sValueName & "' from '" & sKeyPath & "'"
            End If
        Else
            'An error occurred
            mvarLastError = lResult
            mvarErrorDescription = GetErrorMessage(lResult)
            mvarErrorAction = "Attempting to get the value '" & sValueName & "' from '" & sKeyPath & "'"
        End If
    Else
        mvarLastError = ERROR_BAD_PATHNAME
        mvarErrorDescription = "No key path was supplied"
        mvarErrorAction = "Attempting to get the value '" & sValueName & "'"
    End If
        
    If lHandle Then RegCloseKey lHandle

End Function
Public Function SetRegistryValue(ByVal sValueName As String, ByVal vValue As Variant, lDataType As DataType, Optional lHkey As Variant, Optional ByVal sKeyPath As Variant, Optional CreateKeyifNotExist As Boolean = True, Optional AppendPathToCurrentKeyPath As Boolean = True) As Boolean
    
    Dim lResult As Long
    Dim lHandle As Long
    Dim sValue As String
    Dim yBuffer() As Byte
    Dim lTemp As Long
    Dim saTemp() As String
    Dim binValue() As Byte
    Dim sTemp As String
    Dim Length As Integer
    Dim i As Long

    Call ClearError

    If IsMissing(lHkey) Then
        lHkey = mvarhkey
    Else
        lHkey = CLng(lHkey)
    End If
    
    If IsMissing(sKeyPath) Then
        sKeyPath = mvarKeyPath
    Else
        sKeyPath = CStr(sKeyPath)
    End If
    
    If AppendPathToCurrentKeyPath Then
        sKeyPath = mvarKeyPath & sKeyPath
    End If
    'Make sure the keypath passed does not begin with the "\" character becuase this
    'might cause problems
    If Left(sKeyPath, 1) = "\" Then sKeyPath = Mid(sKeyPath, 2)
    
    SetRegistryValue = False
KeyCreated:
    'Attempt to open the key
    If Len(sKeyPath) Then
        lResult = RegOpenKeyEx(lHkey, sKeyPath, 0, KEY_WRITE, lHandle)
        If lResult = ERROR_SUCCESS Then
            'The key was successfully opened, now we need to enumerate the sub keys
            Select Case lDataType
                Case REG_DWORD
                    Select Case VarType(vValue)
                        Case vbString
                            'The "0x" prefex is the registries way of saying hex.  If
                            'this is in the string, we'll assume hex and add the
                            '"&H".  Vb recognizes a number with the "&H" prefix as a
                            'hexadecimal number.  Any hexadecimal numbers will be
                            'automatically converted to long
                            vValue = Replace(vValue, "0x", "&H")
                            lTemp = CLng(vValue)
                        Case vbInteger, vbSingle, vbLong, vbDouble, vbCurrency
                            lTemp = CLng(vValue)
                        Case Else
                            'Unsupported data type
                            mvarLastError = ERROR_INVALID_PARAMETER
                            mvarErrorDescription = "The data type passed is unsupported.  Only strings or numbers are supported."
                            mvarErrorAction = "Attempting to set the value '" & sValueName & "' to '" & sKeyPath & "'"
                            Exit Function
                    End Select
                    
                    'Now we need to store the value to the registry
                    lResult = RegSetValueEx(lHandle, sValueName, ByVal 0&, REG_DWORD, lTemp, 4)
                Case REG_BINARY
                    binValue = vValue
                    Length = UBound(binValue) - LBound(binValue) + 1
                    'Now we need to store the value to the registry
                    lResult = RegSetValueEx(lHandle, sValueName, ByVal 0&, REG_BINARY, binValue(LBound(binValue)), Length)
                Case REG_SZ
                    Select Case VarType(vValue)
                        Case vbString
                            'We are really only interested in strings but we'll trap
                            'anything else being passed
                            sTemp = CStr(vValue)
                        Case Else
                            'Unsupported data type
                            mvarLastError = ERROR_INVALID_PARAMETER
                            mvarErrorDescription = "The data type passed is unsupported.  Only strings are supported for REG_SZ."
                            mvarErrorAction = "Attempting to set the value '" & sValueName & "' to '" & sKeyPath & "'"
                            Exit Function
                    End Select
                    'Now we need to store the value to the registry
                    lResult = RegSetValueEx(lHandle, sValueName, ByVal 0&, REG_SZ, ByVal sTemp, Len(sTemp))
                Case REG_EXPAND_SZ
                    Select Case VarType(vValue)
                        Case vbString
                            'We are really only interested in strings but we'll trap
                            'anything else being passed
                            sTemp = CStr(vValue)
                        Case Else
                            'Unsupported data type
                            mvarLastError = ERROR_INVALID_PARAMETER
                            mvarErrorDescription = "The data type passed is unsupported.  Only strings are supported."
                            mvarErrorAction = "Attempting to set the value '" & sValueName & "' to '" & sKeyPath & "'"
                            Exit Function
                    End Select
                    
                    'Now we need to store the value to the registry
                    lResult = RegSetValueEx(lHandle, sValueName, ByVal 0&, REG_EXPAND_SZ, ByVal sTemp, Len(sTemp))
                Case REG_MULTI_SZ
                    Select Case VarType(vValue)
                        Case vbString
                            'We are really only interested in strings but we'll trap
                            'anything else being passed
                            sTemp = CStr(vValue)
                            If Not Right(sTemp, 2) = Chr(0) & Chr(0) Then
                                If Right(sTemp, 1) = Chr(0) Then
                                    sTemp = sTemp & Chr(0)
                                Else
                                    sTemp = sTemp & Chr(0) & Chr(0)
                                End If
                            End If
                        Case Else
                            'Unsupported data type
                            mvarLastError = ERROR_INVALID_PARAMETER
                            mvarErrorDescription = "The data type passed is unsupported.  Only strings are supported."
                            mvarErrorAction = "Attempting to set the value '" & sValueName & "' to '" & sKeyPath & "'"
                            Exit Function
                    End Select
                    
                    'Now we need to store the value to the registry
                    lResult = RegSetValueEx(lHandle, sValueName, ByVal 0&, REG_MULTI_SZ, ByVal sTemp, Len(sTemp))
                Case Else
                    'Unsupported type
                    mvarLastError = ERROR_INVALID_PARAMETER
                    mvarErrorDescription = "You selected in unknown data type.  The only data types supported are REG_DWORD, REG_BINARY, REG_SZ, REG_EXPAND_SZ and REG_MULTI_SZ."
                    mvarErrorAction = "Attempting to set the value '" & sValueName & "' to '" & sKeyPath & "'"
                    Exit Function
            End Select
        
            If lResult = ERROR_SUCCESS Then
                SetRegistryValue = True
            Else
                'An error occurred
                mvarLastError = lResult
                mvarErrorDescription = GetErrorMessage(lResult)
                mvarErrorAction = "Attempting to set the value '" & sValueName & "' from '" & sKeyPath & "'"
            End If
        Else
            If CreateKeyifNotExist And lResult = 2 Then
                CreateKey , sKeyPath
                GoTo KeyCreated
            End If
            'An error occurred
            mvarLastError = lResult
            mvarErrorDescription = GetErrorMessage(lResult)
            mvarErrorAction = "Attempting to set the value '" & sValueName & "' from '" & sKeyPath & "'"
        End If
    Else
        'No key was passed, log the error
        mvarLastError = ERROR_BAD_PATHNAME
        mvarErrorDescription = "No key path was supplied"
        mvarErrorAction = "Attempting to set the value '" & sValueName & "' to '" & sValue & "'"
    End If
    
    'Close the registry handle
    If lHandle Then RegCloseKey lHandle

End Function
Public Function GetRemoteRegistryHandle(ByVal sMachineName As String, ByVal lRemoteHKey As HKeys) As Long

    Dim lResult As Long
    Dim lHandle As Long
    
    'Clear any error values
    Call ClearError
    
    'Make sure the machine name is properly formatted
    If sMachineName = "" Then
        'Passing NULL will make the API use the local machine name
        sMachineName = vbNullString
    Else
        If Not Left(sMachineName, 2) = "\\" Then
            sMachineName = "\\" & sMachineName
        End If
    End If
    
    'This function attempts to connect to the specified root key on a remote machine
    lResult = RegConnectRegistry(sMachineName, lRemoteHKey, lHandle)
    If lResult = ERROR_SUCCESS Then
        GetRemoteRegistryHandle = lHandle
    Else
        'An error occurred
        mvarLastError = lResult
        mvarErrorDescription = GetErrorMessage(lResult)
        mvarErrorAction = "Attempting to connect to the registry on '" & sMachineName & "'"
    
        GetRemoteRegistryHandle = 0
    End If

End Function
Public Function LoadHive(ByVal sFileName As String, ByVal sKeyName As String, Optional lHkey As Variant) As Boolean

    'This function will load a registry hive from a file.  A new key, with the supplied
    'name, will be created under the currently set lHKey.  This new key will contain
    'the contents of the loaded hive.
    Dim lResult As Long
    
    'Clear any error values
    Call ClearError
    
    LoadHive = False
    
    If IsMissing(lHkey) Then
        lHkey = mvarhkey
    Else
        lHkey = CLng(lHkey)
    End If
    
    'Set the required privileges
    If SetPrivilege(SE_BACKUP_NAME) And SetPrivilege(SE_RESTORE_NAME) Then
        'Load the registry hive
        lResult = RegLoadKey(lHkey, sKeyName, sFileName)
        
        If lResult = ERROR_SUCCESS Then
            LoadHive = True
        Else
            'An error occurred
            mvarLastError = lResult
            mvarErrorDescription = Replace(GetErrorMessage(lResult), vbCrLf, " ")
            mvarErrorAction = "Attempting to load the registry file '" & sFileName & "' into '" & sKeyName & "'"
        End If
    Else
        'An error occurred
        lResult = Err.LastDllError
        mvarLastError = lResult
        mvarErrorDescription = GetErrorMessage(lResult)
        mvarErrorAction = "Attempting to load the registry file '" & sFileName & "' into '" & sKeyName & "'"
    End If
    
End Function
Public Function UnLoadHive(ByVal sKeyName As String, Optional lHkey As Variant) As Boolean

    'Unload a previously Loaded registry Hive.
    Dim lResult As Long
    
    'Clear any error values
    Call ClearError
    
    UnLoadHive = False
    
    If IsMissing(lHkey) Then
        lHkey = mvarhkey
    Else
        lHkey = CLng(lHkey)
    End If

    'Attemp to unload the registry hive
    lResult = RegUnLoadKey(lHkey, sKeyName)
    If lResult = ERROR_SUCCESS Then
        UnLoadHive = True
    Else
        'An error occurred
        mvarLastError = lResult
        mvarErrorDescription = GetErrorMessage(lResult)
        mvarErrorAction = "Attempting to UnLoad the registry hive '" & sKeyName & "'"
    End If

End Function
Public Function DoesKeyExist(Optional ByVal lHkey As Variant, Optional ByVal sKeyPath As Variant) As Boolean

    'This function simply attempts to open a key and return true if the key was found
    Dim lHandle As Long
    Dim lResult As Long
    
    If IsMissing(lHkey) Then
        lHkey = mvarhkey
    Else
        lHkey = CLng(lHkey)
    End If
    
    If IsMissing(sKeyPath) Then
        sKeyPath = mvarKeyPath
    Else
        sKeyPath = CStr(sKeyPath)
    End If
    
    'Try and open the key
    lResult = RegOpenKeyEx(lHkey, sKeyPath, 0, KEY_READ, lHandle)
    If lResult = ERROR_SUCCESS Then
        DoesKeyExist = True
    Else
        DoesKeyExist = False
    End If

End Function
Private Function SetPrivilege(sPrivilege As String) As Boolean

    'This function will attempt to apply the seBackupPrivilege and seRestorePrivilege
    'privileges to the programs process token.  These privileges are required to be
    'able to use the RegLoadKey and RegUnloadKey functions.  Of course this won't work
    'if the current user doesn't have the privilege to use AdjustTokenPrivileges.
    Dim lHandle As Long
    Dim lReturn As Long
    Dim lError As Long
    Dim udtTokenPrivs1 As TOKEN_PRIVILEGES
    Dim udtTokenPrivs2 As TOKEN_PRIVILEGES
    Dim lLuid As luid
    Dim lBufferSize As Long
    Dim i As Integer
    
    'Initialize the return value
    SetPrivilege = False
    
    'Get a handle to the thread token
    lReturn = OpenThreadToken(GetCurrentThread, TOKEN_MY_ACCESS, 0, lHandle)
    If (lReturn = 0) Then
        'Error occurred
        lError = Err.LastDllError
        If (lError = ERROR_NO_TOKEN) Then
            'Try to get the process token
            lReturn = OpenProcessToken(GetCurrentProcess, TOKEN_MY_ACCESS, lHandle)
            If (lReturn = 0) Then
                'Error occurred
                lError = Err.LastDllError
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If

    'Lookup the LUID of the requested privilege
    lReturn = LookupPrivilegeValue(vbNullString, sPrivilege, lLuid)
    If lReturn = 0 Then
        'We were not successful, get the error ID and log the error
        lError = Err.LastDllError
        Exit Function
    End If
    
    'Now we want to try and get the token information, we will first call the API with
    'invalid information so we can get the size of the buffer we need
    Call GetTokenInformation(lHandle, ByVal TokenPrivileges, 0, 0, lBufferSize)
    ReDim DataBuffer((lBufferSize \ 4) - 1) As Long
        
    lReturn = GetTokenInformation(lHandle, ByVal TokenPrivileges, DataBuffer(0), lBufferSize, lBufferSize)
    If lReturn = 0 Then
        'We were not successful, get the error ID and log the error
        lError = Err.LastDllError
        Exit Function
    End If
    
    'Copy the buffer information into our structure
    Call CopyMemory(udtTokenPrivs1, DataBuffer(0), Len(udtTokenPrivs1))
    
    'Loop through the structure and try and find the privilege we are interested in
    For i = 0 To udtTokenPrivs1.PrivilegeCount
        'Compare the current privilege's luid with the luid we looked up earlier
        If udtTokenPrivs1.Privileges(i).pLuid.LowPart = lLuid.LowPart And udtTokenPrivs1.Privileges(i).Attributes = SE_PRIVILEGE_ENABLED Then
            'The privilege is already set
            SetPrivilege = True
            Exit Function
        End If
    Next i
    
    'If we are here, then the privilege in question was not already set for the
    'token.  Now we need to initialize a new token privilege structure with the LUID
    udtTokenPrivs2.PrivilegeCount = 1
    udtTokenPrivs2.Privileges(0).pLuid.LowPart = lLuid.LowPart
    udtTokenPrivs2.Privileges(1).pLuid.HighPart = lLuid.HighPart
    udtTokenPrivs2.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    
    'Enable or disable the privilege
    AdjustTokenPrivileges lHandle, 0, udtTokenPrivs2, 16, 0, 0
    
    'Get the error code to determine if we were successful
    lError = Err.LastDllError
    If (lError <> ERROR_SUCCESS) And (lError <> ERROR_NOT_ALL_ASSIGNED) Then
        'Yes, we had an error
        Exit Function
    Else
        'Possible error
        If (lError = ERROR_NOT_ALL_ASSIGNED) Then
        End If
    End If
    
    'Return true
    SetPrivilege = True
    
End Function
Private Function GetErrorMessage(lErrorNumber As Long) As String
    
    'This function takes a windows error number and looks up the description for it
    Dim sBuffer As String
    
    sBuffer = Space(200)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrorNumber, LANG_NEUTRAL, sBuffer, 200, ByVal 0&
    
    GetErrorMessage = Replace(TrimNull(sBuffer), "file", "key")
    
End Function
Private Function TrimNull(sItem As String) As String
Attribute TrimNull.VB_Description = "Internal function used to remove the nulls from a string returned by various API functions."
    
    'Finds a null then trims the string
    Dim iPos As Integer
        
    iPos = InStr(sItem, vbNullChar)
    If iPos > 0 Then TrimNull = Left(sItem, iPos - 1) Else TrimNull = sItem

End Function
Private Sub ClearError()

    'Clear any current error information
    mvarLastError = 0
    mvarErrorDescription = ""
    mvarErrorAction = ""
    
End Sub

