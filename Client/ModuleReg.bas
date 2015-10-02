Attribute VB_Name = "ModuleReg"
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal _
    hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass _
    As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes _
    As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal _
    hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType _
    As Long, lpData As Any, ByVal cbData As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, _
    ByVal lpSubKey As String) As Long

Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

'nycklar som kan användas
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006 'endast Windows 95/98
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004 'endast Windows NT/2000
Public Const HKEY_USERS = &H80000003

'skriv/läs åtkomst för registret
Public Const KEY_ALL_ACCESS = &HF003F 'Permission for all types of access
Public Const KEY_CREATE_LINK = &H20 'Permission to create symbolic links
Public Const KEY_CREATE_SUB_KEY = &H4 'Permission to create subkeys
Public Const KEY_ENUMERATE_SUB_KEYS = &H8 'Permission to enumerate subkeys
Public Const KEY_EXECUTE = &H20019 'Same as KEY_READ
Public Const KEY_NOTIFY = &H10 'Permission to give change notification
Public Const KEY_QUERY_VALUE = &H1 'Permission to query subkey data
Public Const KEY_READ = &H20019 'Permission for general read access
Public Const KEY_SET_VALUE = &H2 'Permission to set subkey data
Public Const KEY_WRITE = &H20006 'Permission for general write access

'datatyp för det nya värdet som skrivs in används i RegSetValueEx Function
Public Const REG_BINARY = 3 'A non-text sequence of bytes
Public Const REG_DWORD = 4 'Same as REG_DWORD_LITTLE_ENDIAN
Public Const REG_DWORD_BIG_ENDIAN = 5 'A 32-bit integer stored in big-endian format. This is the opposite of the way Intel-based computers normally store numbers the word order is reversed
Public Const REG_DWORD_LITTLE_ENDIAN = 4 'A 32-bit integer stored in little-endian format. This is the way Intel-based computers normally store numbers
Public Const REG_EXPAND_SZ = 2 'A null-terminated string which contains unexpanded environment variables
Public Const REG_LINK = 6 'A Unicode symbolic link
Public Const REG_MULTI_SZ = 7 'A series of strings, each separated by a null character and the entire set terminated by a two null characters
Public Const REG_NONE = 0 'No data type
Public Const REG_RESOURCE_LIST = 8 'A list of resources in the resource map
Public Const REG_SZ = 1 'A string terminated by a null character




