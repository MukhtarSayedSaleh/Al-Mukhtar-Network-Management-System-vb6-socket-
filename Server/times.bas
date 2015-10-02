Attribute VB_Name = "Module1"
'****************************************
'* Al-Mokhtar For Networks Server       *
'*  By : Mokhtar saied saleh            *
'*      Syria - Abokamal                *
'*      WWW.ABOKAMAL.COM                *
'*  MOKHTAR_SS@HOTMAIL.COM              *
'*       0096394467547                  *
'****************************************
Public StartTime As String      ' Holds the begining time in short format
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const SYNCHRONIZE = &H100000
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_NOTIFY = &H10
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_SET_VALUE = &H2
Public Const KEY_QUERY_VALUE = &H1
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private lngTopKey As Long
Private strSubKey As String

Public Sub DeleteRegValue(strKey As String, strValue As String)
  Dim hKey As Long

  ParseKey strKey
  RegOpenKeyEx lngTopKey, strSubKey, 0&, KEY_ALL_ACCESS, hKey
  RegDeleteValue hKey, strValue
  RegCloseKey hKey
End Sub

Private Sub ParseKey(strKey As String)
  Dim intPos As Integer
  
  intPos = InStr(strKey, "\")

  Select Case Left(strKey, intPos - 1)
    Case "HKEY_CLASSES_ROOT"
      lngTopKey = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
      lngTopKey = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
      lngTopKey = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
      lngTopKey = HKEY_USERS
    Case "HKEY_CURRENT_CONFIG"
      lngTopKey = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
      lngTopKey = HKEY_DYN_DATA
  End Select
  strSubKey = Right(strKey, Len(strKey) - intPos)
End Sub



Public Function GetMins(tm1 As String, tm2 As Date) As Long
    Dim m1 As Long, m2 As Long
    Dim strm1 As String
    
    strm1 = Right(tm1, InStr(tm1, ":") - 1)
    m1 = strm1
   
    m2 = Minute(tm2)
    
    
    GetMins = m2 - m1

End Function
Public Function GetHours(tm1 As String, tm2 As Date) As Long
    
    Dim h1 As Long, h2 As Long


    Dim strh1 As String
    strh1 = Left(tm1, InStr(tm1, ":") - 1)
    h1 = strh1
    
    h2 = Hour(tm2)
        
    GetHours = h2 - h1

End Function

