Attribute VB_Name = "RegirtaEstensione"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

'========Read registry key values
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
'Note that if you declare the lpData parameter as String, you must pass it By Value. (In RegQueryValueEx)
Public phkResult As Long
Public lpSubKey As String
Public lpData As String
Public lpcbData As Long
Public RC As Long
'Root Key Constants ...................................
'Public Const HKEY_CLASSES_ROOT = &H80000000
'Reg DataType Constants ...............................
'Public Const REG_SZ = 1 ' Unicode null terminated string
'===============Create and delete key in registry
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
      ' Return codes from Registration functions.
      Const ERROR_SUCCESS = 0&
      Const ERROR_BADDB = 1&
      Const ERROR_BADKEY = 2&
      Const ERROR_CANTOPEN = 3&
      Const ERROR_CANTREAD = 4&
      Const ERROR_CANTWRITE = 5&
      Const ERROR_OUTOFMEMORY = 6&
      Const ERROR_INVALID_PARAMETER = 7&
      Const ERROR_ACCESS_DENIED = 8&
      Private Const MAX_PATH = 260&
      '==included in Read registry key values
      'Private Const HKEY_CLASSES_ROOT = &H80000000
      'Private Const REG_SZ = 1
'This sub puts new default icon on associated files or off if unassociated
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0&
Private Const SHCNF_FLUSHNOWAIT As Long = &H2000


'Public Enum ERegistryClassConstants
'    HKEY_CLASSES_ROOT = &H80000000
'    HKEY_CURRENT_USER = &H80000001
'    HKEY_LOCAL_MACHINE = &H80000002
'    HKEY_USERS = &H80000003
'End Enum

'Public Enum ERegistryValueTypes
    'Predefined Value Types
'    REG_NONE = (0)                         'No value type
'    REG_SZ = (1)                           'Unicode nul terminated string
'    REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
'    REG_BINARY = (3)                       'Free form binary
'    REG_DWORD = (4)                        '32-bit number
'    REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
'    REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
'    REG_LINK = (6)                         'Symbolic Link (unicode)
'    REG_MULTI_SZ = (7)                     'Multiple Unicode strings
'    REG_RESOURCE_LIST = (8)                'Resource list in the resource map
'    REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
'    REG_RESOURCE_REQUIREMENTS_LIST = (10)
'End Enum

Global Const gstrSEP_DIR$ = "\"
Global Const resERR_REG = 462
'Const ERROR_SUCCESS = 0&
'Const REG_SZ = 1

Public Const KEY_READ = &H20019
Public Const KEY_WRITE = &H20006
Public Const KEY_ALL_ACCESS = &HF003F

'Private Const REG_EXPAND_SZ = 2
Private eValueType As ERegistryValueTypes

'Const STANDARD_RIGHTS_ALL = &H1F0000
'Const STANDARD_RIGHTS_READ = &H20000
'Const STANDARD_RIGHTS_WRITE = &H20000
'Const SYNCHRONIZE = &H100000

Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
'Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
'Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
'Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
'Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
'Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
'Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Declare Function OSRegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, lpbData As Any, ByVal cbData As Long) As Long
Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long

' Hkey cache (used for logging purposes)
Private Type HKEY_CACHE
    hKey As Long
    strHkey As String
End Type

Private hkeyCache() As HKEY_CACHE

'------------------------------------------------------------------------------------------
Private Declare Function FindExecutable Lib "shell32" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal sResult As String) As Long
'Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_FILE_SUCCESS As Long = 32 'my constant
Private Const ERROR_BAD_FORMAT As Long = 11
'------------------------------------------------------------------------------------------

Public Function CercaProgramma(lpFile As String, Optional ldDirectory As String = "") As String
   Dim success As Long
   Dim Pos As Long
   Dim Msg As String
   Dim sResult As String
   
   sResult = Space$(MAX_PATH)
   
   If ldDirectory = "" Then ldDirectory = App.Path

  'lpFile: name of the file of interest
  'lpDirectory: location of lpFile
  'sResult: path and name of executable associated with lpFile
   success = FindExecutable(lpFile, lpFile, sResult)
      
   Select Case success
      Case ERROR_FILE_NO_ASSOCIATION: Msg = "no association"
      Case ERROR_FILE_NOT_FOUND: Msg = "File non trovato"
      Case ERROR_PATH_NOT_FOUND: Msg = "path not found"
      Case ERROR_BAD_FORMAT:     Msg = "bad format"
      Case Is >= ERROR_FILE_SUCCESS:
         Pos = InStr(sResult, Chr$(0))
         If Pos Then
            Msg = Left$(sResult, Pos - 1)
         End If
   End Select
   
   CercaProgramma Msg
   
End Function

Public Function IsRegistryEditable() As Boolean
    ' Determines whether Registry tools have been disabled for the current user
    Dim lValue As Long       ' Variable for value
    Dim sKey As String       ' Key to open
    Dim hKey As Long         ' Handle to registry key
    
    eValueType = REG_DWORD

    sKey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    
    If RegOpenKeyEx(HKEY_CURRENT_USER, sKey, 0, KEY_READ, hKey) <> ERROR_SUCCESS Then
       ' Key does not exist, return True
       IsRegistryEditable = True
    Else
       ' Determine if value exists
       If RegQueryValueEx(hKey, "DisableRegistryTools", 0, eValueType, lValue, Len(lValue)) <> ERROR_SUCCESS Then
       
          ' value does not exist, return True
          IsRegistryEditable = True
       Else
          ' Return opposite of value (0 = Editable, 1 = Disable)
          IsRegistryEditable = Not CBool(lValue)
       End If
    End If

End Function

Public Function GetAssociatedApp(sExten As String) As String
   Dim sBuffer As String, sProgName As String
   Dim sPath As String
   Dim lBuffer As Long, lProgName As Long
   Dim hKey As Long, hProgKey As Long

   sBuffer = Space(20)
   lBuffer = Len(sBuffer)

   ' Open Key
   If RegOpenKeyEx(HKEY_CLASSES_ROOT, sExten, 0, KEY_READ, hKey) <> ERROR_SUCCESS Then
      ' Key does not exist, return null string
      GetAssociatedApp = vbNullString
   Else
      Dim lType As Long
      
      ' Get key's unnamed value
      RegQueryValueEx hKey, vbNullString, 0, 0, ByVal sBuffer, lBuffer
      RegCloseKey hKey
      sBuffer = Left(sBuffer, lBuffer - 1)
      
      ' Open Command key of File Association key's Open subkey
      sPath = sBuffer & "\shell\open\command"
      If RegOpenKeyEx(HKEY_CLASSES_ROOT, sPath, 0, KEY_READ, hProgKey) = ERROR_SUCCESS Then
         ' Determine data type and buffer size of key
         RegQueryValueEx hProgKey, vbNullString, 0, 0, ByVal vbNull, lProgName
      
         ' Retrieve file association
         sProgName = Space(lProgName + 1)
         RegQueryValueEx hProgKey, vbNullString, 0, lType, ByVal sProgName, lProgName
         RegCloseKey hProgKey
         sProgName = Left(sProgName, lProgName - 1)
      
         ' Check if environment string is present
         If lType = REG_EXPAND_SZ Then
            Dim lProg As Long
            Dim sProg As String
         
            sProg = Space(MAX_PATH + 1)
            lProg = ExpandEnvironmentStrings(sProgName, sProg, Len(sProg))
            sProgName = Left(sProg, lProg - 1)
         End If
         GetAssociatedApp = sProgName
      Else
         GetAssociatedApp = ""
      End If
   End If
   
End Function

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    Dim hCurKey As Long
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long

    'Set up default value
    If Not IsEmpty(Default) Then
        GetSettingString = Default
    Else
        GetSettingString = ""
    End If

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuffer = String(lDataBufferSize, " ")
            lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)

            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetSettingString = Left$(strBuffer, intZeroPos - 1)
            Else
                GetSettingString = strBuffer
            End If
        End If
    Else
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)
    
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim hCurKey As Long
    Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)

    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

    If lRegResult <> ERROR_SUCCESS Then
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)
    
End Sub

Public Sub RegistraFileExtension(Extension As String, PathToExecute As String, ApplicationName As String)
    'Extension is three letters without the "."
    'PathToExecute is full path to exe file
    'Application Name is any name you want as description of Extension
    Dim sKeyName As String   'Holds Key Name in registry.
    Dim sKeyValue As String  'Holds Key Value in registry.
    Dim ret&           'Holds error status, if any, from API calls.
    Dim lphKey&        'Holds created key handle from RegCreateKey.

    If Left$(Extension, 1) = "." Then Extension = Right$(Extension, Len(Extension) - 1)

    'This creates a Root entry called 'ApplicationName'.
    sKeyName = ApplicationName
    sKeyValue = ApplicationName
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)

    'This creates a Root entry for the extension to be associated with 'ApplicationName'.
    sKeyName = "." & Extension
    sKeyValue = ApplicationName
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)

    'This sets the command line for 'ApplicationName'.
    sKeyName = ApplicationName
    sKeyValue = PathToExecute & " %1"
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)

    'This sets the default icon
    sKeyName = ApplicationName
    sKeyValue = App.Path & "\" & App.EXEName & ".exe,0"
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH)

    'Force Icon Refresh
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
    
End Sub

Public Sub UnRegistraFileExtension(ByVal Extension As String, ByVal ApplicationName As String, Optional TogliAncheRiferimentiAlProgramma As Boolean = True)
    Dim sKeyName As String   'Finds Key Name in registry.
    Dim sKeyValue As String  'Finds Key Value in registry.
    Dim ret&           'Holds error status, if any, from API calls.

    If Left$(Extension, 1) = "." Then Extension = Right$(Extension, Len(Extension) - 1)
    
    If TogliAncheRiferimentiAlProgramma = True Then
        'This deletes the default icon
        sKeyName = ApplicationName
        ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName & "\DefaultIcon")
    
        'This deletes the command line for "ApplicationName".
        sKeyName = ApplicationName
        ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName & "\shell\open\command")
    
        'This deletes a Root entry called "ApplicationName".
        sKeyName = ApplicationName
        ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName & "\shell\open")
    
        'This deletes a Root entry called "ApplicationName".
        sKeyName = ApplicationName
        ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName & "\shell")
    
        'This deletes a Root entry called "ApplicationName".
        sKeyName = ApplicationName
        ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName)
    End If

    'This deletes the Root entry for the extension to be associated with "ApplicationName".
    sKeyName = "." & Extension
    ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName)

    'Force Icon Refresh (ATTENZIONE: funziona solo quando il programma è compilato)
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST & SHCNF_FLUSHNOWAIT, 0, 0
    'Thanks to Ralf Gerstenberger <ralf.gerstenberger@arcor.de> for pointing out
    'that WinXP seems to require the SHCNF_FLUSHNOWAIT flag in SHChangeNotify
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/functions/shchangenotify.asp
End Sub

