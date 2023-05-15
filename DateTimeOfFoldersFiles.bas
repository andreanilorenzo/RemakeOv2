Attribute VB_Name = "DateTimeOfFoldersFiles"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2006 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const OFS_MAXPATHNAME = 128
Private Const OF_READWRITE = &H2

Private Type OFSTRUCT
   cBytes      As Byte
   fFixedDisk  As Byte
   nErrCode    As Integer
   Reserved1   As Integer
   Reserved2   As Integer
   szPathName(0 To OFS_MAXPATHNAME - 1) As Byte '0-based
End Type

Private Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type

Private Type SYSTEMTIME
  wYear          As Integer
  wMonth         As Integer
  wDayOfWeek     As Integer
  wDay           As Integer
  wHour          As Integer
  wMinute        As Integer
  wSecond        As Integer
  wMilliseconds  As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias         As Long
   StandardName(0 To 31) As Integer  '32, 0-based
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 31) As Integer  '32, 0-based
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hFile As Long) As Long

Private Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hwnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    lpIDList      As Long
    lpClass       As String
    hkeyClass     As Long
    dwHotKey      As Long
    hIcon         As Long
    hProcess      As Long
End Type

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Declare Function ShellExecuteEx Lib "shell32" (SEI As SHELLEXECUTEINFO) As Long

Public Function GetSistemTyme() As FILETIME
    Dim tmp As String
    Dim SYS_TIME As SYSTEMTIME
    'Dim NEW_TIME As FILETIME
  
    'obtain the local system time (adjusts for the GMT deviation of the local time zone)
    GetLocalTime SYS_TIME
    
    'tmp = "Giorno:" & SYS_TIME.wDay
    'tmp = "Mese:" & SYS_TIME.wMonth
    'tmp = "Anno:" & SYS_TIME.wYear
    'tmp = "Completo:" GetSystemDateString(SYS_TIME)

    'convert the system time to a valid file time
    Call SystemTimeToFileTime(SYS_TIME, GetSistemTyme)
       
End Function

Public Function SetDataFile(ByVal PosizioneFile As String, Optional NuovaData As Date = "01-01-1900")
    Dim hFile As Long
    Dim fName As String

    Dim NEW_TIME As FILETIME
    Dim OFS As OFSTRUCT
    
    fName = PosizioneFile
    
    ' Apro il file
    hFile = OpenFile(fName, OFS, OF_READWRITE)
    
    If NuovaData = "01-01-1900" Then
        NEW_TIME = GetSistemTyme
    End If
    
    
    
    'manca il codice qua....................
    
    
    
    'set the created, accessed and modified dates all
    'to the new dates.  A null (0&) could be passed as
    'any of the NEW_TIME parameters to leave that date unchanged.
    Call SetFileTime(hFile, NEW_TIME, NEW_TIME, NEW_TIME)

    ' Chiudo il file
    Call CloseHandle(hFile)

End Function

Public Function GetDataFile(ByVal PosizioneFile As String, Optional ByVal CheDataVuoi As String = "C") As String
    Dim hFile As Long
    Dim fName As String
    
    Dim OFS As OFSTRUCT
    Dim FT_CREATE As FILETIME
    Dim FT_ACCESS As FILETIME
    Dim FT_WRITE As FILETIME
    
    
    fName = PosizioneFile
    
    ' Apro il file
    hFile = OpenFile(fName, OFS, OF_READWRITE)
    
    ' get the FILETIME info for the created, accessed and last write info
    Call GetFileTime(hFile, FT_CREATE, FT_ACCESS, FT_WRITE)
    
    Select Case CheDataVuoi
        Case "C" ' Creazione
            GetDataFile = GetFileDateString(FT_CREATE)
        
        Case "A" ' Accesso
            GetDataFile = GetFileDateString(FT_ACCESS)
            
        Case "M" ' Modifica
            GetDataFile = GetFileDateString(FT_WRITE)
        
        Case Else
            GetDataFile = GetFileDateString(FT_WRITE)
        
    End Select
    
    ' Chiudo il file
    Call CloseHandle(hFile)
    
End Function

Private Function GetFileDateString(CT As FILETIME) As String
    Dim ST As SYSTEMTIME
    Dim ds As Single
  
    'convert the passed FILETIME to a valid SYSTEMTIME format for display
    If FileTimeToSystemTime(CT, ST) Then
        ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
        GetFileDateString = Format$(ds, "DD-MM-YYYY")
    Else
        GetFileDateString = "01-01-1900"
    End If

End Function

Private Function GetSystemDateString(ST As SYSTEMTIME) As String
    Dim ds As Single
    
    ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
    
    If ds Then
        GetSystemDateString = Format$(ds, "DD-MM-YYYY")
    Else
        GetSystemDateString = "01-01-1900"
    End If

End Function
