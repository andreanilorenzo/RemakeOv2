VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResourceFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10
Private Const MAX_SIZE                  As Long = 260

Private Type FILETIME
       dwLowDateTime        As Long
       dwHighDateTime       As Long
End Type

Private Type WIN32_FIND_DATA
       dwFileAttributes     As Long     ' file attributes
       ftCreationTime       As FILETIME ' File creation date and time
       ftLastAccessTime     As FILETIME ' File last accessed date and time
       ftLastWriteTime      As FILETIME ' File last modified date and time
       nFileSizeHigh        As Long     ' file sizes over 2GB (2,147,483,647)
       nFileSizeLow         As Long     ' file sizes under 2GB (2,147,483,647)
       dwReserved0          As Long
       dwReserved1          As Long
       cFilename            As String * MAX_SIZE ' full file name w/o path
       cAlternate           As String * 14       ' short file name w/o path
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private m_lngRetCode    As Long

Public Function SalvaInCartella(ByVal ResourceID As Long, ByVal NomeRisorsa As String, ByVal PercorsoFile As String, ByVal NomeFile As String) As Boolean
    Dim intFile As Integer
    Dim bytData() As Byte
    Dim ret As Boolean
    
    On Error GoTo SalvaInCartella_Errors

    ret = FolderExist(PercorsoFile, True)
    PercorsoFile = AddTrailingSlash(PercorsoFile)

    ' Carico i dati della risorsa nell'array
    bytData = LoadResData(ResourceID, NomeRisorsa)
  
    ' Scrivo il file
    intFile = FreeFile
    Open PercorsoFile & NomeFile For Binary Access Write As #intFile
    Put #intFile, , bytData()
    Close #intFile
    
    ' Success  :-)
    SalvaInCartella = True
    Exit Function

SalvaInCartella_Errors:
    ' Failure  :-(
    SalvaInCartella = False

End Function

Public Function FolderExist(FullPathFolderName As String, Optional SeNonEsisteCreala As Boolean = False) As Boolean
    Dim WFD As WIN32_FIND_DATA
    Dim lngHandle As Long
   
    ' remove trailing slash before verifying
    FullPathFolderName = RemoveTrailingSlash(FullPathFolderName)

    lngHandle = FindFirstFile(FullPathFolderName, WFD)
   
    ' if a valid file handle was returned and the directory attribute is set, then the folder exists
    FolderExist = (lngHandle <> INVALID_HANDLE_VALUE) And (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
    
    ' Close the file handle
    m_lngRetCode = FindClose(lngHandle)
    
    ' Se non esiste, creo la cartella
    If SeNonEsisteCreala = True And FolderExist = False Then MkDir FullPathFolderName
    FolderExist = True
   
End Function

Public Function RemoveTrailingSlash(ByVal strPath As String) As String
    ' Removes the trailing backslash from the path if it exist
      
    strPath = Trim(strPath)
    
    If Right(strPath, 1) = "\" Then
        RemoveTrailingSlash = Left(strPath, Len(strPath) - 1)
    Else
        RemoveTrailingSlash = strPath
    End If
   
End Function

Public Function AddTrailingSlash(ByVal strPath As String) As String

    strPath = Trim(strPath)

    If Right(strPath, 1) = "\" Then
        AddTrailingSlash = strPath
    Else
        AddTrailingSlash = strPath & "\"
    End If
   
End Function
