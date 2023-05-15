Attribute VB_Name = "InternetExplorer"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

'-------------------------------------------------------------------------------------------------------
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000
'-------------------------------------------------------------------------------------------------------

Private Declare Function DoFileDownload Lib "shdocvw" (ByVal lpszFile As String) As Long

Type DllVersionInfo
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

Private Const NOERROR As Long = 0
Private Declare Function DllGetVersion Lib "Shlwapi.dll" (pdvi As DllVersionInfo) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

Public Function DownloadFile(sSourceUrl As String, sLocalFile As String) As Boolean
  ' Download the file.
  ' BINDF_GETNEWESTVERSION forces the API to download from the specified source.
  ' Passing 0& as dwReserved causes the locally-cached copy to be downloaded, if available.
  ' If the API Returns ERROR_SUCCESS (0), DownloadFile returns True.
   DownloadFile = URLDownloadToFile(0&, sSourceUrl, sLocalFile, BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS
   
End Function

Public Sub DownloadFileDialog(sDownload As String)
    
    If sDownload = "" Then
        MsgBox "Non è possibile scaricare il file.", vbInformation, App.Path
        Exit Sub
    End If
    
    sDownload = StrConv(sDownload, vbUnicode)
    Call DoFileDownload(sDownload)
   
End Sub

Private Function GetIEVersion() As String
    Dim DVI As DllVersionInfo
    Dim r As Long
    
    DVI.cbSize = Len(DVI)
    r = DllGetVersion(DVI)
    
    If r = NOERROR Then
        GetIEVersion = DVI.dwMajorVersion & "." & DVI.dwMinorVersion & "." & DVI.dwBuildNumber
    Else
        ' There was an error.. Might be because
        ' IE isn't installed.
        GetIEVersion = "ERROR"
    End If
    
End Function

Public Function IEversion(Optional Full As Boolean = True) As String
    
    If Full = True Then
        IEversion = GetIEVersionFriendlyName & " (" & GetIEVersion & ")"
    Else
         IEversion = "Internet Explorer " & GetIEVersion()
    End If

End Function

Private Function GetIEVersionFriendlyName() As String

   Dim s As String
   Dim DVI As DllVersionInfo
   
   Call GetIEVersionDVI(DVI)
   
   Select Case DVI.dwMajorVersion
      Case 4
      
         Select Case DVI.dwMinorVersion
            Case 40:
               Select Case DVI.dwBuildNumber
                  Case 308: s = "1.0 (Plus! for Windows 95)"
                  Case 520: s = "2.0"
               End Select
            Case 70
               Select Case DVI.dwBuildNumber
                  Case 1155: s = "3.0"
                  Case 1158: s = "3.0 (OSR2)"
                  Case 1215: s = "3.01"
                  Case 1300: s = "3.02 and 3.02a"
                  Case Else: s = "3 (Unknown)"
               End Select
            Case 71
               Select Case DVI.dwBuildNumber
                  Case 544: s = "4.0 Platform Preview 1.0 (PP1)"
                  Case 1008: s = "4.0 4.0 Platform Preview 2.0 (PP2)"
                  Case 1712: s = "4.0"
                  Case Else: s = "4.0 (Unknown)"
               End Select
            Case 72
               Select Case DVI.dwBuildNumber
                  Case 2106: s = "4.01"
                  Case 3110: s = "4.01 Service Pack 1 (Windows 98)"
                  Case 3612: s = "4.01 Service Pack 2"
                  Case 3711: s = "4.x with Update"
                  Case Else: s = "4.0 (Unknown)"
               End Select
            Case Else: s = "(Unknown)"
         End Select
         
      Case 5
      
         Select Case DVI.dwMinorVersion
            Case 0
               Select Case DVI.dwBuildNumber
                  Case 518: s = "5 Developer Preview (Beta 1)"
                  Case 910: s = "5 Beta (Beta 2)"
                  Case 2014: s = "5"
                  Case 2314: s = "5 (Office 2000)"
                  Case 2516: s = "5.01 (Windows 2000 Beta 3, build 5.00.2031)"
                  Case 2614: s = "5 (Windows 98 Second Edition)"
                  Case 2717, 2721, 2723: s = "5 with update"
                  Case 2919: s = "5.01 (Windows 2000 RC1&2/Office 2000 SR-1/Update)"
                  Case 2920: s = "5.01 (Windows 2000, build 5.00.2195)"
                  Case 3103: s = "5.01 Service Pack 1 (Windows 2000)"
                  Case 3105: s = "5.01 Service Pack 1 (Windows 95/98 and Windows NT 4.0)"
                  Case 3314: s = "5.01 Service Pack 2 (Windows 95/98 and Windows NT 4.0)"
                  Case 3315: s = "5.01 Service Pack 2 (Windows 2000)"
                  Case 3502: s = "5.01 SP3 (Windows 2000 SP3 only)"
                  Case 3700: s = "5.01 SP4 (Windows 2000 SP4 only)"
                  Case Else: s = "5 (Unknown)"
               End Select
            
            Case 50
            
               Select Case DVI.dwBuildNumber
                  Case 3825: s = "5.5 Developer Preview (Beta)"
                  Case 4030: s = "5.5 & Internet Tools Beta"
                  Case 4134: s = "5.5"
                  Case 4308: s = "5.5 Advanced Security Privacy Beta"
                  Case 4522: s = "5.5 Service Pack 1"
                  Case 4807: s = "5.5 Service Pack 2"
                  Case Else: s = "5.5 (Unknown)"
               End Select
               
            Case Else: s = GetIEVersion()
         End Select
         
         Case 6
            Select Case DVI.dwMinorVersion
               Case 0
               
                  Select Case DVI.dwBuildNumber
                     Case 2462: s = "6 Public Preview (Beta)"
                     Case 2479: s = "6 Public Preview (Beta) Refresh"
                     Case 2600: s = "6"
                     Case 2800: s = "6 Service Pack 1"
                     Case 2900: s = "6 Service Pack 2"
                     Case 3663: s = "6 for Microsoft Windows Server 2003 RC1"
                     Case 3718: s = "6 for Windows Server 2003 RC2"
                     Case 3790: s = "6 for Windows Server 2003 (release)"
                     Case Else: s = "6 (Unknown build)"
                  End Select
                  
               Case Else: s = GetIEVersion()
         End Select
      Case Else: s = GetIEVersion()
   End Select
   
   GetIEVersionFriendlyName = "Internet Explorer " & s
   
End Function

Private Function GetIEVersionDVI(DVI As DllVersionInfo) As Long
   
   DVI.cbSize = Len(DVI)
   Call DllGetVersion(DVI)
   GetIEVersionDVI = DVI.dwMajorVersion
   
End Function

