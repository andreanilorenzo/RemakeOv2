VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsManifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**************************************
' Please see http://www.1JavaStreet.com/vb/scripts/ShowCode.asp?txtCodeId=51893&lngWId=1
'**************************************

'-------------------------------------------------------------------------------------
' Per la funzione che ignora gli errori di windows
' la uso alla chiusura dell'applicazione per colpa dell'errore
' che a volte si verifica con il file .exe.manifest
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Const SEM_FAILCRITICALERRORS = &H1
Const SEM_NOGPFAULTERRORBOX = &H2
Const SEM_NOOPENFILEERRORBOX = &H8000&
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
' Per cambiare gli attributi di un file
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
'-------------------------------------------------------------------------------------

Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private m_hMod As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Private Sub Class_Initialize()
    
    If Var(UsaManifestFile).Valore = 0 Then
        Call DeleteManifest
    ElseIf Var(UsaManifestFile).Valore = 1 Then
        Call CreateManifest
    End If

    m_hMod = LoadLibrary("shell32.dll")
    InitCommonControls
    
End Sub

Private Sub Class_Terminate()

    FreeLibrary m_hMod
    
    ' Previene tutte le finestre di dialogo e i message box di errore del sistema (UAE):
    SetErrorMode SEM_NOGPFAULTERRORBOX
    
    ' Termino il programma
    End

End Sub

Public Sub DeleteManifest()
    Dim EXEPath As String
    
    ' Get The EXE Path
    EXEPath = App.Path & IIf(Right(App.Path, 1) = "\", vbNullString, "\")
    EXEPath = EXEPath & App.EXEName & IIf(LCase(Right(App.EXEName, 4)) = ".exe", ".manifest", ".exe.manifest")

    ' Controllo se il file manifest esiste
    If Dir(EXEPath, vbReadOnly Or vbSystem Or vbHidden) <> vbNullString Then
        SetFileAttributes EXEPath, FILE_ATTRIBUTE_NORMAL
        DoEvents
        Kill EXEPath
        DoEvents
    End If

End Sub

Private Function CreateManifest() As Boolean
    On Error Resume Next
    Dim EXEPath As String
    
    ' Get The EXE Path
    EXEPath = App.Path & IIf(Right(App.Path, 1) = "\", vbNullString, "\")
    EXEPath = EXEPath & App.EXEName & IIf(LCase(Right(App.EXEName, 4)) = ".exe", ".manifest", ".exe.manifest")
    
    ' Checks if the manifest has already been created
    If Dir(EXEPath, vbReadOnly Or vbSystem Or vbHidden) <> vbNullString Then GoTo ErrorHandler
    
    'Makes sure you are using windows xp
    If WinVersion = "Windows XP" Then
        Dim iFileNumber As Integer
        iFileNumber = FreeFile
        'Save the .manifest file
        Open EXEPath For Output As #iFileNumber
        Print #iFileNumber, FormatManifest
        CreateManifest = True
    Else
        Kill EXEPath
    End If
    
    ' Set the file to be hidden
    Close #iFileNumber
    SetAttr EXEPath, vbHidden Or vbSystem Or vbReadOnly Or vbArchive
    
ErrorHandler:
    m_hMod = LoadLibrary("shell32.dll")
    Call InitCommonControls
    DoEvents
    
End Function

Private Function WinVersion() As String
    ' Get windows version (from Microsoft.com)
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

    With osinfo
        Select Case .dwPlatformID
            Case 1
                If .dwMinorVersion = 0 Then
                    WinVersion = "Windows 95"
                ElseIf .dwMinorVersion = 10 Then
                    WinVersion = "Windows 98"
                End If
            Case 2
                If .dwMajorVersion = 3 Then
                    WinVersion = "Windows NT 3.51"
                ElseIf .dwMajorVersion = 4 Then
                    WinVersion = "Windows NT 4.0"
                ElseIf .dwMajorVersion >= 5 Then
                    WinVersion = "Windows XP"
                End If
            Case Else
                WinVersion = "Failed"
        End Select
    End With
    
End Function

Private Function FormatManifest() As String
    'Create the string for the manifest file
    Dim Header As String
    
    Header = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>"
    Header = Header & vbCrLf & "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">"
    Header = Header & vbCrLf & "<assemblyIdentity"
    Header = Header & vbCrLf & "version=" & Chr(34) & "1.0.0.0" & Chr(34)
    Header = Header & vbCrLf & "processorArchitecture=" & Chr(34) & "X86" & Chr(34)
    Header = Header & vbCrLf & "name=" & Chr(34) & "Microsoft.VisualBasic6.IDE" & Chr(34)
    Header = Header & vbCrLf & "type=" & Chr(34) & "win32" & Chr(34)
    Header = Header & vbCrLf & "/>"
    Header = Header & vbCrLf & "<description>Microsoft Visual Basic 6 IDE</description>"
    Header = Header & vbCrLf & "<dependency>"
    Header = Header & vbCrLf & "<dependentAssembly>"
    Header = Header & vbCrLf & "<assemblyIdentity"
    Header = Header & vbCrLf & "type=" & Chr(34) & "win32" & Chr(34)
    Header = Header & vbCrLf & "name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34)
    Header = Header & vbCrLf & "version=" & Chr(34) & "6.0.0.0" & Chr(34)
    Header = Header & vbCrLf & "processorArchitecture=" & Chr(34) & "X86" & Chr(34)
    Header = Header & vbCrLf & "publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34)
    Header = Header & vbCrLf & "language=" & Chr(34) & "*" & Chr(34)
    Header = Header & vbCrLf & "/>"
    Header = Header & vbCrLf & "</dependentAssembly>"
    Header = Header & vbCrLf & "</dependency>"
    Header = Header & vbCrLf & "</assembly>"

    FormatManifest = Header
    
End Function
