VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOSOutputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The CreatePipe function creates an anonymous pipe, and returns handles to the read and write ends of the pipe.
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
'Used to read the the pipe filled by the process create with the CretaProcessA function
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long

'Structure used by the CreateProcessA function
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'Structure used by the CreateProcessA function
Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

'Structure used by the CreateProcessA function
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

'This function launch the the commend and return the relative process into the PRECESS_INFORMATION structure
Private Declare Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

'Close opened handle
Private Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long

'Consts for the above functions
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1


Private mCommand As String          'Private variable for the CommandLine property
Private mOutputs As String          'Private variable for the ReadOnly Outputs property

'Event that notify the temporary buffer to the object
Public Event ReceiveOutputs(CommandOutputs As String)


'--------------------------------------------------------------------------------------------------------------------------------
' Per le funzioni ExecAndCapture e ShellExecuteCapture
'Private Const STARTF_USESHOWWINDOW     As Long = &H1
'Private Const STARTF_USESTDHANDLES     As Long = &H100
Private Const SW_HIDE                  As Integer = 0

'Private Type SECURITY_ATTRIBUTES
'    nLength As Long
'    lpSecurityDescriptor As Long
'    bInheritHandle As Long
'End Type

'Private Type STARTUPINFO
'    cb As Long
'    lpReserved As Long
'    lpDesktop As Long
'    lpTitle As Long
'    dwX As Long
'    dwY As Long
'    dwXSize As Long
'    dwYSize As Long
'    dwXCountChars As Long
'    dwYCountChars As Long
'    dwFillAttribute As Long
'    dwFlags As Long
'    wShowWindow As Integer
'    cbReserved2 As Integer
'    lpReserved2 As Long
'    hStdInput As Long
'    hStdOutput As Long
'    hStdError As Long
'End Type

'Private Type PROCESS_INFORMATION
'    hProcess As Long
'    hThread As Long
'    dwProcessID As Long
'    dwThreadID As Long
'End Type

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

'Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
'Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, lType As Long, lLenOutBuf As Long, lLenInBuf As Long, lMaxInstances As Long) As Long
   
'Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'--------------------------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------------------------------------
' Utilizzati nella funzione DosClose
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
' Constants used by the API functions
Const WM_CLOSE = &H10
Const INFINITE = &HFFFFFFFF
'--------------------------------------------------------------------------------------------------------------------------------

'This property set and get the DOS command line
'It's possible to set this property directly from the
'parameter of the ExecuteCommand method
Public Property Let CommandLine(DOSCommand As String)
    mCommand = DOSCommand
End Property

Public Property Get CommandLine() As String
    CommandLine = mCommand
End Property

'This property ReadOnly get the complete output after
'a command execution
Public Property Get Outputs()
    Outputs = mOutputs
End Property

Public Function ExecuteCommand(Optional ByVal TipoComandoDos As String = "", Optional ByVal CommandLine As String, Optional ByVal NomeProgrammaDosClose As String = "", Optional ByVal WorkDir As String, Optional ByVal pShowWindow As Boolean = False) As String

    ' If the parameter is not empty update the CommandLine property
    If Len(CommandLine) > 0 Then
        mCommand = CommandLine
    End If
    
    ' If the command line is empty then exit whit a error message
    If Len(mCommand) = 0 Then
        MsgBox "Command Line vuota", vbCritical
        Exit Function
    End If
    
    TipoComandoDos = LCase(TipoComandoDos)
    
    
    Select Case TipoComandoDos
        Case Is = LCase("ExecAndCapture")
            Call ExecAndCapture(CommandLine)
            
        Case Is = LCase("ShellExecuteCapture")
            Call ShellExecuteCapture(CommandLine, pShowWindow, WorkDir)
            
        Case Else
            Call EseguiStandard(CommandLine)
    
    End Select
    
    ' Return the Outputs property with the entire DOS output
    ExecuteCommand = mOutputs

    If NomeProgrammaDosClose <> "" Then
        ' Mi assicuro che il programma sia chiuso, altimenti lo chiudo
        DosClose NomeProgrammaDosClose
    End If
    
End Function

Private Function EseguiStandard(ByVal CommandLine As String, Optional ByVal NomeProgrammaDosClose As String = "") As String
    Dim proc As PROCESS_INFORMATION     'Process info filled by CreateProcessA
    Dim ret As Long                     'long variable for get the return value of the
                                        'API functions
    Dim start As STARTUPINFO            'StartUp Info passed to the CreateProceeeA
                                        'function
    Dim sa As SECURITY_ATTRIBUTES       'Security Attributes passeed to the
                                        'CreateProcessA function
    Dim hReadPipe As Long               'Read Pipe handle created by CreatePipe
    Dim hWritePipe As Long              'Write Pite handle created by CreatePipe
    Dim lngBytesread As Long            'Amount of byte read from the Read Pipe handle
    Dim strBuff As String * 256         'String buffer reading the Pipe

    ' Create the Pipe
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    
    If ret = 0 Then
        'If an error occur during the Pipe creation exit
        MsgBox "CreatePipe failed. Error: " & Err.LastDllError, vbCritical
        Exit Function
    End If
    
    'Launch the command line application
    start.cb = Len(start)
    start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    'set the StdOutput and the StdError output to the same Write Pipe handle
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
    'Execute the command
    ret& = CreateProcessA(0&, mCommand, sa, sa, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
        
    If ret <> 1 Then
        'if the command is not found ....
        MsgBox "Errore nella funzione ExecuteCommand del modulo DOSOutputs." & vbNewLine & "File o comando non trovato", vbCritical
        Exit Function
    End If
    
    'Now We can ... must close the hWritePipe
    ret = CloseHandle(hWritePipe)
    mOutputs = ""
    
    'Read the ReadPipe handle
    Do
        ret = ReadFile(hReadPipe, strBuff, 256, lngBytesread, 0&)
        mOutputs = mOutputs & Left(strBuff, lngBytesread)
        ' Send data to the object via ReceiveOutputs event
        RaiseEvent ReceiveOutputs(Left(strBuff, lngBytesread))
    Loop While ret <> 0
    
    'Close the opened handles
    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)
    
    EseguiStandard = mOutputs

End Function

Private Function ExecAndCapture(ByVal sCommandLine As String, Optional ByVal sStartInFolder As String = vbNullString) As String
    ' Call this sub to execute and capture a console app.
    Const BUFSIZE         As Long = 1024 * 10
    Dim hPipeRead         As Long
    Dim hPipeWrite        As Long
    Dim sa                As SECURITY_ATTRIBUTES
    Dim si                As STARTUPINFO
    Dim pi                As PROCESS_INFORMATION
    Dim baOutput(BUFSIZE) As Byte
    Dim sOutput           As String
    Dim lBytesRead        As Long
    
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1    ' get inheritable pipe handles
    End With 'SA
    
    If CreatePipe(hPipeRead, hPipeWrite, sa, 0) = 0 Then
        Exit Function
    End If

    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE          ' hide the window
        .hStdOutput = hPipeWrite
        .hStdError = hPipeWrite
    End With 'SI
    
    If CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, sStartInFolder, si, pi) Then
        Call CloseHandle(hPipeWrite)
        Call CloseHandle(pi.hThread)
        hPipeWrite = 0
        Do
            DoEvents
            If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then
                Exit Do
            End If
            sOutput = Left$(StrConv(baOutput(), vbUnicode), lBytesRead)
            
            mOutputs = mOutputs & sOutput
            ' Send data to the object via ReceiveOutputs event
            RaiseEvent ReceiveOutputs(sOutput)
            
        Loop
        Call CloseHandle(pi.hProcess)
    End If
    
    ' To make sure...
    Call CloseHandle(hPipeRead)
    Call CloseHandle(hPipeWrite)
    
    ExecAndCapture = mOutputs
    
End Function

Private Function ShellExecuteCapture(ByVal sCommandLine As String, Optional ByVal bShowWindow As Boolean = False, Optional ByVal WorkDir As String) As String
    'Purpose     :  Synchronously runs a DOS command line and returns the captured screen output.
    'Inputs      :  sCommandLine                The DOS command line to run.
    '               [bShowWindow]               If True displays the DOS output window.
    'Outputs     :  Returns the screen output
    'Notes       :  This routine will work only with those program that send their output to
    '               the standard output device (stdout).
    '               Windows NT ONLY.
    'WorkDir     :  la directory di lavoro della finestra DOS aperta
        
    Const clReadBytes As Long = 256, INFINITE As Long = &HFFFFFFFF
    Const STARTF_USESHOWWINDOW = &H1, STARTF_USESTDHANDLES = &H100&
    Const SW_HIDE = 0, SW_NORMAL = 1
    Const NORMAL_PRIORITY_CLASS = &H20&
    
    Const PIPE_CLIENT_END = &H0     'The handle refers to the client end of a named pipe instance. This is the default.
    Const PIPE_SERVER_END = &H1     'The handle refers to the server end of a named pipe instance. If this value is not specified, the handle refers to the client end of a named pipe instance.
    Const PIPE_TYPE_BYTE = &H0      'The named pipe is a byte pipe. This is the default.
    Const PIPE_TYPE_MESSAGE = &H4   'The named pipe is a message pipe. If this value is not specified, the pipe is a byte pipe
    
    Dim tProcInfo As PROCESS_INFORMATION, lRetVal As Long, lSuccess As Long
    Dim tStartupInf As STARTUPINFO
    Dim tSecurAttrib As SECURITY_ATTRIBUTES, lhwndReadPipe As Long, lhwndWritePipe As Long
    Dim lBytesRead As Long, sBuffer As String
    Dim lPipeOutLen As Long, lPipeInLen As Long, lMaxInst As Long
    
    Dim AppPath As String
        
    On Error GoTo Errore
    
    If Dir(WorkDir, vbDirectory) <> "" Then
        AppPath = App.Path
        ' Ottengo la lettera del drive e la imposto come drive di lavoro
        ChDrive Left$(WorkDir, 1)
        ' Imposto la directory di lavoro
        ChDir WorkDir
    End If
    
    tSecurAttrib.nLength = Len(tSecurAttrib)
    tSecurAttrib.bInheritHandle = 1&
    tSecurAttrib.lpSecurityDescriptor = 0&

    lRetVal = CreatePipe(lhwndReadPipe, lhwndWritePipe, tSecurAttrib, 0)
    If lRetVal = 0 Then
        'CreatePipe failed
        Exit Function
    End If

    tStartupInf.cb = Len(tStartupInf)
    tStartupInf.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    tStartupInf.hStdOutput = lhwndWritePipe
    If bShowWindow Then
        'Show the DOS window
        tStartupInf.wShowWindow = SW_NORMAL
    Else
        'Hide the DOS window
        tStartupInf.wShowWindow = SW_HIDE
    End If

    lRetVal = CreateProcessA(0&, sCommandLine, tSecurAttrib, tSecurAttrib, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, tStartupInf, tProcInfo)
    If lRetVal <> 1 Then
        'CreateProcess failed
        Exit Function
    End If
    
    'Process created, wait for completion. Note, this will cause your application
    'to hang indefinately until this process completes.
    ' Questa dà errore con la nuova versione di GPSBabel
    'WaitForSingleObject tProcInfo.hProcess, INFINITE
    ' così funziona...........
    WaitForSingleObject tProcInfo.hProcess, 1000
    
    'Determine pipes contents
    lSuccess = GetNamedPipeInfo(lhwndReadPipe, PIPE_TYPE_BYTE, lPipeOutLen, lPipeInLen, lMaxInst)
    If lSuccess Then
        'Got pipe info, create buffer
        sBuffer = String(lPipeOutLen, 0)
        'Read Output Pipe
        lSuccess = ReadFile(lhwndReadPipe, sBuffer, lPipeOutLen, lBytesRead, 0&)
        
        If lSuccess = 1 Then
            'Pipe read successfully
            ShellExecuteCapture = Left$(sBuffer, lBytesRead)
            
            mOutputs = Left$(sBuffer, lBytesRead)
            ' Send data to the object via ReceiveOutputs event
            RaiseEvent ReceiveOutputs(Left$(sBuffer, lBytesRead))

        End If
    End If
    
    'Close handles
    Call CloseHandle(tProcInfo.hProcess)
    Call CloseHandle(tProcInfo.hThread)
    Call CloseHandle(lhwndReadPipe)
    Call CloseHandle(lhwndWritePipe)
    
    If WorkDir <> "" Then
        ' Reimposto la lettera del drive come drive di lavoro
        ChDrive Left$(AppPath, 1)
        ' Reimposto la directory di lavoro
        ChDir AppPath
    End If
    
    ShellExecuteCapture = mOutputs
    
    Exit Function
    
Errore:
    ShellExecuteCapture = ""
    MsgBox ("Errore nella funzione: ShellExecuteCapture (" & Err.Number & ") " & Err.Description)

End Function

Public Function DosClose(ByVal NomeProgramma As String) As Boolean
    ' Cerca il nome del programma e se lo trova lo chiude
    Dim hWindow As Long
    Dim lngResult As Long
    Dim lngReturnValue As Long

    hWindow = FindWindow(vbNullString, NomeProgramma)
    lngReturnValue = PostMessage(hWindow, WM_CLOSE, vbNull, vbNull)
    lngResult = WaitForSingleObject(hWindow, INFINITE)
    DoEvents
        
        hWindow = FindWindow(vbNullString, NomeProgramma)
        If IsWindow(hWindow) = 1 Then
           ' The handle still exists.
           ' Use the TerminateProcess function to close all related processes to this handle.
           DosClose = False
        Else
           'Handle does not exist.
           DosClose = True
        End If

End Function

