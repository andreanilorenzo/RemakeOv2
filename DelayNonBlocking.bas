Attribute VB_Name = "DelayNonBlocking"
Option Explicit
'********************************************
'*    (c) 1999-2000 Sergey Merzlikin        *
'********************************************

' Rather often Visual Basic programmers need to use Windows API functions which delay program
' execution until the occurence of certain event.
'
' Here is the list of these functions:
' Sleep, SleepEx, WaitForSingleObject, WaitForSingleObjectEx,
' WaitForMultipleObjects, WaitForMultipleObjectsEx, MsgWaitForMultipleObjects,
' MsgWaitForMultipleObjectsEx, SignalObjectAndWait.
'
' So, what problems may encounter the programmer using above API functions with?
'
' Fortunately, there is only one problem, but it is rather serious.
' The thing is that the programs written on Visual Basic with a small exception
' are executing in one OS thread. It means that when one of wait functions starts,
' the "life" of the program completely stops: the visual interface freezes,
' the buttons became unclickable, and TaskManager reports "Not Responding".
' More complete discussion of this problem can be found at http://smsoft.chat.ru/en/vbwait.htm
'
' The offered MsgWaitObj function may be used as a non-blocking equivalent of Sleep,
' WaitForSingleObject and WaitForMultipleObjects functions.

'------------------------------------------------------------------------------------------------
'Per la funzione MsgWaitObj
Private Const STATUS_TIMEOUT = &H102&
Private Const INFINITE1 = -1& ' Infinite interval
Private Const QS_KEY = &H1&
Private Const QS_MOUSEMOVE = &H2&
Private Const QS_MOUSEBUTTON = &H4&
Private Const QS_POSTMESSAGE = &H8&
Private Const QS_TIMER = &H10&
Private Const QS_PAINT = &H20&
Private Const QS_SENDMESSAGE = &H40&
Private Const QS_HOTKEY = &H80&
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'------------------------------------------------------------------------------------------------
'Per la funzione eseguiEattendi
'Private Const INFINITE = &HFFFF
Private Const STARTF_USESHOWWINDOW = &H1

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Long
    cbReserved2 As Long
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum

Private Const CREATE_NEW_PROCESS_GROUP = &H200

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'------------------------------------------------------------------------------------------------

Public Function MsgWaitObj(Interval As Long, Optional hObj As Long = 0&, Optional nObj As Long = 0&) As Long

' The MsgWaitObj function replaces Sleep,
' WaitForSingleObject, WaitForMultipleObjects functions.
'
' Unlike these functions, it
' doesn't block thread messages processing.
'
' Using instead Sleep:
'     MsgWaitObj dwMilliseconds
'
' Using instead WaitForSingleObject:
'     retval = MsgWaitObj(dwMilliseconds, hObj, 1&)
'
' Using instead WaitForMultipleObjects:
'     retval = MsgWaitObj(dwMilliseconds, hObj(0&), n),
'     where n - wait objects quantity,
'     hObj() - their handles array.

Dim t As Long, T1 As Long
If Interval <> INFINITE1 Then
    t = GetTickCount()
    On Error Resume Next
    t = t + Interval
    ' Overflow prevention
    If Err <> 0& Then
        If t > 0& Then
            t = ((t + &H80000000) + Interval) + &H80000000
        Else
            t = ((t - &H80000000) + Interval) - &H80000000
        End If
    End If
    On Error GoTo 0
    ' T contains now absolute time of the end of interval
Else
    T1 = INFINITE1
End If
Do
    If Interval <> INFINITE1 Then
        T1 = GetTickCount()
        On Error Resume Next
     T1 = t - T1
        ' Overflow prevention
        If Err <> 0& Then
            If t > 0& Then
                T1 = ((t + &H80000000) - (T1 - &H80000000))
            Else
                T1 = ((t - &H80000000) - (T1 + &H80000000))
            End If
        End If
        On Error GoTo 0
        ' T1 contains now the remaining interval part
        If IIf((T1 Xor Interval) > 0&, _
            T1 > Interval, T1 < 0&) Then
            ' Interval expired
            ' during DoEvents
            MsgWaitObj = STATUS_TIMEOUT
            Exit Function
        End If
    End If
    ' Wait for event, interval expiration
    ' or message appearance in thread queue
    MsgWaitObj = MsgWaitForMultipleObjects(nObj, hObj, 0&, T1, QS_ALLINPUT)
    ' Let's message be processed
    DoEvents
    If MsgWaitObj <> nObj Then Exit Function
    ' It was message - continue to wait
Loop
End Function

Public Function eseguiEattendi(ByVal ExeFullPath As String, Optional ByVal sParametri As String = vbNullString, Optional ByVal sDirectoryLavoro As String = vbNullString, Optional ByVal sWindowStyle As Integer = 6, Optional ByVal TimeOutValue As Long = 0) As Boolean
' Apre un programma ed attende che questo venga chiuso per continuare la routine
'
' ExeFullPath: Posizione e nome del programma
'
' sParametri: la riga di comando da passare al programma
'
' WindowStyle:
' vbHide 0 La finestra è nascosta e lo stato attivo viene passato alla finestra nascosta.
' vbNormalFocus 1 La finestra è attivata e vengono ripristinate la dimensione e la posizione originali.
' vbMinimizedFocus 2 La finestra è ridotta a icona e attivata.
' vbMaximizedFocus 3 La finestra è ingrandita e attivata.
' vbNormalNoFocus 4 Vengono ripristinate le dimensioni e posizione precedenti della finestra. La finestra attiva resta attiva.
' vbMinimizedNoFocus 6 La finestra è ridotta a icona. La finestra attiva resta attiva.
'
' TimeOutValue: Tempo massimo da attendere oltre il quale esco dalla routine e resituisco Falso

    Dim pclass As Long
    Dim sInfo As STARTUPINFO
    Dim pinfo As PROCESS_INFORMATION
    Dim sec1 As SECURITY_ATTRIBUTES
    Dim sec2 As SECURITY_ATTRIBUTES
    Dim dProc As Double
    Dim hProc As Long
    Dim returnvalue As Long
    Dim lStart As Long
    Dim lInst As Long
    Dim lTimeToQuit As Long
    Dim bPastMidnight As Boolean

    On Error GoTo ErrorHandler

    lStart = CLng(Timer)
    
    'Deal with timeout being reset at Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If

    'Set the structure size
    sec1.nLength = Len(sec1)
    sec2.nLength = Len(sec2)
    
    sInfo.cb = Len(sInfo)
    'Set the flags
    sInfo.dwFlags = STARTF_USESHOWWINDOW
    'Set the window's startup position
    sInfo.wShowWindow = sWindowStyle
    'Set the priority class
    pclass = NORMAL_PRIORITY_CLASS
    
    ' Lancio il programma
    If CreateProcess(ExeFullPath, sParametri, sec1, sec2, False, pclass, ByVal 0&, sDirectoryLavoro, sInfo, pinfo) Then
        'Aspetto finchè non è stato chiuso
        Do
            returnvalue = WaitForSingleObject(pinfo.hProcess, 0)
            DoEvents
            If TimeOutValue And Timer > lTimeToQuit Then
                If bPastMidnight Then
                    If Timer < lStart Then Exit Do
                Else
                    Exit Do
                End If
            End If
            ' Inserisco una pausa di 15 millisecondi per allegerire ulteriormente il processo
            MsgWaitObj 15
        Loop Until returnvalue <> 258
    End If

    eseguiEattendi = True
    Exit Function
    
ErrorHandler:
    eseguiEattendi = False
    Exit Function
    
End Function
