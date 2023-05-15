Attribute VB_Name = "FormMinMaxInfo"
  Option Explicit
  ' A demo project showing how to prevent the user from making a window smaller
  ' or larger than you want them to, through subclassing the WM_GETMINMAXINFO message.
  ' by Bryan Stafford of New Vision Software® - newvision@mvps.org
  ' this demo is released into the public domain "as is" without
  ' warranty or guaranty of any kind.  In other words, use at
  ' your own risk.
  
  ' See the comments at the end of this module for a brief explaination of
  ' what subclassing is.
  
  'Type POINTAPI
  '  x As Long
  '  y As Long
  'End Type

  ' the message we will subclass
  Public Const WM_GETMINMAXINFO As Long = &H24&

  Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
  End Type

  ' this var will hold a pointer to the original message handler so we MUST
  ' save it so that it can be restored before we exit the app.  if we don't
  ' restore it.... CRASH!!!!
  Public g_nProcOld As Long
  
  ' declarations of the API functions used
  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
  Public Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
  Public Const GWL_WNDPROC As Long = (-4&)
  
  ' API call to alter the class data for a window
  Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&) As Long
  
'WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!!
'
' Do NOT try to step through this function in debug mode!!!!
' You WILL crash!!!  Also, do NOT set any break points in this function!!!
' You WILL crash!!!  Subclassing is non-trivial and should be handled with
' EXTREME care!!!
'
' There are ways to use a "Debug" dll to allow you to set breakpoints in
' subclassed code in the IDE but this was not implimented for this demo.
'
'WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!!
  
Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  ' this is *our* implimentation of the message handling routine
  
    Dim Min As Long: Min = 8
    Dim Max As Long: Max = 4
  
  ' determine which message was recieved
  Select Case iMsg
    
    Case WM_GETMINMAXINFO
      ' dimention a variable to hold the structure passed from Windows in lParam
      Dim udtMINMAXINFO As MINMAXINFO
      Dim nWidthPixels&, nHeightPixels&
      
      nWidthPixels = Screen.Width \ Screen.TwipsPerPixelX
      nHeightPixels = Screen.Height \ Screen.TwipsPerPixelY
      
      ' copy the struct to our UDT variable
      CopyMemory udtMINMAXINFO, ByVal lParam, Len(udtMINMAXINFO)
           
      With udtMINMAXINFO
        ' set the width of the form when it's maximized
        .ptMaxSize.X = nWidthPixels - (nWidthPixels \ Max)
        ' set the height of the form when it's maximized
        .ptMaxSize.Y = nHeightPixels - (nHeightPixels \ Max)
        
        ' set the Left of the form when it's maximized
        .ptMaxPosition.X = nWidthPixels \ Min
        ' set the Top of the form when it's maximized
        .ptMaxPosition.Y = nHeightPixels \ Min
        
        ' set the max width that the user can drag the form
        .ptMaxTrackSize.X = .ptMaxSize.X
        ' set the max height that the user can drag the form
        .ptMaxTrackSize.Y = .ptMaxSize.Y
        
        ' set the min width that the user can drag the form
        .ptMinTrackSize.X = nWidthPixels \ Max
        ' set the min width that the user can drag the form
        .ptMinTrackSize.Y = nHeightPixels \ Max
      End With
           
      ' copy our modified struct back to the Windows struct
      CopyMemory ByVal lParam, udtMINMAXINFO, Len(udtMINMAXINFO)
  
      ' return zero indicating that we have acted on this message
      WindowProc = 0&
      
      ' exit the function without letting VB get it's grubby little hands on the message
      Exit Function
      
  End Select
  
  ' pass all messages on to VB and then return the value to Windows
  WindowProc = CallWindowProc(g_nProcOld, hWnd, iMsg, wParam, lParam)

End Function

' What is subclassing anyway?
'
' Windows runs on "messages".  A message is a unique value that, when
' recieved by a window or the operating system, tells either that
' something has happened and that an action of some sort needs to be
' taken.  Sort of like your nervous system passing feeling messages
' to your brain and the brain passing movement messages to your body.
'
' So, each window has what is called a message handler.  This is a
' function where all of the messages FROM Windows are recieved.  Every
' window has one.  This means every button, textbox, picturebox, form,
' etc...  Windows keeps track of where the message handler (called a
' WindowProc [short for PROCedure]) in a "Class" structure associated
' with each window handle (otherwise known as hWnd).
'
' What happens when a window is subclassed is that you insert a new
' window procedure in line with the original window procedure.  In other
' words, Windows sends the messages for the given window to YOUR WindowProc
' FIRST where you are responsible for handling any messages you want to
' handle.  Then you pass the remaining messages on to the default
' WindoProc.  So it looks like this:
'
'  Windows Message Sender --> Your WindowProc --> Default WindowProc
'
' A window can be subclassed MANY times so it could look like this:
'
'  Windows Message Sender --> Your WindowProc --> Another WindowProc _
'  --> Yet Another WindowProc --> Default WindowProc
'
' You can also change the order of when you respond to a message by
' where in your routine you pass the message on to the next WindowProc.
' Let's say that you want to draw something on the window AFTER the
' default WindowProc handles the WM_PAINT message.  This is easily done
' by calling the default proc before you do your drawing.   Like so:
'
' Public Function WindowProc(Byval hWnd, Byval etc....)
'
'   Select Case iMsg
'     Case SOME_MESSAGE
'       DoSomeStuff
'
'     Case WM_PAINT
'       ' pass the message to the defproc FIRST
'       Call CallWindowProc(m_g_nProcOld, hWnd, iMsg, wParam, lParam)
'
'       DoDrawingStuff ' <- do your drawing
'
'       WindowProc = nYourReturnVal ' <- retrun the desired value
'                                   '    to the system
'
'       Exit Function ' <- exit since we already passed the
'                     '    measage to the defproc
'
'   End Select
'
'   ' pass all messages on to VB and then return the value to windows
'   WindowProc = CallWindowProc(m_g_nProcOld, hWnd, iMsg, wParam, lParam)
'
' End Function
'
'
' This is just a basic overview of subclassing but I hope it helps if
' you were fuzzy about the subject before reading this.
'

