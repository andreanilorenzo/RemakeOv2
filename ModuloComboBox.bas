Attribute VB_Name = "ComboBoxModule"
Option Explicit
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

'---------------------------------------------------------------------------------
' Per la funzione AutoSizeLBHeight
' Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
'---------------------------------------------------------------------------------

Private Const LB_GETITEMHEIGHT = &H1A1
Public defWinProc As Long

Public Const GWL_WNDPROC As Long = -4
Private Const CBN_DROPDOWN As Long = 7
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_KEYDOWN As Long = &H100
Private Const VK_F4 As Long = &H73

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Sub Unhook(hwnd As Long)
    
   If defWinProc <> 0 Then
      Call SetWindowLong(hwnd, GWL_WNDPROC, defWinProc)
      defWinProc = 0
   End If
    
End Sub

Public Sub Hook(hwnd As Long)

   'Don't hook twice or you will be unable to unhook it
    If defWinProc = 0 Then
      defWinProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    End If
    
End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  'only if the window is the combo box...
   If hwnd = frmDownload.cmbTipoPDI.hwnd Then
   
      Select Case uMsg
      
        'SEE COMMENTS SECTION RE: THIS NOTIFICATION
    'Case CBN_DROPDOWN  'the list box of a combo
        '                   'box is about to be made visible.
        '
        '  'return 1 to indicate we ate the message
        '   WindowProc = 1
   
         Case WM_KEYDOWN   'prevent the F4 key from showing
                           'the combo's list
            
            If wParam = VK_F4 Then
              'set up the parameters as though a
              'mouse click occurred on the combo,
              'and call this routine again
               Call WindowProc(hwnd, WM_LBUTTONDOWN, 1, 1000)
            Else

              'there's nothing to do keyboard-wise
              'with the combo, so return 1 to
              'indicate we ate the message
               WindowProc = 1
            End If
            
         Case WM_LBUTTONDOWN  'process mouse clicks if the list is hidden, position and show it
            If frmDownload.List1.Visible = False Then
               With frmDownload
                  .List1.Left = .cmbTipoPDI.Left
                  ' Imposto la larghezza della ListBox
                  .List1.Width = .cmbTipoPDI.Width
                  .List1.Top = .cmbTipoPDI.Top + .cmbTipoPDI.Height + 1
                  .List1.Visible = True
                  .List1.ZOrder 0
                  .List1.SetFocus
               End With
               
            Else
              'the list must be visible, so hide it
               frmDownload.List1.Visible = False
            End If
           'return 1 to indicate we processed the message
            WindowProc = 1
         
         Case Else
           'call the default window handler
            WindowProc = CallWindowProc(defWinProc, hwnd, uMsg, wParam, lParam)
      End Select
   End If
   
End Function

Public Function AutoSizeLBHeight(LB As Object) As Boolean
    'PURPOSE: Will automatically set the height of a
    'list box based on the number and height of entries
    
    'PARAMETERS: LB = the ListBox control to autosize
    
    'RETURNS: True if successful, false otherwise
    
    'NOTE: LB's parent's (e.g., form, picturebox)
    'scalemode must be vbTwips, which is the
    'default
    If Not TypeOf LB Is ListBox Then Exit Function
    
    On Error GoTo ErrHandler
    
    Dim lItemHeight As Long
    Dim lRet As Long
    Dim lItems As Long
    Dim sngTwips As Single
    Dim sngLBHeight As Single
    
    If LB.ListCount = 0 Then
        LB.Height = 125
        AutoSizeLBHeight = True
        
    Else
        lItems = LB.ListCount
        
        lItemHeight = SendMessage(LB.hwnd, LB_GETITEMHEIGHT, 0&, 0&)
        If lItemHeight > 0 Then
            sngTwips = lItemHeight * Screen.TwipsPerPixelY
            sngLBHeight = (sngTwips * lItems) + 125
            LB.Height = sngLBHeight
            AutoSizeLBHeight = True
        End If
    End If
    
ErrHandler:
End Function
