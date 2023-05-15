VERSION 5.00
Begin VB.Form frmSysTray 
   Caption         =   "Sys Tray Interface"
   ClientHeight    =   1050
   ClientLeft      =   5610
   ClientTop       =   3360
   ClientWidth     =   3045
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   3045
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   360
      Picture         =   "frmSysTray.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   240
      Width           =   540
   End
   Begin VB.Timer TmrFlash 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   240
   End
   Begin VB.PictureBox Flash1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   960
      Picture         =   "frmSysTray.frx":0614
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox Flash2 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1320
      Picture         =   "frmSysTray.frx":2DB6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   240
      Width           =   300
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "&Popup"
      Begin VB.Menu mnuSysTray 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 03/03/2003
' * Added Unicode support
' * Added support for new tray version (ME,2000 or above required)
' * Added support for balloon tips (ME,2000 or above required)

' frmSysTray.
' Steve McMahon
' Original version based on code supplied from Ben Baird:

'Author:
'        Ben Baird <psyborg@cyberhighway.com>
'        Copyright (c) 1997, Ben Baird
'
'Purpose:
'        Demonstrates setting an icon in the taskbar's
'        system tray without the overhead of subclassing
'        to receive events.

Private Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATAA) As Long
Private Declare Function Shell_NotifyIconW Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NOTIFYICON_VERSION = 3

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Type NOTIFYICONDATAA
   cbSize As Long             ' 4
   hwnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip As String * 128      ' 152
   dwState As Long            ' 156
   dwStateMask As Long        ' 160
   szInfo As String * 256     ' 416
   uTimeOutOrVersion As Long  ' 420
   szInfoTitle As String * 64 ' 484
   dwInfoFlags As Long        ' 488
   guidItem As Long           ' 492
End Type

Private Type NOTIFYICONDATAW
   cbSize As Long             ' 4
   hwnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip(0 To 255) As Byte    ' 280
   dwState As Long            ' 284
   dwStateMask As Long        ' 288
   szInfo(0 To 511) As Byte   ' 800
   uTimeOutOrVersion As Long  ' 804
   szInfoTitle(0 To 127) As Byte ' 932
   dwInfoFlags As Long        ' 936
   guidItem As Long           ' 940
End Type


Private nfIconDataA As NOTIFYICONDATAA
Private nfIconDataW As NOTIFYICONDATAW
Private nID As NOTIFYICONDATA

Private Const NOTIFYICONDATAA_V1_SIZE_A = 88
Private Const NOTIFYICONDATAA_V1_SIZE_U = 152
Private Const NOTIFYICONDATAA_V2_SIZE_A = 488
Private Const NOTIFYICONDATAA_V2_SIZE_U = 936

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Const WM_USER = &H400

Private Const NIN_SELECT = WM_USER
Private Const NINF_KEY = &H1
Private Const NIN_KEYSELECT = (NIN_SELECT Or NINF_KEY)
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

' Version detection:
Private Declare Function GetVersion Lib "kernel32" () As Long

Public Event SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseMove()
Public Event SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
Public Event MenuClick(ByVal lIndex As Long, ByVal sKey As String)
Public Event BalloonShow()
Public Event BalloonHide()
Public Event BalloonTimeOut()
Public Event BalloonClicked()

Public Enum EBalloonIconTypes
   NIIF_NONE = 0
   NIIF_INFO = 1
   NIIF_WARNING = 2
   NIIF_ERROR = 3
   NIIF_NOSOUND = &H10
End Enum

Private LastWindowState As Integer

Private m_bAddedMenuItem As Boolean
Private m_iDefaultIndex As Long

Private m_bUseUnicode As Boolean
Private m_bSupportsNewVersion As Boolean

Public Sub MeQueryUnload(ByRef Form As Form, Cancel As Integer, UnloadMode As Integer)

   If UnloadMode = vbFormControlMenu Then
      ' Cancel by setting Cancel = 1, minimize and hide main window.
      Cancel = 1
      Form.WindowState = vbMinimized
      Form.Hide
   End If
   
End Sub

Public Sub RestoreForm(ByRef Form As Form)

   ' Don't "restore"  FSys is visible and not minimized.
   If (Form.Visible And Form.WindowState <> vbMinimized) Then Exit Sub
   ' Restore LastWindowState
   Form.WindowState = LastWindowState
   Form.Visible = True
   SetForegroundWindow Form.hwnd
   
End Sub

Public Sub Minimize(ByRef Form As Form)
   Form.WindowState = vbMinimized
End Sub

Public Sub MeResize(ByRef Form As Form)

   Select Case Form.WindowState
      Case vbNormal, vbMaximized
         ' Store LastWindowState
         LastWindowState = Form.WindowState
         Form.WindowState = vbNormal
      Case vbMinimized
         Form.Hide
   End Select
   
End Sub

Public Sub ShowBalloonTip(ByVal sMessage As String, Optional ByVal sTitle As String, Optional ByVal eIcon As EBalloonIconTypes, Optional ByVal lTimeOutMs = 30000)
    Dim lR As Long
   
   If (m_bSupportsNewVersion) Then
      If (m_bUseUnicode) Then
         stringToArray sMessage, nfIconDataW.szInfo, 512
         stringToArray sTitle, nfIconDataW.szInfoTitle, 128
         nfIconDataW.uTimeOutOrVersion = lTimeOutMs
         nfIconDataW.dwInfoFlags = eIcon
         nfIconDataW.uFlags = NIF_INFO
         lR = Shell_NotifyIconW(NIM_MODIFY, nfIconDataW)
      Else
         nfIconDataA.szInfo = sMessage
         nfIconDataA.szInfoTitle = sTitle
         nfIconDataA.uTimeOutOrVersion = lTimeOutMs
         nfIconDataA.dwInfoFlags = eIcon
         nfIconDataA.uFlags = NIF_INFO
         lR = Shell_NotifyIconA(NIM_MODIFY, nfIconDataA)
      End If
   Else
      ' can't do it, fail silently.
   End If
   
End Sub

Public Property Get ToolTip() As String
Dim sTip As String
Dim iPos As Long
    sTip = nfIconDataA.szTip
    iPos = InStr(sTip, Chr$(0))
    If (iPos <> 0) Then
        sTip = Left$(sTip, iPos - 1)
    End If
    ToolTip = sTip
End Property

Public Property Let ToolTip(ByVal sTip As String)
   If (m_bUseUnicode) Then
      stringToArray sTip, nfIconDataW.szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
      nfIconDataW.uFlags = NIF_TIP
      Shell_NotifyIconW NIM_MODIFY, nfIconDataW
   Else
      If (sTip & Chr$(0) <> nfIconDataA.szTip) Then
         nfIconDataA.szTip = sTip & Chr$(0)
         nfIconDataA.uFlags = NIF_TIP
         Shell_NotifyIconA NIM_MODIFY, nfIconDataA
      End If
   End If
End Property

Public Property Get IconHandle() As Long
    IconHandle = nfIconDataA.hIcon
End Property

Public Property Let IconHandle(ByVal hIcon As Long)
   
   If (m_bUseUnicode) Then
      If (hIcon <> nfIconDataW.hIcon) Then
         nfIconDataW.hIcon = hIcon
         nfIconDataW.uFlags = NIF_ICON
         Shell_NotifyIconW NIM_MODIFY, nfIconDataW
      End If
   Else
      If (hIcon <> nfIconDataA.hIcon) Then
         nfIconDataA.hIcon = hIcon
         nfIconDataA.uFlags = NIF_ICON
         Shell_NotifyIconA NIM_MODIFY, nfIconDataA
      End If
   End If
   
End Property

Public Function AddMenuItem(ByVal sCaption As String, Optional ByVal sKey As String = "", Optional ByVal bDefault As Boolean = False) As Long
    Dim iIndex As Long
    
    If Not (m_bAddedMenuItem) Then
        iIndex = 0
        m_bAddedMenuItem = True
    Else
        iIndex = mnuSysTray.UBound + 1
        Load mnuSysTray(iIndex)
    End If
    
    mnuSysTray(iIndex).Visible = True
    mnuSysTray(iIndex).Tag = sKey
    mnuSysTray(iIndex).Caption = sCaption
    
    If (bDefault) Then
        m_iDefaultIndex = iIndex
    End If
    
    AddMenuItem = iIndex
    
End Function

Private Function ValidIndex(ByVal lIndex As Long) As Boolean
    ValidIndex = (lIndex >= mnuSysTray.LBound And lIndex <= mnuSysTray.UBound)
End Function

Public Sub EnableMenuItem(ByVal lIndex As Long, ByVal bState As Boolean)
    If (ValidIndex(lIndex)) Then
        mnuSysTray(lIndex).Enabled = bState
    End If
End Sub

Public Function RemoveMenuItem(ByVal iIndex As Long) As Long
    Dim I As Long
    
   If ValidIndex(iIndex) Then
      If (iIndex = 0) Then
         mnuSysTray(0).Caption = ""
      Else
         ' remove the item:
         For I = iIndex + 1 To mnuSysTray.UBound
            mnuSysTray(iIndex - 1).Caption = mnuSysTray(iIndex).Caption
            mnuSysTray(iIndex - 1).Tag = mnuSysTray(iIndex).Tag
         Next I
         Unload mnuSysTray(mnuSysTray.UBound)
      End If
   End If
   
End Function

Public Property Get DefaultMenuIndex() As Long
   DefaultMenuIndex = m_iDefaultIndex
End Property

Public Property Let DefaultMenuIndex(ByVal lIndex As Long)
   If (ValidIndex(lIndex)) Then
      m_iDefaultIndex = lIndex
   Else
      m_iDefaultIndex = 0
   End If
End Property

Public Function ShowMenu()
   SetForegroundWindow Me.hwnd
   If (m_iDefaultIndex > -1) Then
      Me.PopupMenu mnuPopup, 0, , , mnuSysTray(m_iDefaultIndex)
   Else
      Me.PopupMenu mnuPopup, 0
   End If
End Function

Private Sub Form_Load()
   ' Get version:
   Dim lMajor As Long
   Dim lMinor As Long
   Dim bIsNt As Long
   GetWindowsVersion lMajor, lMinor, , , bIsNt

   If (bIsNt) Then
      m_bUseUnicode = True
      If (lMajor >= 5) Then
         ' 2000 or XP
         m_bSupportsNewVersion = True
      End If
   ElseIf (lMajor = 4) And (lMinor = 90) Then
      ' Windows ME
      m_bSupportsNewVersion = True
   End If
   
   
   'Add the icon to the system tray...
   Dim lR As Long
   
   If (m_bUseUnicode) Then
      With nfIconDataW
         .hwnd = Me.hwnd
         .uID = Me.Icon
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon.handle
         stringToArray App.FileDescription, .szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
         If (m_bSupportsNewVersion) Then
            .uTimeOutOrVersion = NOTIFYICON_VERSION
         End If
         .cbSize = nfStructureSize
      End With
      lR = Shell_NotifyIconW(NIM_ADD, nfIconDataW)
      If (m_bSupportsNewVersion) Then
         Shell_NotifyIconW NIM_SETVERSION, nfIconDataW
      End If
   Else
      With nfIconDataA
         .hwnd = Me.hwnd
         .uID = Me.Icon
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon.handle
         .szTip = App.FileDescription & Chr$(0)
         If (m_bSupportsNewVersion) Then
            .uTimeOutOrVersion = NOTIFYICON_VERSION
         End If
         .cbSize = nfStructureSize
      End With
      lR = Shell_NotifyIconA(NIM_ADD, nfIconDataA)
      If (m_bSupportsNewVersion) Then
         lR = Shell_NotifyIconA(NIM_SETVERSION, nfIconDataA)
      End If
   End If
   
End Sub

Private Sub stringToArray(ByVal sString As String, bArray() As Byte, ByVal lMaxSize As Long)
    Dim b() As Byte
    Dim I As Long
    Dim j As Long
    
    If Len(sString) > 0 Then
       b = sString
       For I = LBound(b) To UBound(b)
          bArray(I) = b(I)
          If (I = (lMaxSize - 2)) Then
             Exit For
          End If
       Next I
       For j = I To lMaxSize - 1
          bArray(j) = 0
       Next j
    End If
   
End Sub

Private Function unicodeSize(ByVal lSize As Long) As Long

   If (m_bUseUnicode) Then
      unicodeSize = lSize * 2
   Else
      unicodeSize = lSize
   End If
   
End Function

Private Property Get nfStructureSize() As Long
   If (m_bSupportsNewVersion) Then
      If (m_bUseUnicode) Then
         nfStructureSize = NOTIFYICONDATAA_V2_SIZE_U
      Else
         nfStructureSize = NOTIFYICONDATAA_V2_SIZE_A
      End If
   Else
      If (m_bUseUnicode) Then
         nfStructureSize = NOTIFYICONDATAA_V1_SIZE_U
      Else
         nfStructureSize = NOTIFYICONDATAA_V1_SIZE_A
      End If
   End If
End Property

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lX As Long
   ' VB manipulates the x value according to scale mode:
   ' we must remove this before we can interpret the
   ' message windows was trying to send to us:
   lX = ScaleX(X, Me.ScaleMode, vbPixels)
   Select Case lX
   Case WM_MOUSEMOVE
      RaiseEvent SysTrayMouseMove
   Case WM_LBUTTONUP
      RaiseEvent SysTrayMouseDown(vbLeftButton)
   Case WM_LBUTTONUP
      RaiseEvent SysTrayMouseUp(vbLeftButton)
   Case WM_LBUTTONDBLCLK
      RaiseEvent SysTrayDoubleClick(vbLeftButton)
   Case WM_RBUTTONDOWN
      RaiseEvent SysTrayMouseDown(vbRightButton)
   Case WM_RBUTTONUP
      RaiseEvent SysTrayMouseUp(vbRightButton)
   Case WM_RBUTTONDBLCLK
      RaiseEvent SysTrayDoubleClick(vbRightButton)
   Case NIN_BALLOONSHOW
      RaiseEvent BalloonShow
   Case NIN_BALLOONHIDE
      RaiseEvent BalloonHide
   Case NIN_BALLOONTIMEOUT
      RaiseEvent BalloonTimeOut
   Case NIN_BALLOONUSERCLICK
      RaiseEvent BalloonClicked
   End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If (m_bUseUnicode) Then
      Shell_NotifyIconW NIM_DELETE, nfIconDataW
   Else
      Shell_NotifyIconA NIM_DELETE, nfIconDataA
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Remove
    DoEvents
End Sub

Private Sub mnuSysTray_Click(Index As Integer)
   RaiseEvent MenuClick(Index, mnuSysTray(Index).Tag)
End Sub

Private Sub GetWindowsVersion(Optional ByRef lMajor = 0, Optional ByRef lMinor = 0, Optional ByRef lRevision = 0, Optional ByRef lBuildNumber = 0, Optional ByRef bIsNt = False)
    Dim lR As Long
   
   lR = GetVersion()
   lBuildNumber = (lR And &H7F000000) \ &H1000000
   If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
   lRevision = (lR And &HFF0000) \ &H10000
   lMinor = (lR And &HFF00&) \ &H100
   lMajor = (lR And &HFF)
   bIsNt = ((lR And &H80000000) = 0)
   
End Sub

Public Sub Add()
    'Add the icon from the system tray
   
   If (m_bUseUnicode) Then
         nfIconDataW.uFlags = NIF_ICON
         Shell_NotifyIconW NIM_ADD, nfIconDataW
   Else
         nfIconDataA.uFlags = NIF_ICON
         Shell_NotifyIconA NIM_ADD, nfIconDataA
   End If

End Sub

Public Sub Remove()
    'Remove the icon from the system tray
   
   If (m_bUseUnicode) Then
         nfIconDataW.uFlags = NIF_ICON
         Shell_NotifyIconW NIM_DELETE, nfIconDataW
   Else
         nfIconDataA.uFlags = NIF_ICON
         Shell_NotifyIconA NIM_DELETE, nfIconDataA
   End If

End Sub

Private Sub TmrFlash_Timer()
    Static LastIconWasFlash1 As Boolean
   
    LastIconWasFlash1 = Not LastIconWasFlash1
    
    Select Case LastIconWasFlash1
       Case True
          Me.Icon = Flash2
       Case Else
          Me.Icon = Flash1
    End Select
    
    UpdateIcon NIM_MODIFY
   
End Sub

Public Sub FlashIcona(Optional Attiva As Boolean = True)
    
    If Attiva = True Then
        TmrFlash.Enabled = True
    Else
        TmrFlash.Enabled = False
    End If

End Sub

Private Sub UpdateIcon(Value As Long)
   ' Used to add, modify and delete icon.
   With nID
      .cbSize = Len(nID)
      .hwnd = Me.hwnd
      .uID = vbNull
      .uFlags = NIM_DELETE Or NIF_TIP Or NIM_MODIFY
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
   End With
   Shell_NotifyIcon Value, nID
   
End Sub

