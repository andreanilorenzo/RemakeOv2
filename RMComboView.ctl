VERSION 5.00
Begin VB.UserControl RMComboView 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   KeyPreview      =   -1  'True
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   243
   Begin VB.PictureBox picImages 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      Picture         =   "RMComboView.ctx":0000
      ScaleHeight     =   495
      ScaleWidth      =   765
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Timer tmrRelease 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   120
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   150
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1785
   End
End
Attribute VB_Name = "RMComboView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#############################################################################################################################
'Title:     ComboView (An Owner-drawn Multi-Column/Checkbox enabled ComboBox replacement)
'Author:    Richard Mewett
'Created:   01/06/05
'Version:   1.0.4 (30th January 2006)

'Copyright © 2005 Richard Mewett. All rights reserved.
'
'This software is provided "as-is," without any express or implied warranty.
'In no event shall the author be held liable for any damages arising from the
'use of this software.
'If you do not agree with these terms, do not install "ComboView". Use of
'the program implicitly means you have agreed to these terms.
'
'Permission is granted to anyone to use this software for any purpose,
'including commercial use, and to alter and redistribute it, provided that
'the following conditions are met:
'
'1. All redistributions of source code files must retain all copyright
'   notices that are currently in place, and this list of conditions without
'   any modification.
'
'2. All redistributions in binary form must retain all occurrences of the
'   above copyright notice and web site addresses that are currently in
'   place (for example, in the About boxes).
'
'3. Modified versions in source or binary form must be plainly marked as
'   such, and must not be misrepresented as being the original software.

'Credits:   Paul Caton - Subclassing
'           Matt Usner - Rounded Region Code (+ feedback on beta version - thanks!)
'           Heriberto Mantilla Santamaria - Creating Window from PictureBox
'           Anders Lyman - Mouse Wheel Support
'           Fred.cpp - Solid Part of DrawRect
'           Dana Seaman - DrawText wrapper & Unicode Res file for Demo
'           Gary Noble - API Scroll Bars & concise bug reports

'Updates (dd/mm/yy):
'04/07/05   Added ColumnResize property for end-user Column resizing
'           Subclassed WM_CTLCOLORSCROLLBAR for scrollbars
'05/07/05   Added CustomSort Event, FocusRectColor + FocusRectStyle Properties
'07/07/05   Removed ColumnHeadingHeight Property (always calculated from DropDownFont size) seems -
'           unneccessary to be able to change it!
'           Tweaked Sort Arrow position
'           ListIndex adjusted when Text Property changed and in Standard Style
'08/07/05   Custom Border Style (+ Border Colours, Button Colors)
'15/07/05   Added DrawText wrapper function (to support DrawTextW on NT based systems)
'22/07/05   sblargechange(efsVertical) =  PageScrollItems property (thanks Jeff Mayes)
'27/07/05   Enabled State tweaks
'06/08/05   Converted to API Scroll Bars
'           Index variables changed from Integer to Long to allow more than 32667 Items
'13/10/05   Removed obsolete code
'25/01/06   Redraw & behaviour tweaks + RTE6 when navigating with keyboard on list with >32767 Items
'27/01/06   Bug Fixes (report by Alex - thankyou!)
'           Categorized Properties
'30/01/06   Keyboard tweaks
'#############################################################################################################################

'Windows API Declarations
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
 
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
 
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_BTNFACE = 15

Private Const CLR_INVALID = &HFFFF

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_SINGLELINE = &H20

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_FLAT = &H4000
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const SWP_FRAMECHANGED          As Long = &H20
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOSIZE                As Long = &H1

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_TOOLWINDOW = &H80
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2

Private Const SRCCOPY = &HCC0020
Private Const SRCAND = &H8800C6
Private Const MERGEPAINT = &HBB0226

Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X           As Long
    Y           As Long
End Type

'#############################################################################################################################
'Subclassing Code (all credits to Paul Caton!)
Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                             As Long
  dwFlags                            As TRACKMOUSEEVENT_FLAGS
  hwndTrack                          As Long
  dwHoverTime                        As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private mInCtrl                      As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Enum eMsgWhen
    MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Type tSubData                                                                   'Subclass data type
    hwnd                               As Long                                            'Handle of the window being subclassed
    nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
    nMsgCntA                           As Long                                            'Msg after table entry count
    nMsgCntB                           As Long                                            'Msg before table entry count
    aMsgTblA()                         As Long                                            'Msg after table array
    aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WM_SETFOCUS            As Long = &H7
Private Const WM_KILLFOCUS           As Long = &H8
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSEHOVER As Long = &H2A1
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114
Private Const WM_LBUTTONDOWN         As Long = &H201
Private Const WM_RBUTTONDOWN         As Long = &H204
Private Const WM_GETMINMAXINFO       As Long = &H24
Private Const WM_SIZE                As Long = &H5
Private Const WM_WINDOWPOSCHANGED    As Long = &H47
Private Const WM_WINDOWPOSCHANGING   As Long = &H46

'################################################################
'API Scroll Bars
Private Declare Function InitialiseFlatSB Lib "Comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function FlatSB_EnableScrollBar Lib "Comctl32.dll" (ByVal hwnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "Comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "Comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "Comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "Comctl32.dll" (ByVal hwnd As Long, ByVal Index As Long, ByVal NewValue As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function UninitializeFlatSB Lib "Comctl32.dll" (ByVal hwnd As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Enum ScrollBarOrienationEnum
    Scroll_Horizontal
    Scroll_Vertical
    Scroll_Both
End Enum

Public Enum ScrollBarStyleEnum
    Style_Regular = 1& ' FSB_REGULAR_MODE
    Style_Flat = 0& 'FSB_FLAT_MODE
End Enum

Public Enum EFSScrollBarConstants
    efsHorizontal = 0 'SB_HORZ
    efsVertical = 1 'SB_VERT
End Enum

Private Const SB_BOTTOM = 7
Private Const SB_ENDSCROLL = 8
Private Const SB_HORZ = 0
Private Const SB_LEFT = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_RIGHT = 7
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_VERT = 1

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const ESB_DISABLE_BOTH = &H3
Private Const ESB_ENABLE_BOTH = &H0
Private Const MK_CONTROL = &H8
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&
Private Const FSB_FLAT_MODE = 1&
Private Const FSB_REGULAR_MODE = 0&

Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private m_bInitialised      As Boolean
Private m_eOrientation      As ScrollBarOrienationEnum
Private m_eStyle            As ScrollBarStyleEnum
Private m_hWnd              As Long
Private m_lSmallChangeHorz  As Long
Private m_lSmallChangeVert  As Long
Private m_bEnabledHorz      As Boolean
Private m_bEnabledVert      As Boolean
Private m_bVisibleHorz      As Boolean
Private m_bVisibleVert      As Boolean
Private m_bNoFlatScrollBars As Boolean

'#############################################################################################################################
'User Control Declarations
Private Const BORDER_LEFT = 4
Private Const BORDER_TOP = 3
Private Const BUTTON_WIDTH = 15
Private Const SCROLLBAR_SIZE = 16

Private Const DEF_ALIGNMENT = vbLeftJustify
Private Const DEF_AUTOCOMPLETE = False
Private Const DEF_BACKCOLOR = vbWindowBackground
Private Const DEF_BORDERCOLOR = vbBlack
Private Const DEF_BORDERCURVE = 5
Private Const DEF_BORDERSTYLE = 1
Private Const DEF_BORDERWIDTH = 1
Private Const DEF_BUTTONBACKCOLOR = vbButtonFace
Private Const DEF_CACHE_INCREMENT = 10
Private Const DEF_COLS = 1
Private Const DEF_COLUMNHEADERS = False
Private Const DEF_COLUMNRESIZE = False
Private Const DEF_COLUMNSORT = False
Private Const DEF_DROPDOWNAUTOWIDTH = False
Private Const DEF_DROPDOWNITEMSVISIBLE = 8
Private Const DEF_DROPDOWNWIDTH = 0
Private Const DEF_DEFAULTITEMFORECOLOR = vbWindowText
Private Const DEF_EDITABLE = False
Private Const DEF_ENABLED = True
Private Const DEF_FOCUSRECTCOLOR = &HFFFF&
Private Const DEF_FOCUSRECTSTYLE = 1
Private Const DEF_FORECOLOR = vbWindowText
Private Const DEF_INTEGRALHEIGHT = False
Private Const DEF_LOCKED = False
Private Const DEF_PAGESCROLLITEMS = 8
Private Const DEF_REQUIRECHECKEDITEM = False
Private Const DEF_ROWHEIGHTMIN = 0
Private Const DEF_SCALEUNITS = vbTwips
Private Const DEF_SEARCHCOLUMN = 0
Private Const DEF_STYLE = 0
Private Const DEF_TEXTALL = "-- Tutti --"
Private Const DEF_TEXTNONE = "-- Nessuno --"
Private Const DEF_TEXTSELECTION = "-- Selezione --"

Private Const MAX_ITEMS As Long = 2147483647
Private Const EVENT_TIMEOUT As Long = 500
Private Const AUTOSCROLL_TIMEOUT As Long = 50
Private Const NULL_RESULT = -1

Private Enum FlagsEnum
    flgChecked = 2
    flgSelected = 4
    flgBold = 8
End Enum

Private Enum SearchEnum
    cvEqual = 0
    cvGreaterEqual = 1
    cvLike = 2
End Enum

Public Enum ColAlignmentEnum
    AlignLeftTop = DT_LEFT Or DT_TOP
    AlignLeftCenter = DT_LEFT Or DT_VCENTER
    AlignLeftBottom = DT_LEFT Or DT_BOTTOM
    AlignCenterTop = DT_CENTER Or DT_TOP
    AlignCenterCenter = DT_CENTER Or DT_VCENTER
    AlignCenterBottom = DT_CENTER Or DT_BOTTOM
    AlignRightTop = DT_RIGHT Or DT_TOP
    AlignRightCenter = DT_RIGHT Or DT_VCENTER
    AlignRightBottom = DT_RIGHT Or DT_BOTTOM
End Enum

Public Enum BorderStyleEnum
    BorderNone = 0
    BorderSunken = 1
    BorderRaised = 2
    BorderFlat = 3
    BorderCustom = 4
End Enum

Public Enum ColTypeEnum
    TypeString = 0
    TypeNumeric = 1
    TypeDate = 2
    TypeCustom = 3
End Enum

Public Enum FocusRectStyleEnum
    FocusRectNone = 0
    FocusRectLight = 1
    FocusRectHeavy = 2
End Enum

Public Enum SortOrderEnum
    Ascending = 1
    Descending = 0
End Enum

Public Enum StyleEnum
    Standard = 0
    Checkboxes = 1
    OptionButtons = 2
End Enum

#If False Then
    Private flgChecked, flgSelected, flgBold, ctString, ctNumeric, ctNumeric
#End If

Private Type ColType
    nAlignment As ColAlignmentEnum
    dCustomWidth As Single
    lWidth As Long
    nSortOrder As Integer
    nType As Integer
    bVisible As Boolean
    sHeading As String
End Type

Private Type ItemType
    vImage As Variant
    lForeColor As Long
    lItemData As Long
    nFlags As Byte
    sValue() As String
End Type

'Data
Private mCols() As ColType
Private mItems() As ItemType
Private mPositions() As Long
Private mItemCount As Long
Private mListIndex As Long

'Misc
Private mImageList As Object
Private mHotImageList As Object

Private mInFocus As Boolean
Private mMouseDown As Boolean
Private mButtonIndex As Integer
Private mResizeCol As Integer
Private mButtonRect As RECT
Private mButtonClickTick As Long
Private mScrollTick As Long
Private mIgnoreKeyPress As Boolean
Private mLockTextBoxEvent As Boolean
Private mWindowsNT As Boolean
Private mSelectedText As String

'Properties
Private mAlignment As AlignmentConstants
Private mAutoComplete As Boolean
Private mBackColor As OLE_COLOR
Private mBorderColor As OLE_COLOR
Private mBorderCurve As Long
Private mBorderStyle As BorderStyleEnum
Private mBorderWidth As Long
Private mButtonBackColor As OLE_COLOR
Private mCacheIncrement As Integer
Private mColumnHeaders As Boolean
Private mColumnResize As Boolean
Private mColumnSort As Boolean
Private mDefaultItemForeColor As OLE_COLOR
Private mDisplayEllipsis  As Boolean
Private mDropDownAutoWidth As Boolean
Private mDropDownFont As Font
Private mDropDownItemsVisible As Integer
Private mDropDownWidth As Single
Private mEditable As Boolean
Private mEnabled As Boolean
Private mFocusRectColor As OLE_COLOR
Private mFocusRectStyle As FocusRectStyleEnum
Private mFont As Font
Private mFormatString As String
Private mForeColor As OLE_COLOR
Private mHighlighted As Long
Private mHotBorderColor As OLE_COLOR
Private mHotButtonBackColor As OLE_COLOR
Private mIntegralHeight As Boolean
Private mLocked As Boolean
Private mMaxLength As Integer
Private mPageScrollItems As Integer
Private mRequireCheckedItem As Boolean
Private mRowHeightMin As Single
Private mScaleUnits As ScaleModeConstants
Private mSortColumn As Integer
Private mSortSubColumn As Integer
Private mSearchColumn As Integer
Private mStyle As StyleEnum
Private mTextAll As String
Private mTextNone As String
Private mTextSelection As String
Private mChiudiLista As Boolean

'Events
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Scroll()

Public Event AutoCompleteSearch(ListIndex As Long)
Public Event ClickItem(ListIndex As Long, Button As Integer, Shift As Integer)
Public Event CustomSort(bAscending As Boolean, nCol As Integer, sValue1 As String, sValue2 As String, bSwap As Boolean)
Public Event DropDownClose()
Public Event DropDownOpen()
Public Event RequestItemChecked(ListIndex As Long, bValue As Boolean, bCancel As Boolean)
Public Event RequestListChecked(bValue As Boolean, bCancel As Boolean)
Public Event SelectionChanged()
Public Event SortComplete()
Public Event ClickPopUp(ValoreCliccato As Long)

'
'
'
'
'-------------------------------------------------------------------------------
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

' Per il menu
Private m_hMenu     As Long
'Public Caption As String ' If works as a sub-menu, this would be add as parent menu item

Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, lpNewItem As String) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

Private Const MF_STRING = &H0&
Private Const MF_SEPARATOR = &H800&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_POPUP = &H10&
Private Const MF_MENUBREAK = &H40&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_DISABLED = &H2&
Private Const MF_CHECKED = &H8&

Private Const APIFALSE As Long = 0
Private Const APITRUE As Long = 1

' hMenu E' l'indicatore del menu da mostrare attraverso il Popup-Menu.
'
' fuFlags E' la combinazione dei seguenti flags che specificano come posizionare il menu ed altre caratteristiche della funzione:
' TPM_CENTERALIGN: posiziona il Popup-Menu in modo tale che la X che si andrà ad indicare specificherà la coordinata del centro orizzontale del Popup-Menu;
' TPM_LEFTALIGN: posiziona il Popup-Menu in modo tale che la X rappresenta lo spigolo sinistro del Popup-Menu;
' TPM_RIGHTALIGN: posiziona il Popup-Menu in modo tale che la X rappresenta lo spigolo destro del Popup-Menu;
' TPM_BOTTOMALIGN: posiziona il Popup-Menu in modo tale che la Y che si andrà ad indicare specificherà il limite inferiore del Popup-Menu;
' TPM_TOPALIGN: posiziona il Popup-Menu in modo tale che la Y che si andrà ad indicare specificherà il limite superiore del Popup-Menu;
' TPM_VCENTERALIGN: posiziona il Popup-Menu in modo tale che la Y che si andrà ad indicare specificherà la coordinata del centro verticale del Popup-Menu;
' TPM_NONOTIFY: non manda un messaggio alla finestra alla quale il Popup-Menu è legata per indicarle quale scelta è stata compiuta dall'utente;
' TPM_LEFTBUTTON: permette All'utente di operare scelte sul Popup-Menu solo attraverso il pulsante sinistro del mouse o la tastiera;
' TPM_RIGHTBUTTON: permette All'utente di operare scelte sul Popup-Menu attraverso entrambi i pulsanti del mouse e la tastiera;
' TPM_HORIZONTAL: se il Popup-Menu non può essere mostrato interamente alle coordinate prescelte, preserva le dimensioni in orizzontale adattando quelle verticali;
' TPM_VERTICAL: se il Popup-Menu non può essere mostrato interamente alle coordinate prescelte, preserva le dimensioni in verticale adattando quelle orizzontali;
'
' x La coordinata X del Popup-Menu
' y La coordinata Y del Popup-Menu
' hwnd l'intero di riferimento della finestra che ospita il Popup-Menu. Questa finestra riceverà il messaggio del tipo: WM_COMMAND.
' lptpm Riporta informazioni addizionali relative alla posizione del Popup-Menu.

Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Long) As Long
Private Declare Function GetCursorPosa Lib "user32" Alias "GetCursorPos" (lpPoint As POINTLa) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0

Private Type POINTLa
    X As Long
    Y As Long
End Type
Private Const TPM_RETURNCMD = &H100&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_LEFTBUTTON = &H0&

Private Declare Function GetWindowRecta Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, lpRect As RECTa) As Long
Private Type RECTa
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'-------------------------------------------------------------------------------

Private Sub SetListIndex(NewValue As Long)
    mListIndex = NewValue
    mHighlighted = NewValue
    
    If mListIndex >= 0 And mListIndex <= mItemCount Then
        mSelectedText = mItems(mPositions(mListIndex)).sValue(0)
        ShowText mInFocus
    Else
        mSelectedText = ""
        ShowText mInFocus
    End If
End Sub

'Subclass handler
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'THIS MUST BE THE FIRST PUBLIC ROUTINE IN THIS FILE.
'That includes public properties also
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data

    Dim eBar As EFSScrollBarConstants
    Dim lv As Long, lSC As Long
    Dim lScrollCode As Long
    Dim tSI As SCROLLINFO
    Dim zDelta As Long
    Dim lHSB As Long
    Dim lVSB As Long
    
    lHSB = SBValue(efsHorizontal)
    lVSB = SBValue(efsVertical)
    
    Select Case uMsg
    Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL
        lScrollCode = (wParam And &HFFFF&)

        Select Case uMsg
        
            Case WM_HSCROLL ' Get the scrollbar type
                eBar = efsHorizontal
                
            Case WM_VSCROLL
                eBar = efsVertical
                
            Case Else     'WM_MOUSEWHEEL
                eBar = IIf(lScrollCode And MK_CONTROL, efsHorizontal, efsVertical)
                lScrollCode = IIf(wParam / 65536 < 0, SB_LINEDOWN, SB_LINEUP)
                
        End Select

        Select Case lScrollCode
        
            Case SB_THUMBTRACK
            
                ' Is vertical/horizontal?
                pSBGetSI eBar, tSI, SIF_TRACKPOS
                SBValue(eBar) = tSI.nTrackPos

            Case SB_LEFT, SB_BOTTOM
                 SBValue(eBar) = IIf(lScrollCode = 7, SBMax(eBar), SBMin(eBar))

            Case SB_RIGHT, SB_TOP
                 SBValue(eBar) = SBMin(eBar)

            Case SB_LINELEFT, SB_LINEUP
            
                If SBVisible(eBar) Then
                
                    lv = SBValue(eBar)
                    If (eBar = efsHorizontal) Then
                        lSC = m_lSmallChangeHorz
                    Else
                        lSC = m_lSmallChangeVert
                    End If
                    
                    If (lv - lSC < SBMin(eBar)) Then
                         SBValue(eBar) = SBMin(eBar)
                    Else
                         SBValue(eBar) = lv - lSC
                    End If
                    
                End If

            Case SB_LINERIGHT, SB_LINEDOWN
                If SBVisible(eBar) Then
        
                    lv = SBValue(eBar)
                    
                    If (eBar = efsHorizontal) Then
                        lSC = m_lSmallChangeHorz
                    Else
                        lSC = m_lSmallChangeVert
                    End If
                    
                    If (lv + lSC > SBMax(eBar)) Then
                         SBValue(eBar) = SBMax(eBar)
                    Else
                         SBValue(eBar) = lv + lSC
                    End If
                End If

            Case SB_PAGELEFT, SB_PAGEUP
                 SBValue(eBar) = SBValue(eBar) - SBLargeChange(eBar)

            Case SB_PAGERIGHT, SB_PAGEDOWN
                 SBValue(eBar) = SBValue(eBar) + SBLargeChange(eBar)

            Case SB_ENDSCROLL

        End Select
        
        If (lVSB <> SBValue(efsVertical)) Or (lHSB <> SBValue(efsHorizontal)) Then
            mScrollTick = GetTickCount()
            ShowItems
        End If
    
    Case WM_KILLFOCUS
        'Another Control has got the focus
        DoKillFocus
        
    Case WM_MOUSEMOVE
        SetTimer False
        
        If mEnabled Then
            If Not mInCtrl Then
                mInCtrl = True
                DrawComboBorder
                
                Call TrackMouseLeave(lng_hWnd)
                Call TrackMouseHover(lng_hWnd, 0)
            End If
        End If
        
        If IsMouseInScrollArea() Then
            DoAutoScroll
        End If
    
    Case WM_MOUSELEAVE
        mInCtrl = False
        
        If mBorderStyle = BorderCustom Then
            DrawComboBorder
        End If
            
        If picList.Visible Then
            Call GetAsyncKeyState(VK_LBUTTON)
            Call GetAsyncKeyState(VK_RBUTTON)
    
            SetTimer True
        End If
    
    Case WM_MOUSEHOVER
        If mEnabled Then
            mInCtrl = False
        End If
    
    Case WM_MOUSEWHEEL
        If mInFocus Then
            Select Case wParam
            Case Is > False
                If SBValue(efsVertical) > SBMin(efsVertical) Then
                    SBValue(efsVertical) = SBValue(efsVertical) - 1
                End If
            
            Case Else
                If SBValue(efsVertical) < SBMax(efsVertical) Then
                    SBValue(efsVertical) = SBValue(efsVertical) + 1
                End If
                
            End Select
        End If
          
    Case WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED, WM_GETMINMAXINFO, WM_SIZE, WM_LBUTTONDOWN, WM_RBUTTONDOWN
        'If Parent form is changing we want to close!
        DoKillFocus
        
    End Select
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
    MsgBox "ComboView Control 1.0.4 ©2005, Richard Mewett", vbInformation
End Sub

Public Sub AddItem(ByVal Item As String, Optional Index As Integer = -1, Optional Checked As Boolean, Optional Image As Variant)
    Dim lCount As Long
    Dim nCol As Integer
    Dim sText() As String
    
    'mItems() is an array of the Items in the ComboBox
    'mPositions() is an array of "pointers" to mItems()
    
    'The pointer technique is used to allow much faster Inserts & Sorts
    'since we only need to swap an Integer (2 bytes) rather than a large
    'data structure (a UDT in this case)
    
    'The mItems() is resized incrementally to reduce the Redim Preserve
    'overhead. Since we will only ever be too large by CACHE_INCREMENT (10)
    'the potential unused allocated memory is minimal
    
    'Note MAX_ITEMS is the Max value for a long - in practice if you tried
    'to load anywhere near this many items the memory overhead would be
    'enormous!
    If mItemCount = MAX_ITEMS Then
        Err.Raise 381, "ComboView", "Maximum ListItems Exceeded"
    Else
        mItemCount = mItemCount + 1
        If mItemCount > UBound(mItems) Then
            ReDim Preserve mItems(mItemCount + mCacheIncrement)
            ReDim Preserve mPositions(mItemCount + mCacheIncrement)
        End If
        
        If (Index >= 0) And (Index < mItemCount) Then
            If mItemCount > 1 Then
                For lCount = mItemCount To Index + 1 Step -1
                    mPositions(lCount) = mPositions(lCount - 1)
                Next lCount
                mPositions(Index) = mItemCount
            End If
        Else
            mPositions(mItemCount) = mItemCount
        End If
        
        ReDim mItems(mItemCount).sValue(UBound(mCols))
        
        If UBound(mCols) > 0 Then
            sText() = Split(Item, vbTab)
            For lCount = LBound(sText) To UBound(sText)
                mItems(mItemCount).sValue(nCol) = sText(lCount)
                nCol = nCol + 1
                If nCol > UBound(mCols) Then
                    Exit For
                End If
            Next lCount
        Else
            mItems(mItemCount).sValue(0) = Item
        End If
        
        With mItems(mItemCount)
            .lForeColor = mDefaultItemForeColor
            .vImage = Image
        
            If Checked Then
                SetFlag mItemCount, flgChecked, True
            End If
            
            'Default Bold
            If mDropDownFont.Bold Then
                SetFlag mItemCount, flgBold, True
            End If
        End With
    End If
End Sub

Friend Property Get hMenu() As Long
    ' Per il menuPopUp
    hMenu = m_hMenu
End Property

Public Sub AddItemMenu(ByVal nID As Long, varItem As Variant, Optional bDefault As Boolean = False, _
    Optional bChecked As Boolean = False, Optional bDisabled As Boolean = False, _
    Optional bGrayed As Boolean = False, Optional bNewColumn As Boolean = False)
    ' Per il menuPopUp
    
    If nID <= 0 Then
        MsgBox "Il valore di nID deve essere maggiore di 0" & vbNewLine & "Questo menu non verrà inserito.", vbInformation, "ComboView"
        Exit Sub
    End If
    
    If TypeName(varItem) = "String" Then
        If varItem = "-" Then
            AppendMenu m_hMenu, MF_STRING Or MF_SEPARATOR, nID, ByVal vbNullString
        Else
            AppendMenu m_hMenu, MF_STRING Or IIf(bNewColumn, MF_MENUBARBREAK, 0) Or IIf(bChecked, MF_CHECKED, 0), nID, ByVal varItem
        End If
    'ElseIf TypeOf varItem Is clsMenu Then
        'Dim cSubMenu As clsMenu
        'Set cSubMenu = varItem
        'AppendMenu m_hMenu, MF_STRING Or MF_POPUP Or IIf(bNewColumn, MF_MENUBARBREAK, 0), cSubMenu.hMenu, ByVal cSubMenu.Caption
    End If
    
    If bDefault Then SetMenuDefaultItem m_hMenu, nID, APIFALSE
    If bGrayed Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_GRAYED
    If bDisabled Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_DISABLED
    
End Sub

Private Function TrackPopupCmb(Optional ByVal hwnd As Long = -1, Optional ByVal PopX As Long = -1, Optional ByVal PopY As Long = -1, Optional ByVal hWndOfBeneathControl As Long = -1) As Long
    ' Per il menuPopUp
    Dim h As Long, X As Long, Y As Long
    
    If hwnd = -1 Or hwnd = 0 Then
        ' Find the top window of current process
        Dim hDesktop As Long: hDesktop = GetDesktopWindow()
        Dim hChild As Long: hChild = GetWindow(hDesktop, GW_CHILD)
        Dim idCurrent As Long: idCurrent = GetCurrentProcessId()
        Do While hChild
            Dim idChild As Long
            GetWindowThreadProcessId hChild, idChild
            If idChild = idCurrent Then Exit Do
            hChild = GetWindow(hChild, GW_HWNDNEXT)
        Loop
        
        If hChild = 0 Then Err.Raise -1, "TrackPopupCmb", "Cannot find top window of current process!"
        h = hChild
    Else
        h = hwnd
    End If
    
    If hWndOfBeneathControl <> -1 Then
        Dim rt As RECTa
        GetWindowRecta hWndOfBeneathControl, rt
        X = rt.Left
        Y = rt.Bottom
    Else
        Dim pt As POINTLa
        GetCursorPosa pt
        If PopX = -1 Then X = pt.X Else: X = PopX
        If PopY = -1 Then Y = pt.Y Else: Y = PopY
    End If
    
    TrackPopupCmb = TrackPopupMenu(m_hMenu, TPM_RETURNCMD + TPM_LEFTALIGN + TPM_LEFTBUTTON, X, Y, 0, h, 0)
    
End Function

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Misc"
    Alignment = mAlignment
End Property

Public Property Let Alignment(ByVal NewValue As AlignmentConstants)
    mAlignment = NewValue
    txtCombo.Alignment = mAlignment
    
    PropertyChanged "Alignment"
End Property

Public Property Get AutoComplete() As Boolean
Attribute AutoComplete.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoComplete = mAutoComplete
End Property

Public Property Let AutoComplete(ByVal NewValue As Boolean)
    mAutoComplete = NewValue
    
    PropertyChanged "AutoComplete"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    mBackColor = NewValue
    
    With UserControl
        .BackColor = mBackColor
        .Picture = .Image
    End With
    
    txtCombo.BackColor = mBackColor
    picList.BackColor = mBackColor
    ShowText mInFocus
    
    PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderColor = mBorderColor
End Property

Public Property Let BorderColor(ByVal NewValue As OLE_COLOR)
    mBorderColor = NewValue
    DrawComboBorder
    
    PropertyChanged "BorderColor"
End Property

Public Property Get BorderCurve() As Long
Attribute BorderCurve.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderCurve = mBorderCurve
End Property

Public Property Let BorderCurve(ByVal NewValue As Long)
    mBorderCurve = NewValue
    DrawComboBorder
    
    PropertyChanged "BorderCurve"
End Property

Public Property Get BorderStyle() As BorderStyleEnum
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As BorderStyleEnum)
    mBorderStyle = NewValue
    DrawComboBorder
    ShowText mInFocus
    
    PropertyChanged "BorderStyle"
End Property

Public Property Get BorderWidth() As Long
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderWidth = mBorderWidth
End Property

Public Property Let BorderWidth(ByVal NewValue As Long)
    mBorderWidth = NewValue
    DrawComboBorder
    
    PropertyChanged "BorderWidth"
End Property

Public Property Get ButtonBackColor() As OLE_COLOR
Attribute ButtonBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ButtonBackColor = mButtonBackColor
End Property

Public Property Let ButtonBackColor(ByVal NewValue As OLE_COLOR)
    mButtonBackColor = NewValue
    DrawComboBorder
    
    PropertyChanged "ButtonBackColor"
End Property

Public Sub Clear(Optional CancellaSoloValori As Boolean = False)
    Dim tmpEditable As Boolean
    ReDim mItems(0)
    ReDim mPositions(0)
    
    
    mItemCount = -1
    mListIndex = -1
    tmpEditable = mEditable
    mEditable = True
    Me.Text = ""
    mEditable = tmpEditable
    
    If CancellaSoloValori = False Then
        mButtonIndex = NULL_RESULT
        mSortColumn = NULL_RESULT
        mSortSubColumn = NULL_RESULT
        
        mResizeCol = NULL_RESULT
    End If
    
End Sub

Public Property Get ColAlignment(ByVal Index As Integer) As ColAlignmentEnum
Attribute ColAlignment.VB_ProcData.VB_Invoke_Property = ";List"
    ColAlignment = mCols(Index).nAlignment
End Property

Public Property Let ColAlignment(ByVal Index As Integer, ByVal NewValue As ColAlignmentEnum)
    mCols(Index).nAlignment = NewValue
End Property

Public Property Get ColHeading(ByVal Index As Integer) As String
Attribute ColHeading.VB_ProcData.VB_Invoke_Property = ";List"
    ColHeading = mCols(Index).sHeading
End Property

Public Property Let ColHeading(ByVal Index As Integer, ByVal NewValue As String)
    mCols(Index).sHeading = NewValue
End Property

Public Property Get Cols() As Integer
Attribute Cols.VB_ProcData.VB_Invoke_Property = ";List"
    Cols = UBound(mCols) + 1
End Property

Public Property Let Cols(ByVal NewValue As Integer)
    Dim nCol As Integer
    
    If NewValue > 0 Then
        ReDim mCols(0 To NewValue - 1)
        For nCol = LBound(mCols) To UBound(mCols)
            mCols(nCol).dCustomWidth = 1000
            mCols(nCol).lWidth = ScaleX(mCols(nCol).dCustomWidth, mScaleUnits, vbPixels)
            mCols(nCol).bVisible = True
        Next nCol
    Else
        ReDim mCols(0)
    End If
End Property

Public Property Get ColType(ByVal Index As Integer) As ColTypeEnum
    ColType = mCols(Index).nType
End Property

Public Property Let ColType(ByVal Index As Integer, ByVal NewValue As ColTypeEnum)
    mCols(Index).nType = NewValue
End Property

Public Property Get ColumnHeaders() As Boolean
Attribute ColumnHeaders.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ColumnHeaders = mColumnHeaders
End Property

Public Property Let ColumnHeaders(ByVal NewValue As Boolean)
    mColumnHeaders = NewValue
    
    PropertyChanged "ColumnHeaders"
End Property

Public Property Get ColumnResize() As Boolean
Attribute ColumnResize.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ColumnResize = mColumnResize
End Property

Public Property Let ColumnResize(ByVal NewValue As Boolean)
    mColumnResize = NewValue
    
    PropertyChanged "ColumnResize"
End Property

Public Property Get ColumnSort() As Boolean
Attribute ColumnSort.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ColumnSort = mColumnSort
End Property

Public Property Let ColumnSort(ByVal NewValue As Boolean)
    mColumnSort = NewValue
    
    PropertyChanged "ColumnSort"
End Property

Public Property Get ColVisible(ByVal Index As Integer) As Boolean
    ColVisible = mCols(Index).bVisible
End Property

Public Property Let ColVisible(ByVal Index As Integer, ByVal NewValue As Boolean)
    mCols(Index).bVisible = NewValue
End Property

Public Property Get ColWidth(ByVal Index As Integer) As Single
    ColWidth = mCols(Index).dCustomWidth
End Property

Public Property Let ColWidth(ByVal Index As Integer, ByVal NewValue As Single)
    'dCustomWidth is in the Units the Control is operating in
    mCols(Index).dCustomWidth = NewValue
    
    'lWidth is always Pixels (because thats what API functions require) and
    'is calculated to prevent repeated Width Scaling calculations
    mCols(Index).lWidth = ScaleX(mCols(Index).dCustomWidth, mScaleUnits, vbPixels)
End Property

Public Property Get DefaultItemForeColor() As OLE_COLOR
Attribute DefaultItemForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DefaultItemForeColor = mDefaultItemForeColor
End Property

Public Property Let DefaultItemForeColor(ByVal NewValue As OLE_COLOR)
    mDefaultItemForeColor = NewValue
    
    PropertyChanged "DefaultItemForeColor"
End Property

Public Property Get DisplayEllipsis() As Boolean
Attribute DisplayEllipsis.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DisplayEllipsis = mDisplayEllipsis
End Property

Public Property Let DisplayEllipsis(ByVal NewValue As Boolean)
    mDisplayEllipsis = NewValue
    
    PropertyChanged "DisplayEllipsis"
End Property

Private Sub DoAutoScroll()
    Const MAX_COUNT As Long = 2147483647
    
    Static bActive As Boolean
    Dim uPoint  As POINTAPI
    Dim uRect As RECT
    Dim lCount As Long
    
    'This scrolls the list up/down when the mouse moves outside the DropDown
    'and the left button is pressed. It will terminate as soon as the mouse
    'moves back into the DropDown or the control loses focus
    
    'Prevent recursion
    If Not bActive Then
        bActive = True
        'Debug.Print "DoAutoScroll >"
        
        Call GetWindowRect(picList.hwnd, uRect)
        
        Do While mInFocus
            If (GetTickCount() - mScrollTick) > AUTOSCROLL_TIMEOUT Then
                mScrollTick = GetTickCount()
                
                Call GetCursorPos(uPoint)
                
                If (uPoint.Y < uRect.Top) Then
                    If SBValue(efsVertical) > SBMin(efsVertical) Then
                        mHighlighted = mHighlighted - 1
                        SBValue(efsVertical) = SBValue(efsVertical) - 1
                        ShowItems
                    End If
                ElseIf (uPoint.Y > uRect.Bottom) Then
                    If SBValue(efsVertical) < SBMax(efsVertical) Then
                        mHighlighted = mHighlighted + 1
                        SBValue(efsVertical) = SBValue(efsVertical) + 1
                        ShowItems
                    End If
                Else
                    Exit Do
                End If
            End If
            
            lCount = lCount + 1
            If (lCount Mod 10) = 0 Then
                DoEvents
            ElseIf lCount = MAX_COUNT Then
                lCount = 0
            End If
        Loop
        
        bActive = False
        'Debug.Print "DoAutoScroll <"
    End If
End Sub

Private Sub DoKillFocus()
    If picList.Visible Then
        SetDropDown
    End If

    If mInFocus Then
        mInFocus = False
        ShowText False
    End If
End Sub

Private Sub DoSort()
    If (mSortColumn = NULL_RESULT) And (mSortSubColumn <> NULL_RESULT) Then
        mSortColumn = mSortSubColumn
        mSortSubColumn = NULL_RESULT
    ElseIf mSortColumn = mSortSubColumn Then
        mSortSubColumn = NULL_RESULT
    End If
    
    SortArray LBound(mItems), mItemCount, mSortColumn, mCols(mSortColumn).nSortOrder
    SortSubList
    
    RaiseEvent SortComplete
End Sub

Private Sub DrawComboBorder()
    Const ARROW_HEIGHT = 3
    Const ARROW_WIDTH = 5

    Static bResetRegion As Boolean
    
    Dim r As RECT
    Dim lColor As Long
    Dim hBrush As Long
    Dim hRgn1  As Long
    Dim hRgn2  As Long
    Dim lX As Long
    Dim lY As Long
    
    '#############################################################################################################################
    'This draws the Border of the ComboBox and the Dropdown Button
    '#############################################################################################################################
    
    On Local Error GoTo DrawComboBorderError
    
    With mButtonRect
        .Left = txtCombo.Width + BORDER_LEFT + 1
        .Top = BORDER_TOP - 1
        .Right = .Left + BUTTON_WIDTH
        .Bottom = .Top + UserControl.ScaleHeight - BORDER_TOP - 1
    End With
    
    With UserControl
        Call SetRect(r, 0, 0, .ScaleWidth, .ScaleHeight)
        DrawRect .hdc, mButtonRect, TranslateColor(mBackColor), True
        
        If mBorderStyle = BorderCustom Then
            Call SetRect(r, txtCombo.Width + BORDER_LEFT + 1, 0, .ScaleWidth, .ScaleHeight)
            If mInCtrl Then
                DrawRect .hdc, r, TranslateColor(mHotButtonBackColor), True
            Else
                DrawRect .hdc, r, TranslateColor(mButtonBackColor), True
            End If
        Else
            DrawRect .hdc, mButtonRect, TranslateColor(mButtonBackColor), True
        
            If bResetRegion Then
                hRgn1 = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
                SetWindowRgn hwnd, hRgn1, True
                
                SetWindowRgn picList.hwnd, hRgn1, True
                DeleteObject hRgn1
                
                bResetRegion = False
            End If
            
            Call DrawEdge(.hdc, mButtonRect, EDGE_RAISED, BF_RECT)
        End If

        Select Case mBorderStyle
        Case BorderSunken
            Call DrawEdge(.hdc, r, EDGE_SUNKEN, BF_RECT)
        
        Case BorderRaised
            Call DrawEdge(.hdc, r, EDGE_RAISED, BF_RECT)
        
        Case BorderFlat
            Call DrawEdge(.hdc, r, EDGE_SUNKEN, BF_RECT Or BF_FLAT)
        
        Case BorderCustom
            If mInCtrl Then
                lColor = TranslateColor(mHotBorderColor)
            Else
                lColor = TranslateColor(mBorderColor)
            End If
            
            hRgn1 = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, mBorderCurve, mBorderCurve)
            hRgn2 = CreateRoundRectRgn(mBorderWidth, mBorderWidth, ScaleWidth - mBorderWidth, ScaleHeight - mBorderWidth, mBorderCurve, mBorderCurve)
            CombineRgn hRgn2, hRgn1, hRgn2, 3
            
            hBrush = CreateSolidBrush(lColor)
            FillRgn hdc, hRgn2, hBrush
            
            SetWindowRgn hwnd, hRgn1, True
            SetWindowRgn picList.hwnd, hRgn1, True
            
            DeleteObject hRgn2
            DeleteObject hBrush
            DeleteObject hRgn1
            
            bResetRegion = True
        
        Case Else
            .Picture = Nothing
        
        End Select
         
        lX = mButtonRect.Left + (BUTTON_WIDTH / 2) - (ARROW_WIDTH / 2)
        lY = (.ScaleHeight / 2) - (ARROW_HEIGHT / 2)
    
        If mEnabled Then
            Call BitBlt(.hdc, lX, lY, ARROW_WIDTH, ARROW_HEIGHT, picImages.hdc, 42, 0, MERGEPAINT)
            Call BitBlt(.hdc, lX, lY, ARROW_WIDTH, ARROW_HEIGHT, picImages.hdc, 42, 0, SRCAND)
        Else
            Call BitBlt(.hdc, lX, lY, ARROW_WIDTH, ARROW_HEIGHT, picImages.hdc, 42, 4, MERGEPAINT)
            Call BitBlt(.hdc, lX, lY, ARROW_WIDTH, ARROW_HEIGHT, picImages.hdc, 42, 4, SRCAND)
        End If
        
        .Picture = .Image
    End With
    Exit Sub
    
DrawComboBorderError:
    Exit Sub
End Sub

Private Sub DrawRect(hdc As Long, RC As RECT, lColor As Long, bFilled As Boolean)
    Dim lNewBrush As Long
  
    lNewBrush = CreateSolidBrush(lColor)
    
    If bFilled Then
        Call FillRect(hdc, RC, lNewBrush)
    Else
        Call FrameRect(hdc, RC, lNewBrush)
    End If

    Call DeleteObject(lNewBrush)
End Sub

Private Sub DrawText(ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long)
    If mWindowsNT Then
        DrawTextW hdc, StrPtr(lpString), nCount, lpRect, wFormat
    Else
        DrawTextA hdc, lpString, nCount, lpRect, wFormat
    End If
End Sub

Public Property Get DropDownAutoWidth() As Boolean
Attribute DropDownAutoWidth.VB_ProcData.VB_Invoke_Property = ";List"
    DropDownAutoWidth = mDropDownAutoWidth
End Property

Public Property Let DropDownAutoWidth(ByVal NewValue As Boolean)
    mDropDownAutoWidth = NewValue
End Property

Public Property Get DropDownFont() As Font
Attribute DropDownFont.VB_ProcData.VB_Invoke_Property = ";Font"
   Set DropDownFont = mDropDownFont
End Property

Public Property Set DropDownFont(ByVal NewValue As StdFont)
    Set mDropDownFont = NewValue
    
    PropertyChanged "DropDownFont"
End Property

Public Property Get DropDownItemsVisible() As Integer
Attribute DropDownItemsVisible.VB_ProcData.VB_Invoke_Property = ";List"
    DropDownItemsVisible = mDropDownItemsVisible
End Property

Public Property Let DropDownItemsVisible(ByVal NewValue As Integer)
    mDropDownItemsVisible = NewValue
    
    PropertyChanged "DropDownItemsVisible"
End Property

Public Property Get DropDownWidth() As Single
Attribute DropDownWidth.VB_ProcData.VB_Invoke_Property = ";List"
    DropDownWidth = mDropDownWidth
End Property

Public Property Let DropDownWidth(ByVal NewValue As Single)
    mDropDownWidth = NewValue
    
    PropertyChanged "DropDownWidth"
End Property

Public Property Get Editable() As Boolean
Attribute Editable.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Editable = mEditable
End Property

Public Property Let Editable(ByVal NewValue As Boolean)
    mEditable = NewValue
    txtCombo.Visible = mEditable
    
    ShowText mInFocus
    
    PropertyChanged "Editable"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    mEnabled = NewValue
    txtCombo.Enabled = mEnabled
    
    DrawComboBorder
    ShowText mInFocus
End Property

Public Property Get FocusRectColor() As OLE_COLOR
Attribute FocusRectColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FocusRectColor = mFocusRectColor
End Property

Public Property Let FocusRectColor(ByVal NewValue As OLE_COLOR)
    mFocusRectColor = NewValue
    
    PropertyChanged "FocusRectColor"
End Property

Public Property Get FocusRectStyle() As FocusRectStyleEnum
Attribute FocusRectStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FocusRectStyle = mFocusRectStyle
End Property

Public Property Let FocusRectStyle(ByVal NewValue As FocusRectStyleEnum)
    mFocusRectStyle = NewValue
    
    PropertyChanged "FocusRectStyle"
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
   Set Font = mFont
End Property

Public Property Set Font(ByVal NewValue As StdFont)
    Set mFont = NewValue
    
    Set UserControl.Font = mFont
    Set txtCombo.Font = mFont
   
    If mIntegralHeight Then
        UserControl.Height = ScaleY(UserControl.TextHeight("A") + (BORDER_TOP * 2), vbPixels, vbTwips)
        UserControl_Resize
    End If
    ShowText mInFocus
   
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    mForeColor = NewValue
    
    With UserControl
        .ForeColor = mForeColor
    End With
    
    txtCombo.ForeColor = mForeColor
    ShowText mInFocus
    
    PropertyChanged "ForeColor"
End Property

Public Property Get FormatString() As String
Attribute FormatString.VB_ProcData.VB_Invoke_Property = ";List"
    FormatString = mFormatString
End Property

Public Property Let FormatString(ByVal NewValue As String)
    Dim lCol As Long
    Dim sCols() As String
    
    mFormatString = NewValue
    
    sCols() = Split(NewValue, "|")
    If UBound(sCols()) > UBound(mCols) Then
        Cols = UBound(sCols()) + 1
    End If
    
    For lCol = LBound(sCols) To UBound(sCols)
        Select Case Mid$(sCols(lCol), 1, 1)
        Case "^"
            mCols(lCol).sHeading = Mid$(sCols(lCol), 2)
            mCols(lCol).nAlignment = AlignCenterTop
        Case "<"
            mCols(lCol).sHeading = Mid$(sCols(lCol), 2)
            mCols(lCol).nAlignment = AlignLeftTop
        Case ">"
            mCols(lCol).sHeading = Mid$(sCols(lCol), 2)
            mCols(lCol).nAlignment = AlignRightTop
        Case Else
            mCols(lCol).sHeading = sCols(lCol)
        End Select
        
        mCols(lCol).dCustomWidth = 1000
        mCols(lCol).lWidth = ScaleX(mCols(lCol).dCustomWidth, mScaleUnits, vbPixels)
        mCols(lCol).bVisible = True
    Next lCol
    
    PropertyChanged "FormatString"
End Property

Private Function GetColFromX(X As Single, Optional lColPosX As Long) As Integer
    Dim lX As Long
    Dim nCol As Integer
    
    GetColFromX = -1
    
    For nCol = SBValue(efsHorizontal) To UBound(mCols)
        If (X > lX) And (X < lX + mCols(nCol).lWidth) Then
            lColPosX = lX
            GetColFromX = nCol
        End If
        
        lX = lX + mCols(nCol).lWidth
    Next nCol
End Function

Private Function GetColumnHeadingHeight() As Long
    With picList
        GetColumnHeadingHeight = .TextHeight("A") + 4
    End With
End Function

Private Function GetFlag(ByVal Index As Long, nFlag As FlagsEnum) As Boolean
    'Gets information by bit flags for a ListItem.
    
    If mItems(Index).nFlags And nFlag Then
        GetFlag = True
    End If
End Function

Private Function GetRowFromY(Y As Single) As Long
    Dim lColumnHeadingHeight As Long
    Dim lRow As Long
    
    With picList
        If mColumnHeaders Then
            lColumnHeadingHeight = GetColumnHeadingHeight()
            
            If Y > lColumnHeadingHeight Then
                lRow = ((Y - lColumnHeadingHeight) \ GetRowHeight()) + SBValue(efsVertical)
            Else
                lRow = -1
            End If
        Else
            lRow = (Y \ GetRowHeight()) + SBValue(efsVertical)
        End If
    End With
    
    If lRow <= mItemCount Then
        GetRowFromY = lRow
    Else
        GetRowFromY = -1
    End If
End Function

Private Function GetRowHeight() As Long
    If mRowHeightMin > 0 Then
        GetRowHeight = picList.ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
    Else
        GetRowHeight = picList.TextHeight("A")
    End If
End Function

Public Property Get hdc() As Long
   hdc = UserControl.hdc
End Property

Public Property Get HotBorderColor() As OLE_COLOR
Attribute HotBorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HotBorderColor = mHotBorderColor
End Property

Public Property Let HotBorderColor(ByVal NewValue As OLE_COLOR)
    mHotBorderColor = NewValue
    DrawComboBorder
    
    PropertyChanged "HotBorderColor"
End Property

Public Property Get HotButtonBackColor() As OLE_COLOR
Attribute HotButtonBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HotButtonBackColor = mHotButtonBackColor
End Property

Public Property Let HotButtonBackColor(ByVal NewValue As OLE_COLOR)
    mHotButtonBackColor = NewValue
    DrawComboBorder
    
    PropertyChanged "HotButtonBackColor"
End Property

Public Property Get HotImageList() As Object
    Set HotImageList = mHotImageList
End Property

Public Property Let HotImageList(ByVal NewValue As Object)
    Set mHotImageList = NewValue
End Property

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property

Public Property Let ImageList(ByVal NewValue As Object)
    Set mImageList = NewValue
End Property

Public Property Get IntegralHeight() As Boolean
Attribute IntegralHeight.VB_ProcData.VB_Invoke_Property = ";Behavior"
    IntegralHeight = mIntegralHeight
End Property

Public Property Let IntegralHeight(ByVal NewValue As Boolean)
    mIntegralHeight = NewValue
    
    If mIntegralHeight Then
        UserControl.Height = ScaleY(UserControl.TextHeight("A") + (BORDER_TOP * 2), vbPixels, vbTwips)
        UserControl_Resize
        ShowText mInFocus
    End If
    
    PropertyChanged "IntegralHeight"
End Property

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    Call FreeLibrary(hMod)
  End If
End Function
'END Subclassing Code===================================================================================

Private Function IsMouseInScrollArea() As Boolean
    Dim uPoint  As POINTAPI
    Dim uRect As RECT
    
    Call GetWindowRect(picList.hwnd, uRect)
    Call GetCursorPos(uPoint)
    
    If (uPoint.Y < uRect.Top) Or (uPoint.Y > uRect.Bottom) Then
        IsMouseInScrollArea = True
    End If
End Function

Public Property Get ItemChecked(ByVal Index As Long) As Boolean
    ItemChecked = GetFlag(mPositions(Index), flgChecked)
End Property

Public Property Let ItemChecked(ByVal Index As Long, ByVal NewValue As Boolean)
    SetFlag mPositions(Index), flgChecked, NewValue
End Property

Public Property Let ItemData(ByVal Index As Long, NewValue As Long)
    mItems(mPositions(Index)).lItemData = NewValue
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
    ItemData = mItems(mPositions(Index)).lItemData
End Property

Public Property Get ItemFontBold(ByVal Index As Long) As Boolean
    ItemFontBold = mItems(mPositions(Index)).nFlags And flgBold
End Property

Public Property Let ItemFontBold(ByVal Index As Long, ByVal NewValue As Boolean)
    SetFlag Index, flgBold, NewValue
End Property

Public Property Get ItemForeColor(ByVal Index As Long) As Long
    ItemForeColor = mItems(mPositions(Index)).lForeColor
End Property

Public Property Let ItemForeColor(ByVal Index As Long, ByVal NewValue As Long)
    mItems(mPositions(Index)).lForeColor = NewValue
End Property

Public Property Let ItemImage(ByVal Index As Long, NewValue As Variant)
    mItems(mPositions(Index)).vImage = NewValue
End Property

Public Property Get ItemImage(ByVal Index As Long) As Variant
    ItemImage = mItems(mPositions(Index)).vImage
End Property

Public Property Get ItemText(ByVal Index As Long, ByVal Item As Long) As String
    If UBound(mItems(mPositions(Index)).sValue) >= Item Then
        ItemText = mItems(mPositions(Index)).sValue(Item)
    End If
End Property

Public Property Let ItemText(ByVal Index As Long, ByVal Item As Long, NewValue As String)
    If UBound(mItems(mPositions(Index)).sValue) >= Item Then
        mItems(mPositions(Index)).sValue(Item) = NewValue
    End If
End Property

Public Property Get List(ByVal Index As Long) As String
    List = mItems(mPositions(Index)).sValue(0)
End Property

Public Property Get ListCount() As Long
    ListCount = mItemCount + 1
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = mListIndex
End Property

Public Property Let ListIndex(ByVal NewValue As Long)
    SetListIndex NewValue
    RaiseEvent Click
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Locked = mLocked
End Property

Public Property Let Locked(ByVal NewValue As Boolean)
    mLocked = NewValue
    txtCombo.Locked = mLocked
    
    PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Integer
Attribute MaxLength.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MaxLength = mMaxLength
End Property

Public Property Let MaxLength(ByVal NewValue As Integer)
    mMaxLength = NewValue
    txtCombo.MaxLength = mMaxLength
    
    PropertyChanged "MaxLength"
End Property

Private Function NavigateDown() As Boolean
    If mHighlighted < mItemCount Then
        NavigateDown = True
        
        mHighlighted = mHighlighted + 1
        If mHighlighted >= (SBValue(efsVertical) + mDropDownItemsVisible) Then
            SBValue(efsVertical) = SBValue(efsVertical) + 1
        End If
        
        ShowItems
    End If
End Function

Private Function NavigateUp() As Boolean
    If mHighlighted > 0 Then
        NavigateUp = True
        
        mHighlighted = mHighlighted - 1
        If mHighlighted < SBValue(efsVertical) Then
            SBValue(efsVertical) = SBValue(efsVertical) - 1
        End If
        
        ShowItems
    End If
End Function

Public Property Get NewIndex() As Long
    NewIndex = mItemCount
End Property

Private Property Get Orientation() As ScrollBarOrienationEnum
    SBOrientation = m_eOrientation
End Property

Public Property Get PageScrollItems() As Integer
Attribute PageScrollItems.VB_ProcData.VB_Invoke_Property = ";List"
    PageScrollItems = mPageScrollItems
End Property

Public Property Let PageScrollItems(ByVal NewValue As Integer)
    mPageScrollItems = NewValue
    SBLargeChange(efsVertical) = mPageScrollItems
    
    PropertyChanged "PageScrollItems"
End Property

Private Sub picList_Click()
    RaiseEvent Click
End Sub

Private Sub picList_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As RECT
    Dim lItemIndex As Long
    Dim lX As Long
    Dim bCancel As Boolean
    Dim bValue As Boolean
    
    If (Button = vbRightButton) Or (Button = vbLeftButton) And Not mLocked Then
        Call SetCapture(picList.hwnd)
        mMouseDown = True
        
        lItemIndex = GetRowFromY(Y)
        
        If lItemIndex >= 0 Then
            mListIndex = lItemIndex
            RaiseEvent ClickItem(lItemIndex, Button, Shift)
        
            Select Case mStyle
                Case Checkboxes
                    If Button = vbRightButton Then
                        bValue = Not GetFlag(mPositions(mListIndex), flgChecked)
                        mChiudiLista = False
                    ElseIf Button = vbLeftButton Then
                        bValue = True
                        mChiudiLista = True
                    End If
                    
                    RaiseEvent RequestItemChecked(mPositions(mListIndex), bValue, bCancel)
                    
                Case OptionButtons
                    bValue = True
                    RaiseEvent RequestItemChecked(mPositions(mListIndex), bValue, bCancel)
            End Select
            
            If Not bCancel Then
                SetFlag mPositions(mListIndex), flgChecked, bValue
                ShowItems
                SetText mListIndex

                RaiseEvent SelectionChanged
            End If
            
        ElseIf mColumnSort And (picList.MousePointer <> vbSizeWE) Then
            ' Si è fatto click sulle colonne.....
            mButtonIndex = GetColFromX(X, lX)
            If mButtonIndex <> NULL_RESULT Then
                With picList
                    Call SetRect(r, lX, 0, lX + mCols(mButtonIndex).lWidth, GetColumnHeadingHeight())
                    Call DrawEdge(.hdc, r, EDGE_SUNKEN, BF_RECT)
                    
                    .Refresh
                End With
            End If
        End If
    End If
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static lResizeX As Long
    
    Dim lWidth As Long
    Dim lItemIndex As Long
    Dim nPointer As Integer
    
    If Not mLocked Then
        If (Button = vbLeftButton) And (mResizeCol >= 0) Then
            'We are resizing a Column
            lWidth = (X - lResizeX)
            If lWidth > 1 Then
                mCols(mResizeCol).lWidth = lWidth
                mCols(mResizeCol).dCustomWidth = ScaleX(mCols(mResizeCol).lWidth, vbPixels, mScaleUnits)
                
                ShowItems
                picList.Refresh
            End If
        ElseIf Button = 0 Then
            'Only check for resize cursor if no buttons depressed
            lResizeX = 0
            mResizeCol = NULL_RESULT
            
            lItemIndex = GetRowFromY(Y)
            nPointer = vbDefault
            
            If (lItemIndex >= 0) Then
                If (mHighlighted <> lItemIndex) Then
                    mHighlighted = lItemIndex
                    ShowItems
                End If
            ElseIf mColumnResize Then
                 For lItemIndex = LBound(mCols) To UBound(mCols)
                    lWidth = lWidth + mCols(lItemIndex).lWidth
                    
                    If (X < lWidth + 2) And (X > lWidth - 2) Then
                        nPointer = vbSizeWE
                        mResizeCol = lItemIndex
                        Exit For
                    End If
                    
                    lResizeX = lResizeX + mCols(lItemIndex).lWidth
                Next lItemIndex
            End If
        
            With picList
                If .MousePointer <> nPointer Then
                    .MousePointer = nPointer
                End If
            End With
        End If
    End If
End Sub

Private Sub picList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mMouseRow As Long

    If Button = vbLeftButton Or Button = vbRightButton Then
        Call ReleaseCapture
        
        mMouseDown = False
        mMouseRow = GetRowFromY(Y)
        
        If (mResizeCol >= 0) Then
            SetScrollBars
            ShowItems
            
        ElseIf mLocked Then
            SetDropDown
            
        ElseIf (mMouseRow < 0) Then
            If (GetColFromX(X) = mButtonIndex) And (mButtonIndex <> NULL_RESULT) Then
                If (Shift And vbCtrlMask) And (mSortColumn <> NULL_RESULT) Then
                    If mSortSubColumn <> mButtonIndex Then
                        mCols(mButtonIndex).nSortOrder = 0
                    End If
                    mSortSubColumn = mButtonIndex
                Else
                    If mSortColumn <> mButtonIndex Then
                        mCols(mButtonIndex).nSortOrder = 0
                        mSortSubColumn = NULL_RESULT
                    End If
                    mSortColumn = mButtonIndex
                End If
                
                If mCols(mButtonIndex).nSortOrder = 0 Then
                    mCols(mButtonIndex).nSortOrder = 1
                Else
                    mCols(mButtonIndex).nSortOrder = 0
                End If
                
                DoSort
                ShowItems
                
            ElseIf mButtonIndex >= 0 Then
                ShowItems
            End If
            
        ElseIf (mStyle = Standard Or mChiudiLista = True) And (mResizeCol < 0) Then
            mListIndex = mMouseRow
            SetDropDown
        End If
    End If
End Sub

Private Sub pSBClearUp()
    If m_hWnd <> 0 Then
        On Error Resume Next
        ' Stop flat scroll bar if we have it:
        If Not (m_bNoFlatScrollBars) Then
            UninitializeFlatSB m_hWnd
        End If

        On Error GoTo 0
    End If
    m_hWnd = 0
    m_bInitialised = False
End Sub

Private Sub pSBCreateScrollBar()
    Dim lR As Long
    Dim hParent As Long

    On Error Resume Next
    lR = InitialiseFlatSB(m_hWnd)
    If (Err.Number <> 0) Then
        'Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
        ' Means we have version prior to 4.71
        ' We get standard scroll bars.
        m_bNoFlatScrollBars = True
    Else
        SBStyle = m_eStyle
    End If
End Sub

Private Sub pSBGetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim Lo As Long

    Lo = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        GetScrollInfo m_hWnd, Lo, tSI
    Else
        FlatSB_GetScrollInfo m_hWnd, Lo, tSI
    End If

End Sub

Private Sub pSBLetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim Lo As Long

    Lo = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        SetScrollInfo m_hWnd, Lo, tSI, True
    Else
        FlatSB_SetScrollInfo m_hWnd, Lo, tSI, True
    End If

End Sub

Private Sub pSBSetOrientation()
    ShowScrollBar m_hWnd, SB_HORZ, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Horizontal))
    ShowScrollBar m_hWnd, SB_VERT, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Vertical))
End Sub

Public Sub Refresh()
    'Reallocate the buffer to remove pre-allocated items
    If UBound(mItems) > mItemCount Then
        ReDim Preserve mItems(mItemCount)
        ReDim Preserve mPositions(mItemCount)
    End If
    
    SetText mListIndex
    
    If picList.Visible Then
        ShowItems
    End If
End Sub

Public Sub RemoveItem(ByVal Index As Long)
    Dim lCount As Long
    Dim lPosition As Long
   
    '#############################################################################################################################
    'See AddItem for details of the Arrays used
    '#############################################################################################################################
   
    lPosition = mPositions(Index)
    
    'Reset Item Data
    For lCount = mPositions(Index) To mItemCount - 1
        mItems(lCount) = mItems(lCount + 1)
    Next lCount
    
    'Adjust Item Pointers
    For lCount = Index To mItemCount - 1
        mPositions(lCount) = mPositions(lCount + 1)
    Next lCount
    
    'Validate Pointers for Items after deleted Item
    For lCount = 1 To mItemCount - 1
        If mPositions(lCount) > lPosition Then
            mPositions(lCount) = mPositions(lCount) - 1
        End If
    Next lCount
    
    mItemCount = mItemCount - 1
    If (mItemCount + mCacheIncrement) < UBound(mItems) Then
        ReDim Preserve mItems(mItemCount)
        ReDim Preserve mPositions(mItemCount)
    End If
End Sub

Public Property Get RequireCheckedItem() As Boolean
Attribute RequireCheckedItem.VB_ProcData.VB_Invoke_Property = ";Behavior"
    RequireCheckedItem = mRequireCheckedItem
End Property

Public Property Let RequireCheckedItem(ByVal NewValue As Boolean)
    mRequireCheckedItem = NewValue
    
    PropertyChanged "RequireCheckedItem"
End Property

Public Property Get RowHeightMin() As Single
Attribute RowHeightMin.VB_ProcData.VB_Invoke_Property = ";List"
    RowHeightMin = mRowHeightMin
End Property

Public Property Let RowHeightMin(ByVal NewValue As Single)
    mRowHeightMin = NewValue
    
    PropertyChanged "RowHeightMin"
End Property

Private Property Get SBCanBeFlat() As Boolean
    SBCanBeFlat = Not (m_bNoFlatScrollBars)
End Property

Private Sub SBCreate(ByVal hWndA As Long)
    pSBClearUp
    m_hWnd = hWndA
    pSBCreateScrollBar
End Sub

Private Property Get SBEnabled(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        SBEnabled = m_bEnabledHorz
    Else
        SBEnabled = m_bEnabledVert
    End If
End Property

Private Property Let SBEnabled(ByVal eBar As EFSScrollBarConstants, ByVal bEnabled As Boolean)
    Dim Lo As Long
    Dim lF As Long

    Lo = eBar
    If (bEnabled) Then
        lF = ESB_ENABLE_BOTH
    Else
        lF = ESB_DISABLE_BOTH
    End If
    If (m_bNoFlatScrollBars) Then
        EnableScrollBar m_hWnd, Lo, lF
    Else
        FlatSB_EnableScrollBar m_hWnd, Lo, lF
    End If

End Property

Private Property Get SBLargeChange(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_PAGE
    SBLargeChange = tSI.nPage
End Property

Private Property Let SBLargeChange(ByVal eBar As EFSScrollBarConstants, ByVal iLargeChange As Long)
    Dim tSI As SCROLLINFO

    pSBGetSI eBar, tSI, SIF_ALL
    tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
    tSI.nPage = iLargeChange
    pSBLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE
End Property

Private Property Get SBMax(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
    SBMax = tSI.nMax                                  ' - tSI.nPage
End Property

Private Property Let SBMax(ByVal eBar As EFSScrollBarConstants, ByVal iMax As Long)
    Dim tSI As SCROLLINFO
    tSI.nMax = iMax + SBLargeChange(eBar)
    tSI.nMin = SBMin(eBar)
    pSBLetSI eBar, tSI, SIF_RANGE
End Property

Private Property Get SBMin(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_RANGE
    SBMin = tSI.nMin
End Property

Private Property Let SBMin(ByVal eBar As EFSScrollBarConstants, ByVal iMin As Long)
    Dim tSI As SCROLLINFO
    tSI.nMin = iMin
    tSI.nMax = SBMax(eBar) + SBLargeChange(eBar)
    pSBLetSI eBar, tSI, SIF_RANGE
End Property

Private Property Let SBOrientation(ByVal eOrientation As ScrollBarOrienationEnum)
    m_eOrientation = eOrientation
    pSBSetOrientation
End Property

Private Sub SBRefresh()
    EnableScrollBar m_hWnd, SB_VERT, ESB_ENABLE_BOTH
End Sub

Private Property Get SBSmallChange(ByVal eBar As EFSScrollBarConstants) As Long
    If (eBar = efsHorizontal) Then
        SBSmallChange = m_lSmallChangeHorz
    Else
        SBSmallChange = m_lSmallChangeVert
    End If
End Property

Private Property Let SBSmallChange(ByVal eBar As EFSScrollBarConstants, ByVal lSmallChange As Long)
    If (eBar = efsHorizontal) Then
        m_lSmallChangeHorz = lSmallChange
    Else
        m_lSmallChangeVert = lSmallChange
    End If
End Property

Private Property Get SBStyle() As ScrollBarStyleEnum
    SBStyle = m_eStyle
End Property

Private Property Let SBStyle(ByVal eStyle As ScrollBarStyleEnum)
    Dim lR As Long
    If (m_bNoFlatScrollBars) Then
        ' can't do it..
        'Debug.Print "Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71."
        Exit Property
    Else
        If (m_eOrientation = Scroll_Horizontal) Or (m_eOrientation = Scroll_Both) Then
            lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, eStyle, True)
        End If
        If (m_eOrientation = Scroll_Vertical) Or (m_eOrientation = Scroll_Both) Then
            lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, eStyle, True)
        End If
        'Debug.Print lR
        m_eStyle = eStyle
    End If

End Property

Private Property Get SBValue(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_POS
    SBValue = tSI.nPos
End Property

Private Property Let SBValue(ByVal eBar As EFSScrollBarConstants, ByVal iValue As Long)
    Dim tSI As SCROLLINFO
    If (iValue <> SBValue(eBar)) Then
        tSI.nPos = iValue
        pSBLetSI eBar, tSI, SIF_POS
        'ReDrawList
    End If
End Property

Private Property Get SBVisible(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        SBVisible = m_bVisibleHorz
    Else
        SBVisible = m_bVisibleVert
    End If
End Property

Private Property Let SBVisible(ByVal eBar As EFSScrollBarConstants, ByVal bState As Boolean)
    If (eBar = efsHorizontal) Then
        m_bVisibleHorz = bState
    Else
        m_bVisibleVert = bState
    End If
    If (m_bNoFlatScrollBars) Then
        ShowScrollBar m_hWnd, eBar, Abs(bState)
    Else
        FlatSB_ShowScrollBar m_hWnd, eBar, Abs(bState)
    End If
End Property

Public Property Get ScaleUnits() As ScaleModeConstants
Attribute ScaleUnits.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleUnits = mScaleUnits
End Property

Public Property Let ScaleUnits(ByVal NewValue As ScaleModeConstants)
    mScaleUnits = NewValue
    
    PropertyChanged "ScaleUnits"
End Property

Private Function ScaleValue(ByVal lValue As Long, ByVal lMin As Long, ByVal lMax As Long) As Long
    If lValue > lMax Then
        ScaleValue = lMax
    ElseIf lValue < lMin Then
        ScaleValue = lMin
    Else
        ScaleValue = lValue
    End If
End Function

Private Function SearchCode(sCode As String, nMode As SearchEnum) As Long
    Dim lCount As Long
    
    SearchCode = NULL_RESULT
    
    For lCount = LBound(mItems) To mItemCount
        Select Case nMode
        Case cvEqual
            If UCase$(mItems(mPositions(lCount)).sValue(mSearchColumn)) = sCode Then
                SearchCode = lCount
                Exit For
            End If
        
        Case cvGreaterEqual
            If UCase$(Left$(mItems(mPositions(lCount)).sValue(mSearchColumn), Len(sCode))) >= sCode Then
                SearchCode = lCount
                Exit For
            End If
        
        Case cvLike
            If UCase$(mItems(mPositions(lCount)).sValue(mSearchColumn)) Like sCode & "*" Then
                SearchCode = lCount
                Exit For
            End If

        End Select
        
    Next lCount
End Function

Public Property Get SelCount() As Long
    Dim lCount As Long
    Dim lSelected As Long
    
    For lCount = LBound(mItems) To mItemCount
        If GetFlag(lCount, flgChecked) Then
            lSelected = lSelected + 1
        End If
    Next lCount
    
    SelCount = lSelected
End Property

Public Property Get SelLength() As Integer
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = txtCombo.SelLength
End Property

Public Property Let SelLength(ByVal NewValue As Integer)
    txtCombo.SelLength = NewValue
End Property

Public Property Get SelStart() As Integer
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = txtCombo.SelStart
End Property

Public Property Let SelStart(ByVal NewValue As Integer)
    txtCombo.SelStart = NewValue
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
    SelText = txtCombo.SelText
End Property

Public Property Let SelText(ByVal NewValue As String)
    txtCombo.SelText = NewValue
End Property

Private Sub SetDropDown(Optional bDrawButton As Boolean)
    Dim r As RECT
    Dim lHeight As Long
    Dim lColumnsWidth As Long
    Dim lRowHeight As Long
    Dim lWidth As Long
    Dim dLeft As Single
    Dim dTop As Single
    Dim nCount As Integer
    
    With picList
        'If List is open then Close...
        If .Visible Then
            SetTimer False
            SetText mListIndex
            
            If (mStyle = Standard) Or (mChiudiLista = True) And (mListIndex >= 0) Then
                RaiseEvent Click
            End If
            
            .Visible = False
            
            With txtCombo
                If .Visible Then
                    .SetFocus
                End If
            End With
            
            RaiseEvent DropDownClose
            
        ElseIf ListCount() > 0 Then
            RaiseEvent DropDownOpen
            
            If bDrawButton And (mBorderStyle <> BorderCustom) Then
                With UserControl
                    Call DrawEdge(.hdc, mButtonRect, EDGE_SUNKEN, BF_RECT)
                    .Picture = .Image
                End With
                mButtonClickTick = GetTickCount()
            End If
        
            mHighlighted = mListIndex
            
            'Calculate the total width of the visible Columns
            For nCount = LBound(mCols) To UBound(mCols)
                If mCols(nCount).bVisible Then
                    lColumnsWidth = lColumnsWidth + mCols(nCount).lWidth
                End If
            Next nCount
            
            If mDropDownAutoWidth Then
                lWidth = lColumnsWidth + SCROLLBAR_SIZE
            ElseIf mDropDownWidth > 0 Then
                lWidth = ScaleX(mDropDownWidth, mScaleUnits, vbPixels)
            Else
                lWidth = UserControl.ScaleWidth
            End If
            
            Set .Font = mDropDownFont

            'Calculate the visible Rows
            If mRowHeightMin > 0 Then
                lRowHeight = ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
            Else
                lRowHeight = .TextHeight("A")
            End If
            
            If ListCount() > mDropDownItemsVisible Then
                If lColumnsWidth > lWidth Then
                    lHeight = (lRowHeight * mDropDownItemsVisible) + 2 + SCROLLBAR_SIZE
                Else
                    lHeight = (lRowHeight * mDropDownItemsVisible) + 2
                End If
            Else
                lHeight = (lRowHeight * ListCount()) + 2
            End If
            
            If mColumnHeaders Then
                lHeight = lHeight + GetColumnHeadingHeight()
            End If
            
            'Size PictureBox - convert Pixels to Twips
            If lHeight > ScaleY(Screen.Height, vbTwips, vbPixels) Then
                .Height = Screen.Height
            Else
                .Height = ScaleY(lHeight, vbPixels, vbTwips)
            End If
            If lWidth > ScaleX(Screen.Width, vbTwips, vbPixels) Then
                .Width = Screen.Width
            Else
                .Width = ScaleX(lWidth, vbPixels, vbTwips)
            End If
            
            SBValue(efsHorizontal) = SBMin(efsHorizontal)
            SetScrollBars
            
            'Position PictureBox & apply Window attributes
            Call GetWindowRect(hwnd, r)
            dLeft = r.Left * Screen.TwipsPerPixelX
            If (dLeft + .Width) > Screen.Width Then
                dLeft = Screen.Width - .Width
            End If
            
            dTop = r.Bottom * Screen.TwipsPerPixelY
            If (dTop + .Height) > Screen.Height Then
                dTop = (r.Bottom * Screen.TwipsPerPixelY) - (UserControl.Height + .Height)
            End If
            
            Call picList.Move(dLeft, dTop)
            Call SetWindowPos(.hwnd, -1, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)
            
            If mEditable Then
                mListIndex = SearchCode(UCase$(mSelectedText), cvEqual)
            End If
            mHighlighted = mListIndex
            
            'Set Vertical ScrollBar position
            If ListCount() > mDropDownItemsVisible Then
                If mListIndex > SBMax(efsVertical) Then
                    SBValue(efsVertical) = SBMax(efsVertical)
                ElseIf mListIndex > 0 Then
                    SBValue(efsVertical) = mListIndex
                Else
                    SBValue(efsVertical) = 0
                End If
            Else
                SBValue(efsVertical) = 0
            End If
            
            ShowItems
            
            ShowText False
            .Visible = True
            .SetFocus
            
            SetTimer True
        End If
    End With
End Sub

Private Sub SetFlag(ByVal nIndex As Long, nFlag As FlagsEnum, bValue As Boolean)
    Dim lCount As Long

    'Sets information by bit flags for a ListItem.
    If (nFlag = flgChecked) And mRequireCheckedItem Then
        If SelCount() = 1 And Not bValue Then
            bValue = True
        End If
    End If
    
    If bValue = True Then
        ' Azzero tutti i ChechBox
        If (nFlag = flgChecked) And ((mStyle <> Checkboxes) Or mChiudiLista = True) Then
            For lCount = LBound(mItems) To UBound(mItems)
                If mItems(lCount).nFlags And nFlag Then
                    mItems(lCount).nFlags = mItems(lCount).nFlags Xor nFlag
                End If
            Next lCount
        End If
        
        ' Imposto il valore del CheckBox
        mItems(nIndex).nFlags = mItems(nIndex).nFlags Or nFlag
    Else
        If mItems(nIndex).nFlags And nFlag Then
            mItems(nIndex).nFlags = mItems(nIndex).nFlags Xor nFlag
        End If
    End If
End Sub

Private Sub SetFlags(nFlag As FlagsEnum, bValue As Boolean)
    Dim lCount As Long
    
    For lCount = LBound(mItems) To UBound(mItems)
        If bValue Then
            mItems(lCount).nFlags = mItems(lCount).nFlags Or nFlag
        Else
            If mItems(lCount).nFlags And nFlag Then
                mItems(lCount).nFlags = mItems(lCount).nFlags Xor nFlag
            End If
        End If
    Next lCount
End Sub

Public Sub SetItem(ByVal vData As Variant, Optional ByVal nDefault As Integer = -1)
    Dim lCount As Long
    Dim bFound As Boolean
    Dim bItemData As Boolean
    
    If VarType(vData) = vbLong Then
        bItemData = True
    End If

    For lCount = 0 To mItemCount
        If bItemData Then
            If vData = mItems(lCount).lItemData Then
                bFound = True
                ListIndex = lCount
                Exit For
            End If
        Else
            If vData = mItems(lCount).sValue(0) Then
                bFound = True
                ListIndex = lCount
                Exit For
            End If
        End If
    Next lCount
    
    If Not bFound And nDefault >= 0 Then
        ListIndex = nDefault
    End If
End Sub

Private Sub SetScrollBars()
    Dim lWidth As Long
    Dim lHeight As Long
    Dim lRowHeight As Long
    Dim nCount As Integer
    Dim nDropDownItemsVisible As Integer
    
    '#############################################################################################################################
    'Sets the visibilty of scroll bars and sets max scroll values
    '#############################################################################################################################
    
    'Calculate total width of columns
    For nCount = LBound(mCols) To UBound(mCols)
        If mCols(nCount).bVisible Then
            lWidth = lWidth + mCols(nCount).lWidth
        End If
    Next nCount
    
    SBVisible(efsVertical) = False
    
    If (lWidth > picList.ScaleWidth) Then
        SBMax(efsHorizontal) = UBound(mCols) - 1
        SBVisible(efsHorizontal) = True
    Else
        SBMax(efsHorizontal) = UBound(mCols)
        SBVisible(efsHorizontal) = False
    End If
    
    'Calculate total height available for drawing Items
    If mColumnHeaders Then
        lHeight = picList.ScaleHeight - ScaleY(GetColumnHeadingHeight(), mScaleUnits, vbTwips) - 4
    Else
        lHeight = picList.ScaleHeight - 4
    End If
     
    If mRowHeightMin > 0 Then
        lRowHeight = ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
    Else
        lRowHeight = picList.TextHeight("A")
    End If
    
    'This may differ from the DropDownItemsVisible Property if scroll bars
    'have been forced by user dragging a column wider
    nDropDownItemsVisible = (lHeight / lRowHeight)
    
    If ListCount() > nDropDownItemsVisible Then
        SBMax(efsVertical) = mItemCount - nDropDownItemsVisible
        SBVisible(efsVertical) = True
    Else
        SBMax(efsVertical) = mItemCount
        SBVisible(efsVertical) = False
    End If
End Sub

Private Sub SetText(lItemIndex As Long)
    Dim lCount As Long
    Dim lSelCount(1) As Long
    Dim lSelStart As Long
    Dim nCol As Integer
    
    If mStyle = Standard Then
        If lItemIndex >= 0 Then
            mSelectedText = mItems(mPositions(lItemIndex)).sValue(0)
            ShowText mInFocus
        End If
    Else
        For lCount = LBound(mCols) To UBound(mCols)
            If mCols(lCount).bVisible And mCols(lCount).dCustomWidth > 0 Then
                nCol = lCount
                Exit For
            End If
        Next lCount
        
        lSelStart = -1
        For lCount = LBound(mItems) To mItemCount
            If GetFlag(lCount, flgChecked) Then
                lSelCount(0) = lSelCount(0) + 1
                If lSelStart < 0 Then
                    lSelStart = lCount
                End If
            Else
                lSelCount(1) = lSelCount(1) + 1
            End If
        Next lCount
        
        If lSelCount(0) = 1 Then
            mListIndex = lSelStart
            mSelectedText = mItems(lSelStart).sValue(nCol)
        ElseIf (lSelCount(0) > 0) And (lSelCount(1) = 0) Then
            mSelectedText = mTextAll
        ElseIf (lSelCount(0) = 0) Then
            mSelectedText = mTextNone
        Else
            mSelectedText = mTextSelection
        End If
        
        ShowText mInFocus
    End If
End Sub

Private Sub SetTimer(bEnabled As Boolean)
    If tmrRelease.Enabled <> bEnabled Then
        If bEnabled Then
            tmrRelease.Enabled = True
            'Debug.Print "Timer ON"
        Else
            tmrRelease.Enabled = False
            'Debug.Print "Timer OFF"
        End If
    End If
End Sub

Private Sub ShowItems()
    Const CHECKBOX_SIZE = 11
    Const OPTIONBUTTON_SIZE = 10
    Const SORTARROW_SIZE = 8
    Const SMALL_SORTARROW_SIZE = 6
    
    Const HEADER_LEFT = 3
    Const IMAGE_LEFT = 2
    
    Dim r As RECT
    Dim lX As Long
    Dim lY As Long
    
    Dim lLeftImage As Long
    Dim lLeftText As Long
    Dim lTextHeight As Long
    Dim lColumnHeadingHeight As Long
    
    Dim lCBSpace As Long
    Dim lImageSpace As Long
    Dim lSortSpace As Long
    Dim nCount As Integer
    Dim lItem As Long
    Dim bRenderImages As Boolean
    Dim sText As String
    
    'Left Position to Draw Text
    If mStyle <> Standard Then
        lLeftText = 15
    Else
        lLeftText = 3
    End If
    
    'Left Position to Draw Images
    lLeftImage = ScaleX(lLeftText, vbPixels, vbTwips)
    
    'Adjust Text Position for Images
    If Not mImageList Is Nothing Then
        bRenderImages = True
        lImageSpace = ((GetRowHeight() - mImageList.ImageHeight) / 2)
        lLeftText = lLeftText + mImageList.ImageWidth + 2
    End If
    
    lCBSpace = ((GetRowHeight() - CHECKBOX_SIZE) / 2)
    
    With picList
        .Cls
        .DrawWidth = 1
        .ForeColor = vbWindowText
        
        lColumnHeadingHeight = GetColumnHeadingHeight()
        lTextHeight = .TextHeight("A")
        
        'Column Headers
        If mColumnHeaders Then
            Call SetRect(r, 0, 0, .ScaleWidth, lColumnHeadingHeight)
            DrawRect .hdc, r, GetSysColor(COLOR_BTNFACE), True
            
            For nCount = SBValue(efsHorizontal) To UBound(mCols)
                 If mCols(nCount).bVisible Then
                    'Draw the Column Header Buttons
                    Call SetRect(r, lX, 0, lX + mCols(nCount).lWidth, lColumnHeadingHeight)
                    Call DrawEdge(.hdc, r, EDGE_RAISED, BF_RECT)
                 
                    Call SetRect(r, lX + HEADER_LEFT, (lColumnHeadingHeight / 2) - (lTextHeight / 2), (lX + mCols(nCount).lWidth) - HEADER_LEFT, lColumnHeadingHeight)
                    
                    sText = mCols(nCount).sHeading
                    
                    'Format/Render Text
                    If mDisplayEllipsis Then
                        Call DrawText(.hdc, sText, -1, r, mCols(nCount).nAlignment Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
                    Else
                        Call DrawText(.hdc, sText, -1, r, mCols(nCount).nAlignment Or DT_SINGLELINE)
                    End If
                    
                    'Render Sort Arrows
                    If mCols(nCount).lWidth > SORTARROW_SIZE Then
                        If nCount = mSortColumn Then
                            lSortSpace = (lColumnHeadingHeight / 2) - (SORTARROW_SIZE / 2)
                            If mCols(nCount).nSortOrder = 1 Then
                                Call BitBlt(.hdc, r.Right - SORTARROW_SIZE, lY + lSortSpace, SORTARROW_SIZE, 7, picImages.hdc, 25, 14, MERGEPAINT)
                                Call BitBlt(.hdc, r.Right - SORTARROW_SIZE, lY + lSortSpace, SORTARROW_SIZE, 7, picImages.hdc, 1, 14, SRCAND)
                            Else
                                Call BitBlt(.hdc, r.Right - SORTARROW_SIZE, lY + lSortSpace, SORTARROW_SIZE, 7, picImages.hdc, 37, 14, MERGEPAINT)
                                Call BitBlt(.hdc, r.Right - SORTARROW_SIZE, lY + lSortSpace, SORTARROW_SIZE, 7, picImages.hdc, 13, 14, SRCAND)
                            End If
                        ElseIf nCount = mSortSubColumn Then
                            lSortSpace = (lColumnHeadingHeight / 2) - (SMALL_SORTARROW_SIZE / 2)
                            If mCols(nCount).nSortOrder = 1 Then
                                Call BitBlt(.hdc, r.Right - SMALL_SORTARROW_SIZE, lY + lSortSpace, SMALL_SORTARROW_SIZE, 5, picImages.hdc, 26, 23, MERGEPAINT)
                                Call BitBlt(.hdc, r.Right - SMALL_SORTARROW_SIZE, lY + lSortSpace, SMALL_SORTARROW_SIZE, 5, picImages.hdc, 2, 23, SRCAND)
                            Else
                                Call BitBlt(.hdc, r.Right - SMALL_SORTARROW_SIZE, lY + lSortSpace, SMALL_SORTARROW_SIZE, 5, picImages.hdc, 38, 23, MERGEPAINT)
                                Call BitBlt(.hdc, r.Right - SMALL_SORTARROW_SIZE, lY + lSortSpace, SMALL_SORTARROW_SIZE, 5, picImages.hdc, 14, 23, SRCAND)
                            End If
                        End If
                    End If
                    
                    lX = lX + mCols(nCount).lWidth
                End If
            Next nCount
            
            lY = lColumnHeadingHeight
        End If
        
        lTextHeight = GetRowHeight()
        
        'List Items
        For lItem = SBValue(efsVertical) To (SBValue(efsVertical) + mDropDownItemsVisible) - 1
            If lItem > mItemCount Then
                Exit For
            End If
            
            If lItem = mHighlighted Then
                'Draw Highlight & Focus Rectangles
                Call SetRect(r, lLeftText, lY, .ScaleWidth, lY + lTextHeight)
                DrawRect .hdc, r, GetSysColor(COLOR_HIGHLIGHT), True
                
                Select Case mFocusRectStyle
                Case FocusRectLight
                    Call DrawFocusRect(.hdc, r)
                Case FocusRectHeavy
                    DrawRect .hdc, r, TranslateColor(mFocusRectColor), False
                End Select
                
                .ForeColor = vbHighlightText
            Else
                .ForeColor = mItems(mPositions(lItem)).lForeColor
            End If
            .FontBold = mItems(mPositions(lItem)).nFlags And flgBold
            
            'Blit appropriate Checkbox Image
            Select Case mStyle
                Case Checkboxes
                    If mItems(mPositions(lItem)).nFlags And flgChecked Then
                        Call BitBlt(.hdc, IMAGE_LEFT, lY + lCBSpace, CHECKBOX_SIZE, CHECKBOX_SIZE, picImages.hdc, 11, 0, SRCCOPY)
                    Else
                        Call BitBlt(.hdc, IMAGE_LEFT, lY + lCBSpace, CHECKBOX_SIZE, CHECKBOX_SIZE, picImages.hdc, 0, 0, SRCCOPY)
                    End If
                
                Case OptionButtons
                    If mItems(mPositions(lItem)).nFlags And flgChecked Then
                        Call BitBlt(.hdc, IMAGE_LEFT, lY + lCBSpace, OPTIONBUTTON_SIZE, OPTIONBUTTON_SIZE, picImages.hdc, 32, 0, SRCCOPY)
                    Else
                        Call BitBlt(.hdc, IMAGE_LEFT, lY + lCBSpace, OPTIONBUTTON_SIZE, OPTIONBUTTON_SIZE, picImages.hdc, 22, 0, SRCCOPY)
                    End If
    
            End Select
            
            If bRenderImages Then
                'If we have an Image Index then Draw it
                If mItems(mPositions(lItem)).vImage <> Empty Then
                    If lItem = mHighlighted Then
                        mImageList.ListImages(mItems(mPositions(lItem)).vImage).Draw .hdc, lLeftImage, ScaleY(lY + lImageSpace, vbPixels, vbTwips), 2
                    Else
                        mImageList.ListImages(mItems(mPositions(lItem)).vImage).Draw .hdc, lLeftImage, ScaleY(lY + lImageSpace, vbPixels, vbTwips), 1
                    End If
                End If
            End If
            
            lX = -1
            For nCount = SBValue(efsHorizontal) To UBound(mCols)
                If mCols(nCount).bVisible Then
                    If lX < 0 Then
                        lX = 1
                        Call SetRect(r, lLeftText, lY, (lLeftText + mCols(nCount).lWidth) - lLeftText, lY + lTextHeight)
                     Else
                        Call SetRect(r, lX, lY, (lX + mCols(nCount).lWidth) - 3, lY + lTextHeight)
                     End If
                
                    sText = mItems(mPositions(lItem)).sValue(nCount)
                    
                    'Format/Render Text
                    If mDisplayEllipsis Then
                        Call DrawText(.hdc, sText, -1, r, mCols(nCount).nAlignment Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
                    Else
                        Call DrawText(.hdc, sText, -1, r, mCols(nCount).nAlignment Or DT_SINGLELINE)
                    End If
                    
                    lX = lX + mCols(nCount).lWidth
                End If
            Next nCount
            
            lY = lY + lTextHeight
        Next lItem
    End With
End Sub

Private Sub ShowText(Optional bFocus As Boolean)
    Dim r As RECT
    
    With UserControl
        .Cls
        
        'Are are using a Textbox or drawing Text?
        If mEditable Then
            mLockTextBoxEvent = True
            txtCombo.Text = mSelectedText
            mLockTextBoxEvent = False
            
            If bFocus Then
                txtCombo.SelStart = 0
                txtCombo.SelLength = Len(txtCombo.Text)
            End If
        Else
            If (mBorderStyle = BorderCustom) And (mBorderCurve > 0) Then
                Call SetRect(r, BORDER_LEFT + 2, BORDER_TOP, BORDER_LEFT + txtCombo.Width, BORDER_TOP + txtCombo.Height)
            Else
                Call SetRect(r, BORDER_LEFT - 1, BORDER_TOP, BORDER_LEFT + txtCombo.Width, BORDER_TOP + txtCombo.Height)
            End If
                        
            If mEnabled Then
                If bFocus Then
                    'Draw Highlight & Focus Rectangles
                    DrawRect .hdc, r, GetSysColor(COLOR_HIGHLIGHT), True
                    Call DrawFocusRect(.hdc, r)
                    .ForeColor = vbHighlightText
                Else
                    'Clear any previous Highlight/Focus Rectangles
                    DrawRect .hdc, r, TranslateColor(mBackColor), True
                End If
            Else
                DrawRect .hdc, r, TranslateColor(mBackColor), True
                .ForeColor = vbGrayText
            End If
                         
            r.Left = r.Left + 1
            r.Top = r.Top + 1
            
            Select Case txtCombo.Alignment
            Case vbRightJustify
                Call DrawText(.hdc, mSelectedText, -1, r, DT_RIGHT)
            Case vbCenter
                Call DrawText(.hdc, mSelectedText, -1, r, DT_CENTER)
            Case Else
                Call DrawText(.hdc, mSelectedText, -1, r, DT_LEFT)
            End Select
            
            .ForeColor = mForeColor
        End If
    End With
End Sub

Public Sub Sort(Column As Integer, SortOrder As SortOrderEnum, Optional SubColumn As Integer = -1, Optional SubSortOrder As Integer)
    mSortColumn = Column
    mCols(Column).nSortOrder = SortOrder
    
    mSortSubColumn = SubColumn
    If SubColumn >= 0 Then
        mCols(SubColumn).nSortOrder = SubSortOrder
    End If
    
    DoSort
End Sub

Private Sub SortArray(ByVal lFirst As Long, ByVal lLast As Long, nSortColumn As Integer, nSortType As Integer)
    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mPositions(lFirst), mPositions((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            Select Case mCols(nSortColumn).nType
            Case TypeDate
                bSwap = CDate(mItems(mPositions(lIndex)).sValue(nSortColumn)) > CDate(mItems(mPositions(lFirst)).sValue(nSortColumn))
            Case TypeNumeric
                bSwap = Val(mItems(mPositions(lIndex)).sValue(nSortColumn)) > Val(mItems(mPositions(lFirst)).sValue(nSortColumn))
            Case TypeCustom
                RaiseEvent CustomSort(True, nSortColumn, mItems(mPositions(lIndex)).sValue(nSortColumn), mItems(mPositions(lFirst)).sValue(nSortColumn), bSwap)
            
            Case Else
                bSwap = mItems(mPositions(lIndex)).sValue(nSortColumn) > mItems(mPositions(lFirst)).sValue(nSortColumn)
            End Select
        Else
            Select Case mCols(nSortColumn).nType
            Case TypeDate
                bSwap = CDate(mItems(mPositions(lIndex)).sValue(nSortColumn)) < CDate(mItems(mPositions(lFirst)).sValue(nSortColumn))
            Case TypeNumeric
                bSwap = Val(mItems(mPositions(lIndex)).sValue(nSortColumn)) < Val(mItems(mPositions(lFirst)).sValue(nSortColumn))
            Case TypeCustom
                RaiseEvent CustomSort(False, nSortColumn, mItems(mPositions(lIndex)).sValue(nSortColumn), mItems(mPositions(lFirst)).sValue(nSortColumn), bSwap)
            
            Case Else
                bSwap = mItems(mPositions(lIndex)).sValue(nSortColumn) < mItems(mPositions(lFirst)).sValue(nSortColumn)
            End Select
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mPositions(lBoundary), mPositions(lIndex)
        End If
    Next lIndex

    SwapLng mPositions(lFirst), mPositions(lBoundary)
    SortArray lFirst, lBoundary - 1, nSortColumn, nSortType
    SortArray lBoundary + 1, lLast, nSortColumn, nSortType
End Sub

Private Sub SortSubList()
    Dim lCount As Long
    Dim lStartSort As Long
    Dim bDifferent As Boolean
    Dim sMajorSort As String

    If mSortSubColumn > NULL_RESULT Then
        'Re-Sort the Items by a secondary column, preserving the sort sequence of the
        'primary sort
        
        lStartSort = LBound(mItems)
        For lCount = LBound(mItems) To mItemCount
            bDifferent = mItems(mPositions(lCount)).sValue(mSortColumn) <> sMajorSort
            If bDifferent Or lCount = mItemCount Then
                If lCount > 1 Then
                    If lCount - lStartSort > 1 Then
                        If lCount = mItemCount And Not bDifferent Then
                            SortArray lStartSort, lCount, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                        Else
                            SortArray lStartSort, lCount - 1, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                        End If
                    End If
                    lStartSort = lCount
                End If
                
                sMajorSort = mItems(mPositions(lCount)).sValue(mSortColumn)
            End If
        Next lCount
    End If
End Sub

Public Property Get Style() As StyleEnum
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Style = mStyle
End Property

Public Property Let Style(ByVal NewValue As StyleEnum)
    mStyle = NewValue
    
    If mStyle = OptionButtons Then
        SetFlags flgChecked, False
    End If
    SetText mListIndex
    
    PropertyChanged "Style"
End Property

'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'======================================================================================================================================================
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs

'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim I                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    I = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, I, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      I = I + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
Errs:
End Function

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
Errs:
End Sub

'Stop all subclassing
Private Sub Subclass_StopAll()
On Error GoTo Errs
  Dim I As Long
  
  I = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While I >= 0                                                                       'Iterate through each element
    With sc_aSubData(I)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    I = I - 1                                                                           'Next element
  Loop
Errs:
End Sub

Private Sub SwapLng(Value1 As Long, Value2 As Long)
    Dim lTemp As Long

    lTemp = Value1
    Value1 = Value2
    Value2 = lTemp
End Sub

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = mSelectedText
End Property

Public Property Let Text(ByVal NewValue As String)
    Dim lCount As Long
    
    'If mEditable Then
        If mStyle = Standard Then
            mListIndex = NULL_RESULT
            
            For lCount = LBound(mItems) To mItemCount
                If mItems(mPositions(lCount)).sValue(mSearchColumn) = NewValue Then
                    mListIndex = lCount
                    Exit For
                End If
            Next lCount
        End If
        
        mSelectedText = NewValue
        ShowText mInFocus
    'Else
    '    Err.Raise 383, "ComboView", "Text is Read-Only"
    'End If
End Property

Public Property Get TextAll() As String
Attribute TextAll.VB_ProcData.VB_Invoke_Property = ";Text"
    TextAll = mTextAll
End Property

Public Property Let TextAll(ByVal NewValue As String)
    mTextAll = NewValue
    
    PropertyChanged "TextAll"
End Property

Public Property Get TextNone() As String
Attribute TextNone.VB_ProcData.VB_Invoke_Property = ";Text"
    TextNone = mTextNone
End Property

Public Property Let TextNone(ByVal NewValue As String)
    mTextNone = NewValue
    
    PropertyChanged "TextNone"
End Property

Public Property Get TextSelection() As String
Attribute TextSelection.VB_ProcData.VB_Invoke_Property = ";Text"
    TextSelection = mTextSelection
End Property

Public Property Let TextSelection(ByVal NewValue As String)
    mTextSelection = NewValue
    
    PropertyChanged "TextSelection"
End Property

Private Sub tmrRelease_Timer()
    Dim uPoint  As POINTAPI
    Dim uRect As RECT
    Dim nLB As Integer
    Dim nRB As Integer
    
    '#############################################################################################################################
    'This is soley for detecting if we have clicked on a container which does not generate
    'WM_KILLFOCUS message for us to detect. i.e. the parent Form or a Frame
    
    'I don't like Timers in UserControls but wanted to make the Control behave as a normal Combo which
    'closes DropDown when the above situation occurs. I may still remove this "feature"!
    
    'NOTE: This Timer is only Enabled when we detect a WM_MOUSELEAVE so it does not fire unneccessarily
    'while the DropDown is displayed. It is Disabled as soon as the mouse re-enters the DropDown.
    '#############################################################################################################################
    
    Call GetCursorPos(uPoint)
    Call GetWindowRect(picList.hwnd, uRect)
        
    nLB = GetAsyncKeyState(VK_LBUTTON)
    nRB = GetAsyncKeyState(VK_RBUTTON)
    
    If (uPoint.X >= uRect.Left) And (uPoint.X <= uRect.Right) And (uPoint.Y >= uRect.Top) And (uPoint.Y <= uRect.Bottom) Then
        'The mouse pointer is within the Dropdown list
    ElseIf nLB Or nRB Then
        Select Case WindowFromPoint(uPoint.X, uPoint.Y)
        Case UserControl.hwnd
            'The mouse pointer is within the Control
        Case Else
            If (GetTickCount() - mScrollTick) > EVENT_TIMEOUT Then
                If picList.Visible Then
                    SetDropDown
                Else
                    SetTimer False
                End If
            End If
        
        End Select
    End If
End Sub

'Track the mouse hovering the indicated window
Private Sub TrackMouseHover(ByVal lng_hWnd As Long, lHoverTime As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_HOVER
      .hwndTrack = lng_hWnd
      .dwHoverTime = lHoverTime
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub

Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional hPalette As Long = 0) As Long
    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Sub txtCombo_Change()
    Dim lResult As Long
    Dim lSelStart As Long
    Dim sText As String
    
    If Not mLockTextBoxEvent Then
        If mAutoComplete Then
            With txtCombo
                lSelStart = .SelStart
                sText = Left$(.Text, lSelStart)
                If Len(sText) > 0 Then
                    'lResult = SearchCode(UCase$(sText), cvLike)
                    lResult = SearchCode(UCase$(sText), cvGreaterEqual)
                    'RaiseEvent AutoCompleteSearch(lResult)
                    
                    If (lResult > NULL_RESULT) Then
                        mLockTextBoxEvent = True
                        .SelText = Mid$(mItems(mPositions(lResult)).sValue(0), lSelStart + 1)
                        .SelStart = lSelStart
                        .SelLength = Len(.Text) - lSelStart
                        mLockTextBoxEvent = False
                    End If
                End If
            End With
        End If
        
        RaiseEvent Change
    End If
End Sub

Private Sub txtCombo_Click()
    RaiseEvent Click
End Sub

Private Sub txtCombo_DblClick()
    Dim bCancel As Boolean
    Dim bValue As Boolean
    
    If mStyle = Checkboxes Then
        bValue = (SelCount() <> ListCount())
        RaiseEvent RequestListChecked(bValue, bCancel)
    Else
        bCancel = True
    End If
    
    If bCancel Then
        RaiseEvent DblClick
    Else
        SetFlags flgChecked, bValue
        Refresh
        
        RaiseEvent SelectionChanged
    End If
End Sub

Private Sub txtCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lResult As Long
    
    If mAutoComplete And (Len(txtCombo.Text) > 0) Then
        Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            mLockTextBoxEvent = True
            txtCombo.SelText = ""
            mLockTextBoxEvent = False
        
        Case vbKeyReturn
            lResult = SearchCode(UCase$(txtCombo.Text), cvEqual)
            RaiseEvent AutoCompleteSearch(lResult)
            
            If (lResult > NULL_RESULT) Then
                ListIndex = lResult
                If picList.Visible Then
                    SetDropDown
                End If
            End If

        End Select
    End If
    
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtCombo_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtCombo_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtCombo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtCombo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtCombo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_DblClick()
    If mEnabled And (GetTickCount() - mButtonClickTick) > EVENT_TIMEOUT Then
        txtCombo_DblClick
    End If
End Sub

Private Sub UserControl_EnterFocus()
    'Debug.Print "UserControl_EnterFocus"
    
    mInFocus = True
    
    If Not picList.Visible Then
        ShowText True
    End If
End Sub

Private Sub UserControl_ExitFocus()
    DoKillFocus
End Sub

Private Sub UserControl_Initialize()
    Dim OS As OSVERSIONINFO
      
    OS.dwOSVersionInfoSize = Len(OS)
    Call GetVersionEx(OS)
    
    mWindowsNT = ((OS.dwPlatformID And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    
    ReDim mCols(0)
    Clear

    m_hMenu = CreatePopupMenu()

End Sub

Private Sub UserControl_InitProperties()
    Set mFont = Ambient.Font
    Set mDropDownFont = Ambient.Font
    
    mAlignment = DEF_ALIGNMENT
    mAutoComplete = DEF_AUTOCOMPLETE
    mBackColor = DEF_BACKCOLOR
    mBorderColor = DEF_BORDERCOLOR
    mBorderCurve = DEF_BORDERCURVE
    mBorderStyle = DEF_BORDERSTYLE
    mBorderWidth = DEF_BORDERWIDTH
    mButtonBackColor = DEF_BUTTONBACKCOLOR
    mCacheIncrement = DEF_CACHE_INCREMENT
    mColumnHeaders = DEF_COLUMNHEADERS
    mColumnResize = DEF_COLUMNRESIZE
    mColumnSort = DEF_COLUMNSORT
    mDefaultItemForeColor = DEF_DEFAULTITEMFORECOLOR
    mDropDownAutoWidth = DEF_DROPDOWNAUTOWIDTH
    mDropDownItemsVisible = DEF_DROPDOWNITEMSVISIBLE
    mDropDownWidth = DEF_DROPDOWNWIDTH
    mEditable = DEF_EDITABLE
    mEnabled = DEF_ENABLED
    mFocusRectColor = DEF_FOCUSRECTCOLOR
    mFocusRectStyle = DEF_FOCUSRECTSTYLE
    mForeColor = DEF_FORECOLOR
    mIntegralHeight = DEF_INTEGRALHEIGHT
    mLocked = DEF_LOCKED
    mMaxLength = 0
    mPageScrollItems = DEF_PAGESCROLLITEMS
    mRequireCheckedItem = DEF_REQUIRECHECKEDITEM
    mRowHeightMin = DEF_ROWHEIGHTMIN
    mScaleUnits = DEF_SCALEUNITS
    mSearchColumn = DEF_SEARCHCOLUMN
    mStyle = DEF_STYLE
    mTextAll = DEF_TEXTALL
    mTextNone = DEF_TEXTNONE
    mTextSelection = DEF_TEXTSELECTION
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If mEnabled Then
        If Shift And vbCtrlMask Then
            mIgnoreKeyPress = True
            
            If KeyCode = vbKeyA Then
                KeyCode = 0
                
                SetFlags flgChecked, (SelCount() <> ListCount())
                Refresh
                
                RaiseEvent SelectionChanged
            End If
        End If
        
        If picList.Visible Then
            Select Case KeyCode
            Case vbKeyF4
                mListIndex = mHighlighted
                SetDropDown
                
            Case vbKeyEscape
                SetDropDown
            
            Case vbKeyUp
                If Shift And vbAltMask Then
                    mListIndex = mHighlighted
                    SetDropDown
                ElseIf NavigateUp() Then
                    KeyCode = 0
                    SetText mHighlighted
                End If
            Case vbKeyDown
                If Shift And vbAltMask Then
                    mListIndex = mHighlighted
                    SetDropDown
                ElseIf NavigateDown() Then
                    KeyCode = 0
                    SetText mHighlighted
                End If
            
            Case vbKeyPageUp
                If mHighlighted > 0 Then
                    KeyCode = 0
                    mHighlighted = (mHighlighted - mDropDownItemsVisible) + 1
                    If mHighlighted < 0 Then
                        mHighlighted = 0
                    End If
                    
                    SBValue(efsVertical) = mHighlighted
                    ShowItems
                    
                    SetText mHighlighted
                End If
            
            Case vbKeyPageDown
                If mHighlighted < mItemCount Then
                    KeyCode = 0
                    mHighlighted = ScaleValue((mHighlighted + mDropDownItemsVisible) - 1, 0, mItemCount)
                    SBValue(efsVertical) = ScaleValue(mHighlighted, 0, SBMax(efsVertical))
                    ShowItems
                    SetText mHighlighted
                End If
            
            Case vbKeySpace
                If mHighlighted >= 0 Then
                    mIgnoreKeyPress = True
                    KeyCode = 0
                    
                    SetFlag mHighlighted, flgChecked, Not GetFlag(mHighlighted, flgChecked)
                    ShowItems
                    
                    RaiseEvent SelectionChanged
                End If
            
            Case vbKeyHome
                KeyCode = 0
                SetListIndex 0
                SBValue(efsVertical) = 0
                ShowItems
                
            Case vbKeyEnd
                KeyCode = 0
                SetListIndex mItemCount
                SBValue(efsVertical) = mItemCount
                ShowItems
                
            End Select
        Else
            Select Case KeyCode
            Case vbKeyF4
                SetDropDown
            
            Case vbKeyUp
                If Shift And vbAltMask Then
                    SetDropDown
                ElseIf mListIndex > 0 Then
                    KeyCode = 0
                    mListIndex = mListIndex - 1
                    SetText mListIndex
                    
                    RaiseEvent Click
                End If
            Case vbKeyDown
                If Shift And vbAltMask Then
                    SetDropDown
                ElseIf mListIndex < mItemCount Then
                    KeyCode = 0
                    mListIndex = mListIndex + 1
                    SetText mListIndex
                    
                    RaiseEvent Click
                End If
            
             Case vbKeyPageUp
                If mListIndex > 0 Then
                    KeyCode = 0
                    mListIndex = ScaleValue((mListIndex - mDropDownItemsVisible) + 1, 0, mItemCount)
                    SetText mListIndex
                    
                    RaiseEvent Click
                End If
           
             Case vbKeyPageDown
                If mListIndex < mItemCount Then
                    KeyCode = 0
                    mListIndex = ScaleValue((mListIndex + mDropDownItemsVisible) - 1, 0, mItemCount)
                    SetText mListIndex
                    
                    RaiseEvent Click
                End If
           
            End Select
        End If
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Static lTime As Long
    Static sCode As String
    Dim lResult As Long
    
    If picList.Visible And Not mIgnoreKeyPress Then
        If (GetTickCount() - lTime) < 1000 Then
            sCode = sCode & Chr$(KeyAscii)
        Else
            sCode = Chr$(KeyAscii)
        End If
        
        lTime = GetTickCount()
        
        lResult = SearchCode(UCase$(sCode), cvGreaterEqual)
        If lResult > NULL_RESULT Then
            mListIndex = lResult
            mHighlighted = lResult
            If mListIndex > SBMax(efsVertical) Then
                SBValue(efsVertical) = SBMax(efsVertical)
            Else
                SBValue(efsVertical) = mListIndex
            End If
            ShowItems
        End If
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    mIgnoreKeyPress = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled And (Button = vbLeftButton) Then
        If mStyle = Checkboxes Then
            If (X > mButtonRect.Left) Then
                SetDropDown True
            End If
        Else
            SetDropDown True
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled Then
        If X > UserControl.ScaleWidth Or X < 0 Or Y > UserControl.ScaleHeight Or Y < 0 Then
            ReleaseCapture
            mInCtrl = False
        ElseIf mInCtrl Then
            RaiseEvent MouseMove(Button, Shift, X, Y)
        Else
            mInCtrl = True
            Call TrackMouseLeave(UserControl.hwnd)
 
            If mBorderStyle = BorderCustom Then
                DrawComboBorder
            End If
            
            RaiseEvent MouseMove(Button, Shift, X, Y)
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ret As Long
    
    If mEnabled And (mBorderStyle <> BorderCustom) Then
        With UserControl
            Call DrawEdge(.hdc, mButtonRect, EDGE_RAISED, BF_RECT)
            .Picture = .Image
        End With
    End If
    
    ' Apro il menu PopUp
    If Button = vbRightButton And mEnabled = True Then
        ret = TrackPopupCmb
        If ret > 0 Then
            RaiseEvent ClickPopUp(ret)
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mAlignment = PropBag.ReadProperty("Alignment", DEF_ALIGNMENT)
    mAutoComplete = PropBag.ReadProperty("AutoComplete", DEF_AUTOCOMPLETE)
    mBackColor = PropBag.ReadProperty("BackColor", DEF_BACKCOLOR)
    mBorderColor = PropBag.ReadProperty("BorderColor", DEF_BORDERCOLOR)
    mBorderCurve = PropBag.ReadProperty("BorderCurve", DEF_BORDERCURVE)
    mBorderStyle = PropBag.ReadProperty("BorderStyle", DEF_BORDERSTYLE)
    mBorderWidth = PropBag.ReadProperty("BorderWidth", DEF_BORDERWIDTH)
    mButtonBackColor = PropBag.ReadProperty("ButtonBackColor", DEF_BUTTONBACKCOLOR)
    mCacheIncrement = PropBag.ReadProperty("CacheIncrement", DEF_CACHE_INCREMENT)
    mColumnHeaders = PropBag.ReadProperty("ColumnHeaders", DEF_COLUMNHEADERS)
    mColumnResize = PropBag.ReadProperty("ColumnResize", DEF_COLUMNRESIZE)
    mColumnSort = PropBag.ReadProperty("ColumnSort", DEF_COLUMNSORT)
    mDefaultItemForeColor = PropBag.ReadProperty("DefaultItemForeColor", DEF_DEFAULTITEMFORECOLOR)
    mDisplayEllipsis = PropBag.ReadProperty("DisplayEllipsis", DEF_EDITABLE)
    mDropDownAutoWidth = PropBag.ReadProperty("DropDownAutoWidth", DEF_DROPDOWNAUTOWIDTH)
    mDropDownItemsVisible = PropBag.ReadProperty("DropDownItemsVisible", DEF_DROPDOWNITEMSVISIBLE)
    mDropDownWidth = PropBag.ReadProperty("DropDownWidth", DEF_DROPDOWNWIDTH)
    mEditable = PropBag.ReadProperty("Editable", DEF_EDITABLE)
    mEnabled = PropBag.ReadProperty("Enabled", DEF_ENABLED)
    mFocusRectColor = PropBag.ReadProperty("FocusRectColor", DEF_FOCUSRECTCOLOR)
    mFocusRectStyle = PropBag.ReadProperty("FocusRectStyle", DEF_FOCUSRECTSTYLE)
    mForeColor = PropBag.ReadProperty("ForeColor", DEF_FORECOLOR)
    mHotBorderColor = PropBag.ReadProperty("HotBorderColor", DEF_BORDERCOLOR)
    mHotButtonBackColor = PropBag.ReadProperty("HotButtonBackColor", DEF_BUTTONBACKCOLOR)
    mIntegralHeight = PropBag.ReadProperty("IntegralHeight", DEF_INTEGRALHEIGHT)
    mLocked = PropBag.ReadProperty("Locked", DEF_LOCKED)
    mMaxLength = PropBag.ReadProperty("MaxLength", 0)
    mPageScrollItems = PropBag.ReadProperty("PageScrollItems", DEF_PAGESCROLLITEMS)
    mRequireCheckedItem = PropBag.ReadProperty("RequireCheckedItem", DEF_REQUIRECHECKEDITEM)
    mRowHeightMin = PropBag.ReadProperty("RowHeightMin", DEF_ROWHEIGHTMIN)
    mScaleUnits = PropBag.ReadProperty("ScaleUnits", DEF_SCALEUNITS)
    mSearchColumn = PropBag.ReadProperty("SearchColumn", DEF_SEARCHCOLUMN)
    mStyle = PropBag.ReadProperty("Style", DEF_STYLE)
    mTextAll = PropBag.ReadProperty("TextAll", DEF_TEXTALL)
    mTextNone = PropBag.ReadProperty("TextNone", DEF_TEXTNONE)
    mTextSelection = PropBag.ReadProperty("TextSelection", DEF_TEXTSELECTION)
    
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set DropDownFont = PropBag.ReadProperty("DropDownFont", Ambient.Font)
    Cols = PropBag.ReadProperty("Cols", DEF_COLS)

    '#############################################################################################################################    'Format Controls
    With txtCombo
        .Alignment = mAlignment
        .BackColor = mBackColor
        .ForeColor = mForeColor
        .MaxLength = mMaxLength
        .Visible = mEditable
    End With
    
    picList.BackColor = mBackColor
    
    With UserControl
        .BackColor = mBackColor
        .ForeColor = mForeColor
    End With
    
    SBLargeChange(efsVertical) = mPageScrollItems
    
    '#############################################################################################################################
    'Subclassing
    If Ambient.UserMode Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        
        With UserControl.Parent
            Call Subclass_Start(.hwnd)
            Call Subclass_AddMsg(.hwnd, WM_WINDOWPOSCHANGING, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_WINDOWPOSCHANGED, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_GETMINMAXINFO, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_LBUTTONDOWN, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_SIZE, MSG_AFTER)
        End With

        With UserControl
            Call Subclass_Start(.hwnd)
            Call Subclass_AddMsg(.hwnd, WM_KILLFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_SETFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEWHEEL, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSELEAVE, MSG_AFTER)
        End With

        With picList
            Call Subclass_Start(.hwnd)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEWHEEL, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEMOVE, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSELEAVE, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEHOVER, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_HSCROLL, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_VSCROLL, MSG_AFTER)
        End With
        
        SBStyle = Style_Regular
        SBCreate picList.hwnd

        SBLargeChange(efsVertical) = 5
        SBSmallChange(efsVertical) = 1
        SBLargeChange(efsHorizontal) = 5
        SBSmallChange(efsHorizontal) = 1
    End If
End Sub

Private Sub UserControl_Resize()
    With txtCombo
        .Left = BORDER_LEFT
        .Top = BORDER_TOP
        .Height = UserControl.ScaleHeight - (BORDER_TOP * 2)
        .Width = (UserControl.ScaleWidth - (BORDER_LEFT + BORDER_TOP)) - BUTTON_WIDTH
    End With
    
    With UserControl
        .Picture = Nothing
    End With
    
    DrawComboBorder
End Sub

Private Sub UserControl_Show()
    Dim lResult As Long
    
    'This modifies the PictureBox control so that it is not bound by
    'its Container
    'Dropdown can render over any Container the control is in
    '(such as a Frame) and is not restricted by the Forms Boundaries
    
    lResult = GetWindowLong(picList.hwnd, GWL_EXSTYLE)
    Call SetWindowLong(picList.hwnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW)
    Call SetWindowPos(picList.hwnd, picList.hwnd, 0, 0, 0, 0, 39)
    Call SetWindowLong(picList.hwnd, -8, Parent.hwnd)
    Call SetParent(picList.hwnd, 0)
End Sub

Private Sub UserControl_Terminate()
    On Local Error GoTo UserControl_TerminateError
    
    Call Subclass_Stop(UserControl.Parent.hwnd)
    Call Subclass_Stop(UserControl.hwnd)
    Call Subclass_Stop(picList.hwnd)
    
    DestroyMenu m_hMenu
  
UserControl_TerminateError:
    Exit Sub
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", mFont, Ambient.Font)
    Call PropBag.WriteProperty("DropDownFont", mDropDownFont, Ambient.Font)
    Call PropBag.WriteProperty("Cols", UBound(mCols) + 1, DEF_COLS)
    
    Call PropBag.WriteProperty("Alignment", mAlignment, DEF_ALIGNMENT)
    Call PropBag.WriteProperty("AutoComplete", mAutoComplete, DEF_AUTOCOMPLETE)
    Call PropBag.WriteProperty("BackColor", mBackColor, DEF_BACKCOLOR)
    Call PropBag.WriteProperty("BorderColor", mBorderColor, DEF_BORDERCOLOR)
    Call PropBag.WriteProperty("BorderCurve", mBorderCurve, DEF_BORDERCURVE)
    Call PropBag.WriteProperty("BorderStyle", mBorderStyle, DEF_BORDERSTYLE)
    Call PropBag.WriteProperty("BorderWidth", mBorderWidth, DEF_BORDERWIDTH)
    Call PropBag.WriteProperty("ButtonBackColor", mButtonBackColor, DEF_BUTTONBACKCOLOR)
    Call PropBag.WriteProperty("CacheIncrement", mCacheIncrement, DEF_CACHE_INCREMENT)
    Call PropBag.WriteProperty("ColumnHeaders", mColumnHeaders, DEF_COLUMNHEADERS)
    Call PropBag.WriteProperty("ColumnResize", mColumnResize, DEF_COLUMNRESIZE)
    Call PropBag.WriteProperty("ColumnSort", mColumnSort, DEF_COLUMNSORT)
    Call PropBag.WriteProperty("DefaultItemForeColor", mDefaultItemForeColor, DEF_DEFAULTITEMFORECOLOR)
    Call PropBag.WriteProperty("DisplayEllipsis", mDisplayEllipsis, DEF_EDITABLE)
    Call PropBag.WriteProperty("DropDownAutoWidth", mDropDownAutoWidth, DEF_DROPDOWNAUTOWIDTH)
    Call PropBag.WriteProperty("DropDownItemsVisible", mDropDownItemsVisible, DEF_DROPDOWNITEMSVISIBLE)
    Call PropBag.WriteProperty("DropDownWidth", mDropDownWidth, DEF_DROPDOWNWIDTH)
    Call PropBag.WriteProperty("Editable", mEditable, DEF_EDITABLE)
    Call PropBag.WriteProperty("Enabled", mEnabled, DEF_ENABLED)
    Call PropBag.WriteProperty("FocusRectColor", mFocusRectColor, DEF_FOCUSRECTCOLOR)
    Call PropBag.WriteProperty("FocusRectStyle", mFocusRectStyle, DEF_FOCUSRECTSTYLE)
    Call PropBag.WriteProperty("ForeColor", mForeColor, DEF_FORECOLOR)
    Call PropBag.WriteProperty("HotBorderColor", mHotBorderColor, DEF_BORDERCOLOR)
    Call PropBag.WriteProperty("HotButtonBackColor", mHotButtonBackColor, DEF_BUTTONBACKCOLOR)
    Call PropBag.WriteProperty("IntegralHeight", mIntegralHeight, DEF_INTEGRALHEIGHT)
    Call PropBag.WriteProperty("Locked", mLocked, DEF_LOCKED)
    Call PropBag.WriteProperty("MaxLength", mMaxLength, 0)
    Call PropBag.WriteProperty("PageScrollItems", mPageScrollItems, DEF_PAGESCROLLITEMS)
    Call PropBag.WriteProperty("RequireCheckedItem", mRequireCheckedItem, DEF_REQUIRECHECKEDITEM)
    Call PropBag.WriteProperty("RowHeightMin", mRowHeightMin, DEF_ROWHEIGHTMIN)
    Call PropBag.WriteProperty("ScaleUnits", mScaleUnits, DEF_SCALEUNITS)
    Call PropBag.WriteProperty("SearchColumn", mSearchColumn, DEF_SEARCHCOLUMN)
    Call PropBag.WriteProperty("Style", mStyle, DEF_STYLE)
    Call PropBag.WriteProperty("TextAll", mTextAll, DEF_TEXTALL)
    Call PropBag.WriteProperty("TextNone", mTextNone, DEF_TEXTNONE)
    Call PropBag.WriteProperty("TextSelection", mTextSelection, DEF_TEXTSELECTION)
End Sub

'=======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
Errs:
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
Errs:
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
On Error GoTo Errs
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
'  If Not bAdd Then
'    Debug.Assert False                                                                  'hWnd not found, programmer error
'  End If
Errs:

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function


Public Property Get CacheIncrement() As Integer
Attribute CacheIncrement.VB_ProcData.VB_Invoke_Property = ";List"
    CacheIncrement = mCacheIncrement
End Property

Public Property Let CacheIncrement(ByVal NewValue As Integer)
    mCacheIncrement = NewValue
    
    PropertyChanged "CacheIncrement"
End Property


Public Function LoadFile(ByVal FullNameFile As String, Optional ByVal LoadCheck As Boolean = False, Optional ByVal CancellaPrima As Boolean = True) As Boolean
    ' Legge i dati da file di testo e li scrive in una ComboBox
    Dim MyString As String
    Dim File1 As Integer
    Dim strTmp() As String
    Dim Che As Boolean
    Dim tmpEditable As Boolean
    Dim nFlag As FlagsEnum
    Const Separatore As String = "|"
    
    On Error Resume Next
    
    tmpEditable = mEditable
    Me.Editable = True
    
    If FileExistsLst(FullNameFile) = False Then
        LoadFile = False
        Exit Function
    End If
    
    If CancellaPrima = True Then Me.Clear
    
    File1 = FreeFile
    
    Open FullNameFile For Input Access Read As #File1
    Do While EOF(File1) = False
        Line Input #File1, MyString
        
        If LoadCheck = False And InStr(1, MyString, "|", vbTextCompare) <> 0 Then
            MyString = Left$(MyString, Len(MyString) - 2)
        End If
        
        If LoadCheck = True And mStyle = Checkboxes Then
            strTmp = Split(MyString, Separatore, , vbTextCompare)
            If UBound(strTmp) > 0 Then Che = CBool(strTmp(1))
            AddItem strTmp(0), , Che
        Else
            AddItem MyString
        End If
    Loop
    Close #File1
    
    LoadFile = True

    Me.Editable = tmpEditable
    
End Function

Public Function SaveFile(FullNameFile As String, Optional SaveCheck As Boolean = False, Optional AvvertiSeFileEsiste As Boolean = True) As Boolean
    ' Legge i dati da una ComboBox e li scrive su di un file di testo
    Dim ret
    Dim lCol As Long
    Dim savelist As Long
    Dim File2 As Integer
    Dim strTmp As String
    Dim nFlag As FlagsEnum
    Const Separatore As String = "|"

    'On Error Resume Next
    
    If Me.ListCount = 0 Then
        SaveFile = False
        Exit Function
    End If
    
    If AvvertiSeFileEsiste = True And (FileExistsLst(FullNameFile) = True) Then
        ret = MsgBox(FullNameFile & " esiste già." & vbNewLine & "Sostituirlo?", vbYesNo + vbExclamation + vbDefaultButton1, App.ProductName)
        If ret = vbNo Then
            SaveFile = False
            Exit Function
        End If
    End If
    
    Me.Enabled = False
    
    File2 = FreeFile

    Open FullNameFile$ For Output As #File2
    
    For savelist& = 0 To mItemCount
        strTmp = mItems(savelist&).sValue(LBound(mCols))
        For lCol = LBound(mCols) + 1 To UBound(mCols)
            strTmp = strTmp & vbTab & mItems(savelist&).sValue(lCol)
        Next
        If SaveCheck = True And mStyle = Checkboxes Then
            strTmp = strTmp & Separatore & mItems(savelist&).nFlags
        End If
        Print #File2, strTmp
    Next
    
    Close #File2
    
    Me.Enabled = True
    
    SaveFile = True
End Function

Private Function FileExistsLst(ByVal sPath As String) As Boolean
  'Determines if a file exists. This function tests the validity of the file and path. It works only on the local file system or on a remote drive that has been mounted to a drive letter.
  'It will return False for remote file paths that begin with the UNC names \\server or \\server\share. It will also return False if a mounted remote drive is out of service.
  'Requires Version 4.71 and later of Shlwapi.dll
   FileExistsLst = PathFileExists(sPath) = 1
   
End Function

