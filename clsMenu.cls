VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class Name: clsMenu
' Version: 1.0
' Author: Walter Wong (walterwong@msn.com)
' Description:
'   This class can build a hierarchical popup menu
'   Features:
'     o Hierarchical structure
'     o Each item has
' History:
'   2002/5/27:
'     o Removed dependencies on win.tlb, now all needed API declarations included, this means
'       this class can be used alone.
'     o Added a full functional test project to Source+ 2000.

Private m_hMenu     As Long
Public Caption As String ' If works as a sub-menu, this would be add as parent menu item

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
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTL) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0

Private Type POINTL
    X As Long
    Y As Long
End Type
Private Const TPM_RETURNCMD = &H100&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_LEFTBUTTON = &H0&

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Sub Class_Initialize()
    m_hMenu = CreatePopupMenu()
End Sub

Private Sub Class_Terminate()
    DestroyMenu m_hMenu
End Sub

Friend Property Get hMenu() As Long
    hMenu = m_hMenu
End Property

Public Sub AddItem(ByVal nID As Long, varItem As Variant, Optional ByVal bDefault As Boolean = False, _
    Optional ByVal bChecked As Boolean = False, Optional ByVal bDisabled As Boolean = False, _
    Optional ByVal bGrayed As Boolean = False, Optional ByVal bNewColumn As Boolean = False)
    
    If TypeName(varItem) = "String" Then
        If varItem = "-" Then
            AppendMenu m_hMenu, MF_STRING Or MF_SEPARATOR, nID, ByVal vbNullString
        Else
            AppendMenu m_hMenu, MF_STRING Or IIf(bNewColumn, MF_MENUBARBREAK, 0) Or IIf(bChecked, MF_CHECKED, 0), nID, ByVal varItem
        End If
    ElseIf TypeOf varItem Is clsMenu Then
        Dim cSubMenu As clsMenu
        Set cSubMenu = varItem
        AppendMenu m_hMenu, MF_STRING Or MF_POPUP Or IIf(bNewColumn, MF_MENUBARBREAK, 0), cSubMenu.hMenu, ByVal cSubMenu.Caption
    End If
    
    If bDefault Then SetMenuDefaultItem m_hMenu, nID, APIFALSE
    If bGrayed Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_GRAYED
    If bDisabled Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_DISABLED
    
End Sub

Public Function TrackPopup(Optional ByVal hwnd As Long = -1, Optional ByVal PopX As Long = -1, Optional ByVal PopY As Long = -1, Optional ByVal hWndOfBeneathControl As Long = -1) As Long
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
        
        If hChild = 0 Then Err.Raise -1, "clsMenu.TrackPopup", "Cannot find top window of current process!"
        h = hChild
    Else
        h = hwnd
    End If
    
    If hWndOfBeneathControl <> -1 Then
        Dim rt As RECT
        GetWindowRect hWndOfBeneathControl, rt
        X = rt.Left
        Y = rt.Bottom
    Else
        Dim pt As POINTL
        GetCursorPos pt
        If PopX = -1 Then X = pt.X Else: X = PopX
        If PopY = -1 Then Y = pt.Y Else: Y = PopY
    End If
    
    TrackPopup = TrackPopupMenu(m_hMenu, TPM_RETURNCMD + TPM_LEFTALIGN + TPM_LEFTBUTTON, X, Y, 0, h, 0)
End Function

