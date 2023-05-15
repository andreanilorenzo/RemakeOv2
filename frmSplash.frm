VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2250
   ClientLeft      =   6405
   ClientTop       =   5415
   ClientWidth     =   6570
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmSplash.frx":030A
   ScaleHeight     =   2250
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2880
      Top             =   1560
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caricamento in corso Attendere........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------------------------------------------
' Per la funzione GrabScreen
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const SRCCOPY = &HCC0020
'-----------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------
' Per la funzione CreateFormRegion
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Const RGN_COPY = 5
Private ResultRegion As Long
'-----------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------
' Per spostare la form con il mouse
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'-----------------------------------------------------------------------------------------------------------------

Private siEsce As Boolean

Private Sub Form_Initialize()
    lblInfo.Move (Me.Width - lblInfo.Width) / 2, (Me.Height - lblInfo.Height) / 2
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3
End Sub

Private Sub Form_Load()
    Dim nRet As Long
    
    siEsce = False
    Timer1.Enabled = False
    DoEvents

    'If this line are modified or moved a second copy of them may be added again if the form is later Modified by VBSFC.
    nRet = SetWindowRgn(Me.hwnd, CreateFormRegion(1, 1, 0, 0), True)

    SetOnTop frmSplash
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Queste due righe permettono di spostare la form senza utilizzare la barra del titolo
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If siEsce = False Then
        Cancel = True
        siEsce = True
        Timer1.Enabled = True
    Else
        Cancel = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    DeleteObject ResultRegion
    'If the above line is modified or moved a second copy of it may be added again if the form is later Modified by VBSFC.

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    DoEvents
    Unload Me
End Sub

Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long
    Dim STPPX As Integer, STPPY As Integer
    STPPX = Screen.TwipsPerPixelX
    STPPY = Screen.TwipsPerPixelY
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)

    ' Shaped Form Region Definition 2,0,0,442,146,25,25,1
    ObjectRegion = CreateRoundRectRgn(0 * ScaleX * 15 / STPPX + OffsetX, 0 * ScaleY * 15 / STPPY + OffsetY, 442 * ScaleX * 15 / STPPX + OffsetX, 146 * ScaleY * 15 / STPPY + OffsetY, 50 * ScaleX * 15 / STPPX, 50 * ScaleY * 15 / STPPY)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
    
End Function
