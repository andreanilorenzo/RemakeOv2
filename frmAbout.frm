VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   Caption         =   "About ..."
   ClientHeight    =   6450
   ClientLeft      =   11025
   ClientTop       =   1275
   ClientWidth     =   8175
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "frmAbout"
   Picture         =   "frmAbout.frx":030A
   ScaleHeight     =   430
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSound 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4200
      Top             =   4680
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   4680
   End
   Begin VB.PictureBox pctImg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   1320
      Picture         =   "frmAbout.frx":21FBD
      ScaleHeight     =   1365
      ScaleWidth      =   5295
      TabIndex        =   5
      Top             =   840
      Width           =   5295
      Begin Remakeov2.XpBs cmdContAgg 
         Height          =   495
         Left            =   3960
         TabIndex        =   6
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "&Controlla versione"
         ButtonStyle     =   3
         PictureWidth    =   16
         PictureHeight   =   16
         PictureSize     =   0
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
      End
      Begin VB.Image imgPoiGps 
         Appearance      =   0  'Flat
         Height          =   630
         Left            =   0
         MouseIcon       =   "frmAbout.frx":43C70
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":43DC2
         Top             =   0
         Width           =   2595
      End
      Begin VB.Image imgNokioteca 
         Appearance      =   0  'Flat
         Height          =   540
         Left            =   0
         MouseIcon       =   "frmAbout.frx":49354
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":494A6
         Top             =   720
         Width           =   2595
      End
      Begin VB.Image imgGPSbabel 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   2640
         MouseIcon       =   "frmAbout.frx":4B1A8
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":4B2FA
         Top             =   0
         Width           =   1200
      End
      Begin VB.Image imgVBaccelerator 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3960
         MouseIcon       =   "frmAbout.frx":4BDD1
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":4BF23
         Top             =   600
         Width           =   1275
      End
   End
   Begin VB.TextBox txtLicenza 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmAbout.frx":4C47C
      Top             =   2160
      Width           =   7935
   End
   Begin VB.PictureBox PicShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000A80FF&
      Height          =   1905
      Left            =   240
      Picture         =   "frmAbout.frx":4C487
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   4
      Top             =   4080
      Width           =   7695
   End
   Begin VB.CommandButton cmdEsci 
      BackColor       =   &H00404040&
      Cancel          =   -1  'True
      Caption         =   "&Esci  [Esc]"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblSitoWeb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pagina web del programma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2355
      MouseIcon       =   "frmAbout.frx":6E13A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   120
      Width           =   3285
   End
   Begin VB.Label lblMail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Invia una mail per commenti o suggerimenti!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MouseIcon       =   "frmAbout.frx":6E28C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private m_Media As cMedia

Dim vTop As Long            'Stores the Text Top pos
Dim CrdLines() As String    'Stores the text lines

Dim SpecialThank As String    ' Contiene il testo dei ringraziamenti
Dim MsgPrimoAvvio As String   ' Indica se è il primo avvio del programma. Serve per far visualizzare la licenza
Dim sFileMid As String        ' Il file .mid da suonare

Public Function ApriForm(Optional ByVal frmCaption As String = "", Optional ByVal MsgCaption As String = "") As String
    
    MsgPrimoAvvio = MsgCaption
    Load Me
    If frmCaption <> "" Then Me.Caption = frmCaption
    Me.Show vbModal
    
End Function

Private Sub cmdContAgg_Click()
    frmActProg.Show vbModeless
    Unload Me
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If FormIsLoad("frmMain") = True Then frmMain.Visible = False
    If MsgPrimoAvvio <> "" Then MsgBox "Questa è la prima volta che avvii " & Versione & "." & vbNewLine & "Prima di proseguire è necessario che leggi la licenza di utilizzo.", vbExclamation, App.ProductName
End Sub

Sub Form_Load()
    Dim B() As Byte
    Dim fNum As Integer
    
    ' Inizializzo la classe per il file .mid
    Set m_Media = New cMedia

    On Error GoTo Continua
    B = LoadResData(110, "SOUND")
    sFileMid = App.path & "\Sound.mid"
    fNum = FreeFile
    Open sFileMid For Binary Access Write Lock Read As #fNum
    Put #fNum, , B
Continua:
    Close #fNum
    DoEvents
    
    On Error Resume Next
    
    ' Take control of message processing by installing our message handling routine into the chain of message routines for this window
    ' Per la limitazione delle misure della form
    g_nProcOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)

    SpecialThank = "RemakeOv2......... dal 06/12/2005"
    SpecialThank = SpecialThank & vbNewLine
    SpecialThank = SpecialThank & vbNewLine & "Special thanks to:" & vbNewLine & "www.poigps.com"
    SpecialThank = SpecialThank & vbNewLine & "www.nokioteca.net"
    SpecialThank = SpecialThank & vbNewLine & "www.gpsbabel.org"
    SpecialThank = SpecialThank & vbNewLine & "areagps.135.it"
    SpecialThank = SpecialThank & vbNewLine
    SpecialThank = SpecialThank & vbNewLine & "emme @poigps.com"
    SpecialThank = SpecialThank & vbNewLine
    SpecialThank = SpecialThank & vbNewLine & ".....ed in ordine casuale:" & vbNewLine
    SpecialThank = SpecialThank & vbNewLine & "Zzed @poigps.com"
    SpecialThank = SpecialThank & vbNewLine & "electrobose @nokioteca.net"
    SpecialThank = SpecialThank & vbNewLine & "Galileo @poigps.com"
    SpecialThank = SpecialThank & vbNewLine & "ricfranz @poigps.com"
    SpecialThank = SpecialThank & vbNewLine & "giangio1986 @nokioteca.net"
    SpecialThank = SpecialThank & vbNewLine & "Iulo @nokioteca.net"
    SpecialThank = SpecialThank & vbNewLine & "Iscio @nokioteca.net"
    SpecialThank = SpecialThank & vbNewLine & "giando_ing @poigps.net"
    SpecialThank = SpecialThank & vbNewLine & "nicunic @poigps.com"
    SpecialThank = SpecialThank & vbNewLine & vbNewLine & "....e a tutti quanti hanno contribuito alla crescita del programma...."

    If FormIsLoad("frmMain") = True Then frmMain.Visible = False
    If PrimoAvvio = True Then
        cmdContAgg.Enabled = False
    End If
    Me.BackColor = vbWhite
    PicShow.BackColor = Me.BackColor
    txtLicenza.Text = Versione
    txtLicenza.Text = txtLicenza.Text & vbNewLine & vbNewLine & Licenza
    txtLicenza.ZOrder
    
    tmrUpdate.Interval = 55
    vTop = PicShow.Height
    CrdLines = Split(SpecialThank, vbCrLf) 'Getting into lines
   
    Form_Resize
    
    tmrSound.Enabled = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ElaboraLabelLink
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If MsgPrimoAvvio <> "" Then
        Dim ret
        ret = MsgBox(MsgPrimoAvvio, vbInformation + vbYesNoCancel, App.ProductName)
        
        If ret = vbNo Then
            Kill (XmlFileConfig)
            clsManifestFile.DeleteManifest
            DoEvents
            ChiudiProgramma
        ElseIf ret = vbCancel Then
            Cancel = True
            txtLicenza.SetFocus
        End If

    End If
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tmrUpdate.Enabled = False
    
    ' Make sure the form is not minimized
    If Me.WindowState <> vbMinimized Then
        ' Maintain a minimum height and width in order to not set a negative width or height
        'If Me.Height < 6700 Or Me.Width < 8200 Then
            'If Me.Height < 6700 Then Me.Height = 6700
            'If Me.Width < 8200 Then Me.Width = 8200
        'Else
            ' Centro i controlli nella form
            lblMail.Move (Me.ScaleWidth - lblMail.Width) / 2
            lblSitoWeb.Move (Me.ScaleWidth - lblSitoWeb.Width) / 2
            pctImg.Move (Me.ScaleWidth - pctImg.Width) / 2
            
            txtLicenza.Move 0, txtLicenza.Top, Me.ScaleWidth, CInt(Me.Height / 44)
            PicShow.Move 0, txtLicenza.Top + txtLicenza.Height, Me.ScaleWidth, Me.ScaleHeight - PicShow.Top
        'End If
    End If
    
    tmrUpdate.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If FormIsLoad("frmMain") = True And FormIsLoad("frmActProg") = False Then frmMain.Visible = True
    ' give message processing control back to VB if you don't do this you WILL crash!!!
    Call SetWindowLong(hwnd, GWL_WNDPROC, g_nProcOld)
    
    StopPlayMid
    Set m_Media = Nothing
    
End Sub

Private Sub StopPlayMid()
    
    With m_Media
        If LCase(.Status) = "playing" Then
            .mmStop
            .mmClose
            Kill sFileMid
        End If
    End With

End Sub

Private Sub imgNokioteca_Click()
    Screen.MousePointer = vbDefault
    DoEvents
    ShellExecute hwnd, "open", "http://www.nokioteca.net/home/forum/index.php?showtopic=916", vbNullString, vbNullString, SW_SHOW
End Sub
Private Sub imgNokioteca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgNokioteca.ToolTipText = "http://www.nokioteca.net"
End Sub

Private Sub lblMail_Click()
    Screen.MousePointer = vbDefault
    DoEvents
    ShellExecute hwnd, "open", "mailto:remakeov2@poigps.com", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub lblMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ElaboraLabelLink(lblMail)
    lblMail.ToolTipText = "mailto:remakeov2@poigps.com"
End Sub

Private Sub lblSitoWeb_Click()
    Screen.MousePointer = vbDefault
    DoEvents
    ShellExecute hwnd, "open", HomePage, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub lblSitoWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ElaboraLabelLink(lblSitoWeb)
    lblSitoWeb.ToolTipText = HomePage
End Sub

Private Sub imgPoiGps_Click()
    Screen.MousePointer = vbDefault
    DoEvents
    ShellExecute hwnd, "open", "http://www.poigps.com/modules.php?name=Forums&file=viewtopic&t=11153", vbNullString, vbNullString, SW_SHOW
End Sub
Private Sub imgPoiGps_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPoiGps.ToolTipText = "http://www.poigps.com"
End Sub

Private Sub imgGPSbabel_Click()
    Screen.MousePointer = vbDefault
    DoEvents
    ShellExecute hwnd, "open", "http://www.gpsbabel.org/index.html", vbNullString, vbNullString, SW_SHOW
End Sub
Private Sub imgGPSbabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgGPSbabel.ToolTipText = "http://www.gpsbabel.org"
End Sub

Private Sub imgVBaccelerator_Click()
    Screen.MousePointer = vbDefault
    DoEvents
    ShellExecute hwnd, "open", "http://www.vbaccelerator.com", vbNullString, vbNullString, SW_SHOW
End Sub
Private Sub imgVBaccelerator_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgVBaccelerator.ToolTipText = "http://www.vbaccelerator.com"
End Sub

Private Sub ElaboraLabelLink(Optional LabelLink As Object)
    On Error Resume Next

    Dim ctl As Control
    
    Screen.MousePointer = vbDefault
    
    For Each ctl In Controls
        If TypeOf ctl Is Label Then
            If ctl.Name <> LabelLink.Name Then
                ctl.ForeColor = vbBlack
            Else
                ctl.ForeColor = vbBlue
            End If
            
        End If
    Next
    
    DoEvents

End Sub

Private Sub pctImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = vbDefault
End Sub

Private Sub picShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrUpdate.Enabled = False
End Sub

Private Sub picShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrUpdate.Enabled = True
End Sub

Private Sub tmrSound_Timer()
    
    tmrSound.Enabled = False
    
    With m_Media
        .mmClose
        .mmOpen sFileMid
        .mmPlay
        'tmrSound.Enabled = True
    End With

End Sub

Private Sub tmrUpdate_Timer()
    Dim X As Integer
    Dim nTop As Long
    
    PicShow.Cls
    nTop = vTop
    
    For X = 0 To UBound(CrdLines)
        'if the 'top' is inside the picturebox then draw
        If nTop > -50 And nTop < PicShow.Height Then
            SendCredits PicShow, CrdLines(X), (PicShow.ScaleWidth - PicShow.TextWidth(CrdLines(X))) / 2, nTop, PicShow.BackColor, RGB(0, 0, 0), PicShow.BackColor, 1 / 2
        End If
        nTop = nTop + PicShow.TextHeight(CrdLines(X))
    Next X
    
    'Reloading at the end of the file
    If vTop + 20 < -PicShow.TextHeight("A") * UBound(CrdLines) Then
        vTop = PicShow.Height
    End If
    
    vTop = vTop - 1
    
End Sub

Private Function SendCredits(PicBox As PictureBox, txt As String, ByVal X As Integer, ByVal Y As Integer, Optional StartCol As Long = 0, Optional MidCol As Long = 111111, Optional EndCol As Long = 0, Optional ByVal cRegion As Double)
    Dim hLength   As Integer 'Region over which the text fades
    Dim DrawCol   As Long    'The current faded color
    Dim rctDraw   As RECT
    
    hLength = PicBox.Height * cRegion   'Determines the fade region
    
    If Y <= hLength And Y >= -50 Then   'Some Calculations
        DrawCol = GetShade(MidCol, EndCol, (hLength - Y) / (hLength + 20))   'Getting the shaded color
    ElseIf Y <= PicBox.Height And Y >= PicBox.Height * (1 - cRegion) Then
        DrawCol = GetShade(StartCol, MidCol, (PicBox.Height - Y) / hLength)  'Getting the shaded color
    Else
        DrawCol = MidCol
    End If
    
    With rctDraw
        .Left = X
        .Top = Y
        .Right = PicBox.Width
        .Bottom = PicBox.Height
    End With
    
    PicBox.ForeColor = DrawCol  'Setting the DrawColor
    DrawText PicBox.hdc, txt, -1, rctDraw, &H800    'Drawing the text
    
End Function

Private Function GetShade(ByVal StartCol As Long, ByVal EndCol As Long, ByVal ColDepth As Double) As Long
    'Returns the shaded color in the specified color depth
    On Error Resume Next
    Dim sRate As Double
    Dim cBlue As Long, cGreen As Long, cRed As Long   'Determines the pixel color
    Dim sBlue As Long, sGreen As Long, sRed As Long   'Determines the SHADING color
    
    sRate = ColDepth
    GetRGB EndCol, sRed, sGreen, sBlue
    GetRGB StartCol, cRed, cGreen, cBlue
    cRed = cRed + (sRed - cRed) * sRate
    cGreen = cGreen + (sGreen - cGreen) * sRate
    cBlue = cBlue + (sBlue - cBlue) * sRate
    If cRed < 0 Then cRed = -cRed
    If cGreen < 0 Then cGreen = -cGreen
    If cBlue < 0 Then cBlue = -cBlue
    GetShade = RGB(cRed, cGreen, cBlue)

End Function

Private Sub GetRGB(ByVal LngCol As Long, r As Long, G As Long, B As Long)
    'Returns the RGB values
    r = LngCol Mod 256
    G = (LngCol And vbGreen) / 256 'Green
    B = (LngCol And vbBlue) / 65536 'Blue
    
End Sub

Private Sub txtLicenza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ElaboraLabelLink
End Sub
