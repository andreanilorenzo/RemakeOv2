VERSION 5.00
Object = "{551ACF74-6F53-451E-ABA8-F2534B381A63}#1.0#0"; "andrMap.ocx"
Begin VB.Form frmImmagine 
   Caption         =   "frmImmagine"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   Icon            =   "frmImmagine.frx":0000
   LinkTopic       =   "frmImmagine"
   ScaleHeight     =   7215
   ScaleWidth      =   9045
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin andrMap.andrMapCtl andrMapCtl1 
      Height          =   5295
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9340
   End
   Begin VB.TextBox txtInfo 
      Height          =   2325
      Left            =   4920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmImmagine.frx":030A
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox picIntestazione 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   600
      ScaleHeight     =   735
      ScaleWidth      =   7095
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdEsci 
         Cancel          =   -1  'True
         Caption         =   "&Esci  [Esc]"
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame Frame 
         Caption         =   " Elenco mappe disponibili"
         Height          =   735
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4815
         Begin VB.PictureBox pctFr1 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   4545
            TabIndex        =   2
            Top             =   240
            Width           =   4575
            Begin VB.CommandButton cmdInfo 
               Caption         =   "I"
               Enabled         =   0   'False
               Height          =   315
               Left            =   4200
               TabIndex        =   5
               Top             =   0
               Width           =   315
            End
            Begin VB.ComboBox cmbMappe 
               Height          =   315
               Left            =   0
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   0
               Width           =   4095
            End
         End
      End
   End
End
Attribute VB_Name = "frmImmagine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdInfo_Click()
    Static bSt As Boolean
    
    bSt = Not bSt
    txtInfo.Visible = bSt
    txtInfo.ZOrder
    If txtInfo.Visible = True Then txtInfo.SetFocus

End Sub

Private Sub Form_Activate()
    cmbMappe.SetFocus
End Sub

Private Sub Form_Load()
    
    Me.Caption = "Gestione mappe"
    pctFr1.BorderStyle = 0
    
    With andrMapCtl1
        .PanActive = True
        .UseQuickBar = True
        .AllowZoomIn = True
        .AllowZoomOut = True
        .MouseTrack = False
    End With
    
    Call CaricaMappe
    
    Form_Resize

    SetOnTop frmWeb, True
    
End Sub

Private Sub CaricaMappe()
    Dim arrFile() As String
    Dim cnt As Integer
    Dim sEstensioni As String
    
    sEstensioni = ".gif"
    
    If GetFileInFolder(arrFile, Var(CartellaMappe).Valore, sEstensioni) = True Then
        For cnt = 0 To UBound(arrFile)
            cmbMappe.AddItem arrFile(cnt)
        Next
        cmbMappe.ListIndex = 0
        
    Else
        MsgBox "Non sono state trovate mappe nella cartella" & vbNewLine & Var(CartellaMappe).Valore & vbNewLine & vbNewLine & "I formati supportati sono i seguenti: " & sEstensioni, vbInformation, App.ProductName
    End If
    
End Sub

Private Sub cmbMappe_Click()
    Dim strTmp As String
    
    andrMapCtl1.LoadImage Var(CartellaMappe).Valore & "\" & cmbMappe.List(cmbMappe.ListIndex)
    strTmp = andrMapCtl1.LeggiFileCalibrazione(andrMapCtl1.FileName & ".inf")
    If andrMapCtl1.Calibrata = False Then
        ' Message box a tempo
        SetTimer Me.hwnd, NV_CLOSEMSGBOX, 5000&, AddressOf TimerProc
        Call MessageBox(Me.hwnd, strTmp, App.ProductName & " - Chiusura a tempo", MB_ICONQUESTION Or MB_TASKMODAL)
    End If

End Sub

Private Sub CaricaInfoMappa(ByVal sNomeMappa As String)

    'sFileInf = Var(CartellaMappe).Valore & "/" & sNomeMappa & ".inf"
    
    'If FileExists(sFileInf) = True Then
        'sInfo = LeggiFile(sFileInf)
        
        ' Leggo i dati dal file
        'InfMappa.BackColor = ConvertDelphiColor(GetValue(sInfo, "BackColor"))
        'InfMappa.MapCheck = CLng(GetValue(sInfo, "MapCheck"))
        
        
        'andrMapCtl1.BackColor = InfMappa.BackColor
        
        'txtInfo.Text = sInfo
        'cmdInfo.Enabled = True
        
    'Else
        'andrMapCtl1.BackColor = Me.BackColor
        'cmdInfo.Enabled = False
        'MsgBox "Nessun file di configurazione valido trovato per questa mappa", vbInformation, App.ProductName
    'End If
    
End Sub

Private Sub Form_Resize()
    Dim lF As Integer
    Dim tp As Integer
    
    On Error Resume Next
    
    ' Store the position if the window state is Max'd or Min'd.
    ' Do it before resizing, since a restore will return it to this size.
    If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then
        andrMapCtl1.StorePosition
    End If
    
    picIntestazione.Move (Me.ScaleWidth - picIntestazione.Width) / 2
    lF = picIntestazione.Left + pctFr1.Left + cmdInfo.Left
    tp = picIntestazione.Top + pctFr1.Top + cmdInfo.Height
    txtInfo.Move lF, tp
    
    andrMapCtl1.Move 0, picIntestazione.Top + picIntestazione.Height + 60, Me.ScaleWidth, Me.ScaleHeight - andrMapCtl1.Top
    
    ' Recall the position. Will only work on Restore.
    andrMapCtl1.RecallPosition
        
End Sub

Private Sub andrMapCtl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Frame.Caption = X & " -- " & Y
End Sub

Private Sub andrMapCtl1_ZoomChanged(ByVal ZoomPercent As Long)
    ' Toggle quickbar buttons based upon current percentage
    andrMapCtl1.AllowZoomIn = (ZoomPercent < 1000)
    andrMapCtl1.AllowZoomOut = (ZoomPercent > 10)
End Sub

Private Sub andrMapCtl1_ZoomInClick()
    If andrMapCtl1.Zoom < 1000 Then
        andrMapCtl1.Zoom = andrMapCtl1.Zoom + 10
    End If
End Sub

Private Sub andrMapCtl1_ZoomOutClick()
    If andrMapCtl1.Zoom > 10 Then
        andrMapCtl1.Zoom = andrMapCtl1.Zoom - 10
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetOnTop frmWeb, False
    DoEvents
End Sub

Private Sub txtInfo_LostFocus()
    
    txtInfo.Visible = False

End Sub
