VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "frmMain"
   ClientHeight    =   2865
   ClientLeft      =   5490
   ClientTop       =   5715
   ClientWidth     =   9690
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   2865
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRinomina 
      Caption         =   "Rinomi&na File"
      Height          =   470
      Left            =   8040
      OLEDropMode     =   1  'Manual
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "   Aggiorna POI    sul &Telefono"
      Height          =   855
      Left            =   6480
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdImpostazioni 
      Caption         =   "&Impostazioni"
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      ToolTipText     =   $"frmMain.frx":86B0
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdAggiorna 
      Caption         =   "&Aggiorna file"
      Height          =   375
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdRimuoviDuplicati 
      Caption         =   "Rimuovi &duplicati"
      Height          =   375
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "A&bout ..."
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdDownloadManager 
      Caption         =   "&Download manager"
      Height          =   470
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "&Esci  [Esc]"
      Height          =   615
      Left            =   8040
      TabIndex        =   5
      ToolTipText     =   " Esce dal programma "
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdConverti 
      Caption         =   "Ri&para e  converti file"
      Height          =   855
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdRegistraEstensioni 
      Caption         =   "&Reg. estens."
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      ToolTipText     =   $"frmMain.frx":8745
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdVerifica 
      Caption         =   "Verifica &file"
      Height          =   855
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreaModifica 
      Caption         =   "&Crea e modifica"
      Height          =   470
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblGestErr 
      BackStyle       =   0  'Transparent
      Caption         =   "ATTENZIONE: GestioneErrori disattivato"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblDebugMode 
      BackStyle       =   0  'Transparent
      Caption         =   "ATTENZIONE: DebugMode attivato"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape shBordo 
      BackColor       =   &H80000002&
      BorderColor     =   &H80000002&
      BorderWidth     =   4
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   735
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --------------------------------------------------------------------------------------------
'
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
' --------------------------------------------------------------------------------------------

Private PosizioneSplashLeft As Long
Private PosizioneSplashTop As Long

Private WithEvents FormSys As frmSysTray
Attribute FormSys.VB_VarHelpID = -1

Private Sub cmdAbout_Click()
    Load frmAbout
    DoEvents
    frmAbout.Show
End Sub

Private Sub cmdAggiorna_Click()
    Load frmDownload
    Load frmUpdateFile
    DoEvents
    frmUpdateFile.Show
End Sub

Private Sub cmdConverti_Click()
    Load frmRemakeov2
    DoEvents
    frmRemakeov2.Show
End Sub

Private Sub cmdCreaModifica_Click()
    Load frmWeb
    DoEvents
    frmWeb.Show
End Sub

Private Sub cmdCreaModifica_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    If data.GetFormat(vbCFFiles) Then
        For i = 1 To 1 'Data.Files.count
            CommandLineFile = (data.Files(i))
            frmWeb.Show
            frmWeb.ControllaModalità ("Edit")
        Next
    End If
    
End Sub

Private Sub cmdDownloadManager_Click()
    Load frmDownload
    DoEvents
    frmDownload.Show
End Sub

Private Sub cmdDownloadManager_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    If data.GetFormat(vbCFFiles) Then
        For i = 1 To 1 'Data.Files.count
            CommandLineFile = (data.Files(i))
            frmDownload.Show
        Next
    End If
    
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdImpostazioni_Click()
    Load frmImpostazioni
    DoEvents
    frmImpostazioni.Show
End Sub

Private Sub cmdRegistraEstensioni_Click()
    Load frmRegEstensioni
    DoEvents
    frmRegEstensioni.Show
End Sub

Private Sub cmdRimuoviDuplicati_Click()
    Load frmRimuoviDuplicati
    DoEvents
    frmRimuoviDuplicati.Show
End Sub

Private Sub cmdRimuoviDuplicati_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    If data.GetFormat(vbCFFiles) Then
        For i = 1 To 1 'Data.Files.count
            frmRimuoviDuplicati.Show
            frmRimuoviDuplicati.ClickPopUp 20, (data.Files(i))
        Next
    End If
    
End Sub

Private Sub cmdRinomina_Click()
    Load frmRinomina
    DoEvents
    frmRinomina.Show
End Sub

Private Sub cmdVerifica_Click()
    On Error Resume Next
    frmCheckOV2.Show
    DoEvents
End Sub

Private Sub cmdVerifica_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    If data.GetFormat(vbCFFiles) Then
        For i = 1 To 1 'Data.Files.count
            PercorsoFileList = (data.Files(i))
            frmCheckOV2.Show
        Next
    End If
    
End Sub

Private Sub command1_Click()
    Load frmTrasmettiFile
    DoEvents
    frmTrasmettiFile.Show
End Sub

Private Sub Form_Initialize()
    Const largBordo As Integer = 5
    
    With shBordo
        .BorderWidth = largBordo
        .Move largBordo, largBordo, Me.ScaleWidth, Me.ScaleHeight
    End With

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SpostaForm
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    If FormIsLoad("frmAbout") = True Then
        frmAbout.SetFocus
    
    ElseIf FormIsLoad("frmWeb") = True Then
        frmWeb.SetFocus
    
    ElseIf FormIsLoad("frmDownload") = True Then
        If FormIsLoad("frmUpdateFile") = True Then
            frmUpdateFile.SetFocus
        Else
            frmDownload.SetFocus
        End If
    
    ElseIf FormIsLoad("frmCheckOV2") = True Then
        frmCheckOV2.SetFocus
    
    ElseIf FormIsLoad("frmRimuoviDuplicati") = True Then
        frmRimuoviDuplicati.SetFocus
    
    ElseIf FormIsLoad("frmRemakeov2") = True Then
        frmRemakeov2.SetFocus
    
    ElseIf FormIsLoad("frmTrasmettiFile") = True Then
        frmTrasmettiFile.SetFocus
    
    ElseIf FormIsLoad("frmImpostazioni") = True Then
        frmImpostazioni.SetFocus
    
    ElseIf FormIsLoad("frmRegEstensioni") = True Then
        frmRegEstensioni.SetFocus

    ElseIf FormIsLoad("frmActProg") = True Then
        frmActProg.SetFocus

    End If
    
    If Var(DebugMode).Valore = 1 Then
        lblDebugMode.Visible = True
    Else
        lblDebugMode.Visible = False
    End If
    If Var(GestioneErrori).Valore = 1 Then
        lblGestErr.Visible = True
    Else
        lblGestErr.Visible = False
    End If

End Sub

Private Sub Form_Load()
    Dim arrPos
    
    Me.Caption = App.ProductName
    lblVersion.Caption = Versione
    IconaInSysTry (True)

    If Var(MainLastPos).Valore <> "" Then
        arrPos = Split(Var(MainLastPos).Valore, ",")
        ' Sposto Form nella posizione salvata
        Me.Left = CLng(arrPos(0))
        Me.Top = CLng(arrPos(1))
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SpostaForm
End Sub

Private Sub SpostaForm()
    
    ' Queste due righe permettono di sopostare la form senza utilizzare la barra del titolo
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

    Select Case Me.Left
        Case Is > Screen.Width - Me.Width
            PosizioneSplashLeft = Screen.Width - Me.Width
        Case Is > 0
            PosizioneSplashLeft = Me.Left
        Case Is < 0
            PosizioneSplashLeft = "0"
    End Select
    
    Select Case Me.Top
        Case Is > Screen.Height - Me.Height
            PosizioneSplashTop = Screen.Height - Me.Height
        Case Is > 0
            PosizioneSplashTop = Me.Top
        Case Is < 0
            PosizioneSplashTop = "0"
    End Select
    
    Me.Move PosizioneSplashLeft, PosizioneSplashTop
    Me.Refresh

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Not FormSys Is Nothing Then
        FormSys.MeQueryUnload Me, Cancel, UnloadMode
    End If
    
    Cancel = ConfermaChiusuraForm("1")
    
End Sub

Private Sub Form_DblClick()
    Me.WindowState = vbMinimized
    RiduciAdIcona
End Sub

Private Sub Form_Resize()
    RiduciAdIcona
End Sub

Private Sub Form_Unload(Cancel As Integer)
     
    If Not FormSys Is Nothing Then
        Unload FormSys
    End If
    Set FormSys = Nothing
    DoEvents

    lVar(MainLastPos) = Me.Left & "," & Me.Top

    ChiudiProgramma
    
End Sub

Private Sub RiduciAdIcona()
    If (Me.WindowState <> vbNormal) And (Not FormSys Is Nothing) Then
        FormSys.MeResize Me
    End If
End Sub

Private Sub IconaInSysTry(Optional Inserisci As Boolean)
    
    If Inserisci = True Then
        Set FormSys = New frmSysTray
        With FormSys
            .AddMenuItem "&About...", "about"
            '.AddMenuItem "&Restore", "restore"
            .AddMenuItem "-"
            '.AddMenuItem "&Converti File", "conv"
            '.AddMenuItem "&Verifica File", "veri"
            '.AddMenuItem "&Crea e Modifica File", "web"
            '.AddMenuItem "-"
            .AddMenuItem "&Apri", "apri", True
            .AddMenuItem "-"
            .AddMenuItem "&Esci", "esci"
            .ToolTip = "RemakeOv2! Programma per i POI dei navigatori GPS"
        End With
    Else
        If Not FormSys Is Nothing Then Unload FormSys
        Set FormSys = Nothing
    End If

End Sub

Private Sub FormSys_MenuClick(ByVal lIndex As Long, ByVal sKey As String)

    Select Case sKey
        Case "apri"
           Me.Show
           Me.WindowState = vbNormal
           Me.ZOrder
        Case "esci"
           Unload Me
        Case "restore"
            'FormSys.Restore Me
        Case "veri"
            frmCheckOV2.Show
        Case "web"
            frmWeb.Show
        Case "conv"
            frmRemakeov2.Show
        Case "about"
           cmdAbout_Click
    End Select
    
End Sub

Private Sub FormSys_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    On Error Resume Next
    
    Me.WindowState = vbNormal
    Me.Show
    Me.ZOrder
    
End Sub

Private Sub FormSys_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)

    If (eButton = vbRightButton) Then
        FormSys.ShowMenu
    End If
    
End Sub

Private Sub lblVersion_DblClick()
    Me.WindowState = vbMinimized
    RiduciAdIcona
End Sub

Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SpostaForm
End Sub
