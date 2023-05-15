VERSION 5.00
Begin VB.Form frmOperazioni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operazioni"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmOperazioni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8565
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalva 
      Caption         =   "&Salva Impostazioni"
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      ToolTipText     =   "Salva l'attuale impostazione dei file in ingresso/uscita"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&Ok  [Esc]"
      Default         =   -1  'True
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox lstUscita2 
      Height          =   2010
      Left            =   5040
      TabIndex        =   6
      Top             =   2880
      Width           =   3375
   End
   Begin VB.ListBox lstUscita1 
      Height          =   2010
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   4695
   End
   Begin VB.ListBox lstIngresso2 
      Height          =   1620
      Left            =   5040
      TabIndex        =   4
      Top             =   840
      Width           =   3375
   End
   Begin VB.ListBox lstIngresso1 
      Height          =   1620
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.ListBox lstUscita0 
      Height          =   2010
      ItemData        =   "frmOperazioni.frx":030A
      Left            =   120
      List            =   "frmOperazioni.frx":0311
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox lstIngresso0 
      Height          =   1620
      ItemData        =   "frmOperazioni.frx":0321
      Left            =   120
      List            =   "frmOperazioni.frx":0328
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdProgramma 
      Caption         =   "&GPSBabel   (clicca per cambiare modalità)"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton cmdProgramma 
      Caption         =   "&TomTom Tools    (clicca per cambiare modalità)"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label lblUscita 
      Alignment       =   2  'Center
      Caption         =   "File Uscita"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   8295
   End
   Begin VB.Label lblIngresso 
      Alignment       =   2  'Center
      Caption         =   "File Ingresso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   8295
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   8295
   End
End
Attribute VB_Name = "frmOperazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LBHSi0 As clsListBox
Private LBHSi1 As clsListBox
Private LBHSi2 As clsListBox
Private LBHSu0 As clsListBox
Private LBHSu1 As clsListBox
Private LBHSu2 As clsListBox

Private WithEvents objDOS As DOSOutputs
Attribute objDOS.VB_VarHelpID = -1
Private sRetDOScomm As String

Public PreparaLabel As Boolean

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdProgramma_Click(index As Integer)
    
    Select Case index
        Case Is = 0
            Call CaricaArrayGpsBabel
            Call CaricaListBoxGpsBabel("file")
        Case Is = 1
            Call CaricaListBoxTomTom
    End Select
    
    Call CambiaStato
    
End Sub

Private Sub Form_Load()
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    ' Inizializzo la classe per la ListBox
    Set LBHSi0 = New clsListBox
    Set LBHSi1 = New clsListBox
    Set LBHSi2 = New clsListBox
    LBHSi0.Attach lstIngresso0
    LBHSi1.Attach lstIngresso1
    LBHSi2.Attach lstIngresso2

    Set LBHSu0 = New clsListBox
    Set LBHSu1 = New clsListBox
    Set LBHSu2 = New clsListBox
    LBHSu0.Attach lstUscita0
    LBHSu1.Attach lstUscita1
    LBHSu2.Attach lstUscita2
    
    Set objDOS = New DOSOutputs
    
    If PreparaLabel = True Then
        lblVersion.Caption = Left$(CaricaProgDOS(Var(tmpFile).Valore & "\dumpov2.exe"), 22) & vbNewLine
        ' Mi assicuro che il programma sia chiuso
        objDOS.DosClose ("dumpov2")
        lblVersion.Caption = lblVersion.Caption & PrimaMaiuscola(Mid$(CaricaProgDOS(Var(tmpFile).Valore & "\makeov2.exe"), Len(Var(tmpFile).Valore) + 2, 23))
        ' Mi assicuro che il programma sia chiuso
        objDOS.DosClose ("makeov2")
        lblVersion.Caption = lblVersion.Caption & CaricaProgDOS(Var(tmpFile).Valore & "\gpsbabel.exe -V")
        ' Mi assicuro che il programma sia chiuso
        objDOS.DosClose ("gpsbabel")
    End If

    Select Case Opzioni.Programma
        Case "", "GPSBabel"
            Call CaricaArrayGpsBabel
            Call CaricaListBoxGpsBabel("file")
        Case "TomTom"
            Call CaricaListBoxTomTom
        Case Else
            MsgBox "ATTENZIONE! Errore nella Form.Operazioni." & vbNewLine & "Non è stato impostato il programma di conversione da utilizzare.", vbCritical, App.ProductName
            Call CaricaListBoxTomTom
    End Select
    
    frmRemakeov2.Enabled = False
    PreparaLabel = False
    
    Exit Sub

Errore:
    GestErr Err, "Errore nella funzione frmOperazioni.Form_Load."

End Sub

Private Function CaricaProgDOS(commandLine As String) As String
    
    sRetDOScomm = ""
    objDOS.ExecuteCommand , commandLine
    CaricaProgDOS = sRetDOScomm

End Function

Private Sub objDOS_ReceiveOutputs(CommandOutputs As String)
    sRetDOScomm = sRetDOScomm & CommandOutputs
End Sub

Private Sub CambiaStato()
    Dim strMod As String
    
    strMod = "Modalita': " & Opzioni.Programma & " Tools" & " (" & Opzioni.EstIng & " --> " & Opzioni.EstUsc & ")"
    frmRemakeov2.lblStato.Caption = strMod
    frmRemakeov2.Caption = App.ProductName & " - " & strMod

End Sub

Private Sub ContaFileIngUsc()
    Dim arrTmp
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    ' Scrive gli indici delle ListBox selezionati nella label
    lblIngresso.Caption = "File Ingresso: " & (LBHSi0.ListIndex + 1) & " di " & LBHSi1.ListCount
    lblUscita.Caption = "File Uscita:  " & (LBHSu0.ListIndex + 1) & " di " & LBHSu1.ListCount
    
    ' Imposto lo stato del pulsante cmdSalva
    arrTmp = Split(Var(GpsBabel_In_Out).Valore, ",")
    If (LBHSi0.ListIndex <> CInt(arrTmp(0))) Or (LBHSu0.ListIndex <> CInt(arrTmp(1))) Then
        cmdSalva.Enabled = True
    Else
        cmdSalva.Enabled = False
    End If

    Exit Sub

Errore:
    GestErr Err, "Errore nella funzione ContaFileIngUsc."

End Sub

Private Sub CaricaArrayGpsBabel()
    Dim cnt As Long
    Dim w() As String
    Dim strTmp As String
    Dim strGpsBabel() As String

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    Call CartelleTmp(True)
    
    strTmp = CaricaProgDOS(Var(tmpFile).Valore & "\gpsbabel.exe -^2")
    
    If strTmp <> "" Then
        strGpsBabel = Split(strTmp, Chr(13) + Chr(10))
    
        ReDim arrayGpsBabelCap(UBound(strGpsBabel) - 1)
        
        For cnt = 0 To UBound(strGpsBabel) - 1
            w = Split(strGpsBabel(cnt), Chr(9))
            arrayGpsBabelCap(cnt).Tipo = w(0)
            If Mid$(w(1), 1, 1) = "r" Then arrayGpsBabelCap(cnt).rWaypoints = True Else arrayGpsBabelCap(cnt).rWaypoints = False
            If Mid$(w(1), 2, 1) = "w" Then arrayGpsBabelCap(cnt).wWaypoints = True Else arrayGpsBabelCap(cnt).wWaypoints = False
            If Mid$(w(1), 1, 1) = "r" Then arrayGpsBabelCap(cnt).rTracks = True Else arrayGpsBabelCap(cnt).rTracks = False
            If Mid$(w(1), 2, 1) = "w" Then arrayGpsBabelCap(cnt).wTracks = True Else arrayGpsBabelCap(cnt).wTracks = False
            If Mid$(w(1), 1, 1) = "r" Then arrayGpsBabelCap(cnt).rRoutes = True Else arrayGpsBabelCap(cnt).rRoutes = False
            If Mid$(w(1), 2, 1) = "w" Then arrayGpsBabelCap(cnt).wRoutes = True Else arrayGpsBabelCap(cnt).wRoutes = False
            arrayGpsBabelCap(cnt).FormatName = w(2)
            arrayGpsBabelCap(cnt).Estensione = w(3)
            arrayGpsBabelCap(cnt).FileFormat = w(4)
        Next
    End If
    
    'Aggiungo all'array gli i file style creati manualmente
    If FileExists(Var(tmpFile).Valore & "/csvPoiGPS.style") = True Then
        ReDim Preserve arrayGpsBabelCap(UBound(arrayGpsBabelCap) + 2)
        arrayGpsBabelCap(cnt).Tipo = "file"
        arrayGpsBabelCap(cnt).rWaypoints = True
        arrayGpsBabelCap(cnt).wWaypoints = True
        arrayGpsBabelCap(cnt).rTracks = False
        arrayGpsBabelCap(cnt).wTracks = False
        arrayGpsBabelCap(cnt).rRoutes = False
        arrayGpsBabelCap(cnt).wRoutes = False
        arrayGpsBabelCap(cnt).FormatName = "xcsv,style=csvPoiGPS.style"
        arrayGpsBabelCap(cnt).Estensione = "csv"
        arrayGpsBabelCap(cnt).FileFormat = "Comma separated values PoiGPS"
    End If
        
    Exit Sub

Errore:
    GestErr Err, "Errore nella funzione CaricaArrayGpsBabel."

End Sub

Private Sub cmdSalva_Click()

    If MsgBox("Vuoi salvare le impostazioni sotto selezionate?" & vbNewLine & vbNewLine & "(Scegliendo Si le impostazioni verranno ricordate ad ogni apertura del programma)", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        ' Formato stringa GpsBabel_In_Out: PosListBoxIngresso,PosListBoxUscita
        lVar(GpsBabel_In_Out) = LBHSi0.ListIndex & "," & LBHSu0.ListIndex
        cmdSalva.Enabled = False
    End If
    
End Sub

Private Sub CaricaListBoxGpsBabel(Tipo As String)
    Dim cnt As Long
    Dim arrTmp

    cmdProgramma(0).Visible = False
    cmdProgramma(1).Visible = True
    cmdSalva.Visible = True
    DoEvents
    
    LBHSi0.Clear
    LBHSi1.Clear
    LBHSi2.Clear
     
    LBHSu0.Clear
    LBHSu1.Clear
    LBHSu2.Clear
    
    For cnt = 0 To UBound(arrayGpsBabelCap) - 1
        If arrayGpsBabelCap(cnt).Tipo = Tipo Then
            If arrayGpsBabelCap(cnt).rWaypoints = True Then
                LBHSi0.AddItem (arrayGpsBabelCap(cnt).Estensione)
                LBHSi1.AddItem (arrayGpsBabelCap(cnt).FileFormat)
                LBHSi2.AddItem (arrayGpsBabelCap(cnt).FormatName)
            End If
            If arrayGpsBabelCap(cnt).wWaypoints = True Then
                LBHSu0.AddItem (arrayGpsBabelCap(cnt).Estensione)
                LBHSu1.AddItem (arrayGpsBabelCap(cnt).FileFormat)
                LBHSu2.AddItem (arrayGpsBabelCap(cnt).FormatName)
            End If
        End If
    Next
    
    ' Leggo le impostazioni salvate nel file XML
    arrTmp = Split(Var(GpsBabel_In_Out).Valore, ",")
    LBHSi0.Selected CInt(arrTmp(0))
    LBHSu0.Selected CInt(arrTmp(1))
    
    Call ContaFileIngUsc
    
    Opzioni.Programma = "GPSBabel"

End Sub

Private Sub CaricaListBoxTomTom()

    LBHSi0.Clear
    LBHSi1.Clear
    LBHSi2.Clear
     
    LBHSu0.Clear
    LBHSu1.Clear
    LBHSu2.Clear

    LBHSi1.AddItem ("TomTom Tools .ov2 > .asc utilizzando dumpov2.exe")
    LBHSi0.AddItem ("ov2")
    LBHSi2.AddItem ("TomTom Tools")
    LBHSi1.AddItem ("TomTom Tools .asc per makeov2.exe")
    LBHSi0.AddItem ("asc")
    LBHSi2.AddItem ("TomTom Tools")
    LBHSu1.AddItem ("TomTom Tools .asc > .ov2 utilizzando makeov2.exe")
    LBHSu0.AddItem ("ov2")
    LBHSu2.AddItem ("TomTom Tools")
    LBHSu1.AddItem ("TomTom Tools .ov2 > .asc utilizzando dumpov2.exe")
    LBHSu0.AddItem ("asc")
    LBHSu2.AddItem ("TomTom Tools")

    LBHSi0.Selected (0)
    LBHSu0.Selected (0)

    Call ContaFileIngUsc

    cmdProgramma(0).Visible = True
    cmdProgramma(1).Visible = False
    cmdSalva.Visible = False
    
    Opzioni.Programma = "TomTom"

End Sub

Public Sub ControllaOption()
    ' Imposta i dati per le impostazioni delle operazioni da eseguire
   
    Opzioni.FileIng = LBHSi2.List(LBHSi2.ListIndex)
    Opzioni.EstIng = LBHSi0.List(LBHSi0.ListIndex)
    Opzioni.FileUsc = LBHSu2.List(LBHSu2.ListIndex)
    
    If LBHSu0.List(LBHSu0.ListIndex) <> "" Then
        Opzioni.EstUsc = LBHSu0.List(LBHSu0.ListIndex)
    Else
        Opzioni.EstUsc = LBHSu2.List(LBHSu2.ListIndex)
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objDOS = Nothing
    
    Call CambiaStato
    frmRemakeov2.Enabled = True
    
End Sub

Private Sub lstIngresso0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelezionaRigaIng (0)
End Sub

Private Sub lstIngresso1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelezionaRigaIng (1)
End Sub

Private Sub lstIngresso2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelezionaRigaIng (2)
End Sub

Private Sub lstIngresso0_Click()
    SelezionaRigaIng (0)
End Sub

Private Sub lstIngresso1_Click()
    SelezionaRigaIng (1)
End Sub

Private Sub lstIngresso2_Click()
    SelezionaRigaIng (2)
End Sub

Private Sub SelezionaRigaIng(index As Integer)
   
    Select Case index
        Case Is = 0
            LBHSi1.Selected (LBHSi0.ListIndex)
            LBHSi2.Selected (LBHSi0.ListIndex)
        Case Is = 1
            LBHSi0.Selected (LBHSi1.ListIndex)
            LBHSi2.Selected (LBHSi1.ListIndex)
        Case Is = 2
            LBHSi0.Selected (LBHSi2.ListIndex)
            LBHSi1.Selected (LBHSi2.ListIndex)
    End Select

    If LBHSi0.List(LBHSi0.ListIndex) = "" Then
        Estensione = LBHSi2.List(LBHSi0.ListIndex)
    Else
        Estensione = "." & LBHSi0.List(LBHSi0.ListIndex)
    End If
    frmRemakeov2.File1.Pattern = "*" & Estensione
    frmRemakeov2.List1.Clear

    Call ControllaOption
    Call ContaFileIngUsc
    
End Sub

Private Sub lstUscita0_Click()
    SelezionaRigaUsc (0)
End Sub

Private Sub lstUscita1_Click()
    SelezionaRigaUsc (1)
End Sub

Private Sub lstUscita2_Click()
    SelezionaRigaUsc (2)
End Sub

Private Sub lstUscita0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelezionaRigaUsc (0)
End Sub

Private Sub lstUscita1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelezionaRigaUsc (1)
End Sub

Private Sub lstUscita2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelezionaRigaUsc (2)
End Sub

Private Sub SelezionaRigaUsc(index As Integer)

    Select Case index
        Case Is = 0
            LBHSu1.Selected (LBHSu0.ListIndex)
            LBHSu2.Selected (LBHSu0.ListIndex)
        Case Is = 1
            LBHSu0.Selected (LBHSu1.ListIndex)
            LBHSu2.Selected (LBHSu1.ListIndex)
        Case Is = 2
            LBHSu0.Selected (LBHSu2.ListIndex)
            LBHSu1.Selected (LBHSu2.ListIndex)
    End Select
    
    Call ControllaOption
    Call ContaFileIngUsc
    
End Sub

Private Sub lstIngresso0_Scroll()
    ScrollIng (0)
End Sub
Private Sub lstIngresso1_Scroll()
    ScrollIng (1)
End Sub
Private Sub lstIngresso2_Scroll()
    ScrollIng (2)
End Sub

Private Sub ScrollIng(index As Integer)
    
    Select Case index
        Case Is = 0
            LBHSi1.TopIndex = LBHSi0.TopIndex
            LBHSi2.TopIndex = LBHSi0.TopIndex
        Case Is = 1
            LBHSi0.TopIndex = LBHSi1.TopIndex
            LBHSi2.TopIndex = LBHSi1.TopIndex
        Case Is = 2
            LBHSi0.TopIndex = LBHSi2.TopIndex
            LBHSi1.TopIndex = LBHSi2.TopIndex
    End Select

End Sub

Private Sub lstUscita0_Scroll()
    ScrollUsc (0)
End Sub
Private Sub lstUscita1_Scroll()
    ScrollUsc (1)
End Sub
Private Sub lstUscita2_Scroll()
    ScrollUsc (2)
End Sub
Private Sub ScrollUsc(index As Integer)
   
    Select Case index
        Case Is = 0
            LBHSu1.TopIndex = LBHSu0.TopIndex
            LBHSu2.TopIndex = LBHSu0.TopIndex
        Case Is = 1
            LBHSu0.TopIndex = LBHSu1.TopIndex
            LBHSu2.TopIndex = LBHSu1.TopIndex
        Case Is = 2
            LBHSu0.TopIndex = LBHSu2.TopIndex
            LBHSu1.TopIndex = LBHSu2.TopIndex
    End Select
    
End Sub
