VERSION 5.00
Begin VB.Form frmRemakeov2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ov2tools"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   Icon            =   "frmRemakeov2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox cekBeep 
      Caption         =   "Beep al termine"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   7680
      TabIndex        =   17
      Top             =   7560
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox cekIncludiSubDirectory 
      Caption         =   "Includi sotto cartelle"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   550
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancella1 
      Caption         =   "A&nnulla"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      ToolTipText     =   "Ferma l'operazione di trattamento dei file in corso."
      Top             =   6960
      Width           =   1095
   End
   Begin VB.OptionButton optCopiaIn 
      Caption         =   "Copia in &PoiScaricati"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   21
      Top             =   6960
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton optCopiaIn 
      Caption         =   "Copia in &Desktop"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   20
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CheckBox cekHidden 
      Caption         =   "Imposta il file .bmp come nascosto (consigliato per cellulari Nokia)"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   5520
      TabIndex        =   19
      Top             =   7800
      Width           =   5295
   End
   Begin VB.TextBox txtOutputs 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   6735
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   120
      Width           =   8655
   End
   Begin Remakeov2.Xp_ProgressBar Xp_ProgressBar 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   8640
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdAvviaRemake 
      Caption         =   "&Avvia"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOperazioni 
      Caption         =   "&Operazioni"
      Height          =   495
      Left            =   8400
      TabIndex        =   16
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "&Esci  [Esc]"
      Height          =   495
      Left            =   10320
      TabIndex        =   15
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CheckBox cekAltMetodo 
      Caption         =   "Utilizza metodo alternativo (utilizzare solo in caso di errori)"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   3120
      TabIndex        =   14
      Top             =   7560
      Width           =   4575
   End
   Begin VB.CommandButton cmdLogFile 
      Caption         =   "File &Log"
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Top             =   6960
      Width           =   975
   End
   Begin VB.CheckBox cekVisualizzaLogFile 
      Caption         =   "Visualizza sempre il file .log"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   3120
      TabIndex        =   11
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancella 
      Caption         =   "&Ferma ricerca"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      ToolTipText     =   "Ferma l'operazione di ricerca dei file in corso."
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H80000007&
      Height          =   5910
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   8655
   End
   Begin VB.FileListBox File1 
      ForeColor       =   &H80000007&
      Height          =   1845
      Hidden          =   -1  'True
      Left            =   120
      OLEDragMode     =   1  'Automatic
      System          =   -1  'True
      TabIndex        =   4
      Top             =   6240
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      ForeColor       =   &H80000007&
      Height          =   4815
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton cmdInserisciFile 
      Caption         =   "&Inserisci file"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblStato 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   8160
      Width           =   11655
   End
   Begin VB.Label lblIstruzioni 
      BackStyle       =   0  'Transparent
      Caption         =   "lblIstruzioni"
      ForeColor       =   &H80000007&
      Height          =   975
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmRemakeov2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Obbliga a dichiarare le costanti
Option Explicit

Private objFSO As FileSystemObject
Private objFiles As Files
Private objFile As File
Private objFolders As Folders
Private objFolder As folder

Private Cancel As Boolean

Private classeRES As clsResourceFile
Private WithEvents objDOS As DOSOutputs
Attribute objDOS.VB_VarHelpID = -1
Private LBHS As clsListBox

Private Sub cekBeep_Click()
    lVar(EndBeep) = CStr(cekBeep.value)
End Sub

Private Sub cekHidden_Click()
    lVar(BMPnascoste) = CStr(cekHidden.value)
End Sub

Private Sub cekVisualizzaLogFile_Click()
    lVar(VisualizzaLog) = CStr(cekVisualizzaLogFile.value)
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdLogFile_Click()

    If FileExists(App.path & "\Remakeov2.log") = True Then
        frmLogFile.Show
    Else
        SetTimer hwnd, NV_CLOSEMSGBOX, 2000&, AddressOf TimerProc
        Call MessageBox(hwnd, "File Log non disponibile!", "Chiusura a tempo Message Box", MB_ICONQUESTION Or MB_TASKMODAL)
    End If
    
End Sub

Private Sub cmdOperazioni_Click()
    frmOperazioni.PreparaLabel = True
    frmOperazioni.Show
    cmdAvviaRemake.Enabled = False
    optCopiaIn(0).Enabled = False
    optCopiaIn(1).Enabled = False
End Sub

Private Sub Dir1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    If data.GetFormat(vbCFFiles) Or data.GetFormat(vbCFLink) Then
        For i = 1 To 1
            Drive1.Drive = Left$(data.Files(i), 3)
            Dir1.path = DirectoryFromFile(data.Files(i))
        Next
    End If
    
End Sub

Private Sub Form_Activate()
    If FormIsLoad("frmMain") = True Then frmMain.Visible = False
End Sub

Private Sub Form_Load()
    Dim res As Boolean

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    ' Inizializzo la classe per la ListBox
    Set LBHS = New clsListBox
    LBHS.Attach List1
    ' Attivo il DragDrop
    LBHS.DragDropMode = True
    
    ' Inizializzo la classe per i file di risorse
    Set classeRES = New clsResourceFile

    ' Creo ed initializzo objFSO as a FileSystemObject object
    Set objFSO = New FileSystemObject

    lblIstruzioni.Caption = "1° - Seleziona la cartella contenente i file .ov2 da elaborare, quindi tramite il pulsante ""Inserisci File"", oppure cliccando sul nome del file visualizzato, aggiungili nella lista (oppure trascinali direttamente dentro l'elenco qui sotto)." & vbNewLine & _
                            " 2° - Seleziona l'opzione ""Copia in..."" per scegliere dove copiare i file elaborati." & vbNewLine & _
                            " 3° - Premi il tasto ""Avvia"" ed attendi la creazione dei file (possono passare alcuni minuti)."
                            
    
    Drive1.Drive = Left$(Var(PoiScaricati).Valore, 3)
    Dir1.path = Var(PoiScaricati).Valore
    Dir1.ToolTipText = "Trascina qua un file oppure una directory per caricare i file........"
    
    ' Imposto i filtri di visualizzazione per File1
    File1.Pattern = "*" & Estensione
    File1.Hidden = True
    File1.System = True
    
    txtOutputs.Visible = False
    Xp_ProgressBar.Visible = False
    cekHidden.value = Var(BMPnascoste).Valore
    cekBeep.value = Var(EndBeep).Valore
    cekVisualizzaLogFile.value = Var(VisualizzaLog).Valore
    
    Load frmOperazioni
    DoEvents
    Unload frmOperazioni
    
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione frmRemakeOv2.Form_Load."
    
End Sub

Private Sub cmdAvviaRemake_Click()
    Dim strTmp As String
    
    cmdAvviaRemake.Enabled = False
    optCopiaIn(0).Enabled = False
    optCopiaIn(1).Enabled = False
    Cancel1 = False
    ContoErrori = 0
    Me.MousePointer = vbHourglass
    cmdCancella1.Enabled = True
    Set objDOS = New DOSOutputs

    ' Creo la struttura delle cartelle temporanee con i programmi
    Call CartelleTmp(True)

    Call CopiaFileDaElaborare

    If Opzioni.Programma = "TomTom" Then
        Select Case Opzioni.EstIng
            Case Is = "ov2"
                Call Converti_OV2_ASC
                Call CancellaFile("ov2")
            Case Is = "asc"
            
        End Select
        
        Select Case Opzioni.EstUsc
            Case Is = "ov2"
                Call Converti_ASC_OV2
                Call CancellaFile("asc")
                Call CancellaFile("bmp")
            Case Is = "asc"
            
        End Select
        '
    ElseIf Opzioni.Programma = "GPSBabel" Then
        Call ConvertiGPSBabel
        Call CancellaFile(Opzioni.EstIng)
        
    End If
    
    Xp_ProgressBar.value = 0
    Xp_ProgressBar.Visible = False
    
    If cekBeep = 1 Then
        Beep 250, 200
        Beep 1000, 300
    End If
    cmdCancella1.Enabled = False
    
    ' Apro la finestra per visualizzare il file .log
    If (cekVisualizzaLogFile.value = 1) Or (ContoErrori <> 0) Then
        txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Doppio click per chiudere questa casella di testo."
        frmLogFile.Show
        frmLogFile.SetFocus
    Else
        txtOutputs.Text = ""
        txtOutputs.Visible = False
    End If

    ' Copia dei file ottenuti nella cartella finale
    If optCopiaIn(0).value = True Then strTmp = "PoiScaricati"
    If optCopiaIn(1).value = True Then strTmp = Replace(optCopiaIn(1).Caption, "&", "")
    If MsgBox("Procedo alla copia dei file elaborati nella cartella """ & strTmp & """?", vbInformation + vbYesNo, App.ProductName) = vbYes Then
        If optCopiaIn(0).value = True Then
            CopiaFileElaborati (Var(PoiScaricati).Valore)
        ElseIf optCopiaIn(1).value = True Then
            CopiaFileElaborati (CartellaDesktop("RemakedOv2 File"))
        End If
        lblStato.Caption = "Processo terminato. File elaborati " & LBHS.ListCount
    Else
        lblStato.Caption = "Processo terminato. File elaborati " & LBHS.ListCount & " - File copiati: 0"
    End If

    LBHS.Clear

    ' Cancello i file che non servono dalle cartelle temporanee
    Call CartelleTmp(False)
    
    Me.MousePointer = vbDefault
    
    Set objDOS = Nothing

End Sub

Private Sub CopiaFileDaElaborare()
    Dim NomeFileList As String ' Il nome del vecchio file .ov2
    Dim NomeFileIngresso As String
    Dim NomeFileUscita As String
    Dim DirFileList As String  ' La cartella del vecchio file .ov2
    Dim cnt As Integer
    Dim cntArray As Long
    Dim cntErr As Integer
    Dim res As Boolean

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
'---' Copia dei file da elaborare --------------------------------------------------------------------------------
    ' Dimensiono l'ArrayFileTmp in base al numero di righe in ListBox
    cntArray = LBHS.ListCount - 1
    ReDim ArrayFileTmp(cntArray)
    ' Dimensiono la ProgressBar in base al numero di righe in ListBox
    Xp_ProgressBar.Visible = True
    Xp_ProgressBar.ProgressLook = XP_Default
    Xp_ProgressBar.Max = cntArray + 1
    Xp_ProgressBar.value = 0
    cnt = 0
    
    lblStato.Caption = "Copia dei file da elaborare in corso.........................."
    txtOutputs.Visible = True
    DoEvents
    
    For cnt = 0 To cntArray
        cntErr = 0
        PercorsoFileList = LBHS.List(cnt)
        DirFileList = DirectoryFromFile(PercorsoFileList, True)
        NomeFileList = FileNameFromPath(PercorsoFileList, True)
        
        ' Il solo nome del file di uscita
        NomeFileUscita = ElaboraNomeFile(NomeFileList)
        
        ' Il solo nome del file di ingresso
        Select Case Opzioni.Programma
            Case Is = "GPSBabel"
                NomeFileIngresso = "OldTmp" & NomeFileUscita
            Case Is = "TomTom"
                NomeFileIngresso = NomeFileUscita
            Case Else
                MsgBox ("Errore nella funzione CopiaFileDaElaborare!")
                GoTo Errore
        End Select
        
        ' Iserisco i dati nell'array
        ArrayFileTmp(cnt).PfIngEst = Var(tmpFile).Valore & "\" & NomeFileIngresso
        ArrayFileTmp(cnt).PfIngNoEst = Var(tmpFile).Valore & "\" & Left$(NomeFileIngresso, Len(NomeFileIngresso) - 4)
        ArrayFileTmp(cnt).NfIngEst = NomeFileIngresso
        ArrayFileTmp(cnt).NfIngNoEst = Left$(NomeFileIngresso, Len(NomeFileIngresso) - 4)
        ArrayFileTmp(cnt).PfUscNoEst = Var(tmpFile).Valore & "\" & Left$(NomeFileUscita, Len(NomeFileUscita) - 4)
        ArrayFileTmp(cnt).NfUscNoEst = Left$(NomeFileUscita, Len(NomeFileUscita) - 4)
        ArrayFileTmp(cnt).NfUscNoEstFinale = ElaboraNomeFile(ArrayFileTmp(cnt).NfUscNoEst, True)
        ArrayFileTmp(cnt).LSTnome = NomeFileList
        ArrayFileTmp(cnt).LSTdir = DirFileList
    
        ' Copio il file.ov2 nella cartella temporanea
        Call CopiaFile(PercorsoFileList, Var(tmpFile).Valore & "\" & NomeFileIngresso)
        DoEvents
        
        ' Scrivo nel file.log
        If FileExists(Var(tmpFile).Valore & "\" & NomeFileIngresso) = True Then
            ' Scrivo nel file.log
            WriteLog ("> Copiato file in tmpFile: " & NomeFileIngresso)
            txtOutputs.Text = txtOutputs.Text & "Copiato file in tmpFile: " & NomeFileIngresso & vbNewLine
            DoEvents
        Else
            ' Scrivo nel file.log
            WriteLog ("* Errore copia file in tmpFile: " & NomeFileIngresso)
            ContoErrori = ContoErrori + 1
            txtOutputs.Text = txtOutputs.Text & "Errore copia file in tmpFile: " & NomeFileIngresso & vbNewLine
            DoEvents
        End If
        
        If Cancel1 = True Then GoTo PremutoCancella

        ' Copio il file.bmp nella cartella temporanea
        If FileExists(Left$(PercorsoFileList, Len(PercorsoFileList) - 3) & "bmp") = True Then
            Call CopiaFile(Left$(PercorsoFileList, Len(PercorsoFileList) - 3) & "bmp", Var(tmpBMPfile).Valore & "\" & ArrayFileTmp(cnt).NfUscNoEst & ".bmp")
            DoEvents
        Else
            res = classeRES.SalvaInCartella(105, "BITMAP", Var(tmpBMPfile).Valore, ArrayFileTmp(cnt).NfUscNoEst & ".bmp")
            DoEvents
            WriteLog ("° Attenzione manca file.bmp: " & PercorsoFileList & " (il file è stato creato)")
        End If
        
        If FileExists(Var(tmpBMPfile).Valore & "\" & ArrayFileTmp(cnt).NfUscNoEst & ".bmp") = True Then
            WriteLog ("> Copiato file in tmpBMPfile: " & ArrayFileTmp(cnt).NfUscNoEst & ".bmp")
            txtOutputs.Text = txtOutputs.Text & "Copiato file in tmpBMPfile: " & ArrayFileTmp(cnt).NfUscNoEst & ".bmp" & vbNewLine
        Else
            WriteLog ("* Errore copia file.bmp: " & Left$(PercorsoFileList, Len(PercorsoFileList) - 4) & " - il file è stato creato")
            ContoErrori = ContoErrori + 1
            txtOutputs.Text = txtOutputs.Text & "Errore copia file in tmpBMPfile: " & ArrayFileTmp(cnt).NfUscNoEst & ".bmp" & vbNewLine
        End If
        Xp_ProgressBar.value = cnt
        DoEvents
        If Cancel1 = True Then GoTo PremutoCancella
    Next
    
    lblStato.Caption = ""
    
    Exit Sub

PremutoCancella:
    lblStato.Caption = ""
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione CopiaFileDaElaborare."
    Exit Sub

End Sub

Private Sub ConvertiGPSBabel()
    Dim EsitoAperturaProgramma As Boolean
    Dim fileTmp As String
    Dim cnt As Long
    Dim fNum As Integer
    Dim Apri As String

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    ' Dimensiono la ProgressBar in base al numero di righe dell'Array
    Xp_ProgressBar.ProgressLook = XP_Blue
    Xp_ProgressBar.Max = UBound(ArrayFileTmp) + 1
    Xp_ProgressBar.value = 0
    cnt = 0
        
    lblStato.Caption = "Elaborazione dei file con GPSBabel in corso........attendere"
    DoEvents
    
    ' Avvio il file MakeAscBat per la creazione dei file
    If cekAltMetodo = vbChecked Then
        ' Apro il file in modalità for append
        fNum = FreeFile
        fileTmp = Var(MakeAscBat).Valore
        Open fileTmp For Append As fNum
        ' ......Scrivo i dati nel file
        ' cambio la directory di lavoro al dos
        Print #fNum, "cd " & Var(tmpFile).Valore
        For cnt = 0 To UBound(ArrayFileTmp)
            Apri = ParGPSBabel(cnt)
            ' Scrivo il valore nel file
            Print #fNum, Apri
        Next
        ' Chiudo il file
        Close fNum
        ' Utilizzo il vecchio metodo....
        EsitoAperturaProgramma = eseguiEattendi(Var(MakeAscBat).Valore, , Var(tmpFile).Valore, vbNormalFocus)
        
    Else
        txtOutputs.Visible = True
        For cnt = 0 To UBound(ArrayFileTmp)
            txtOutputs.Text = txtOutputs.Text & "Elaborazione del file: " & ArrayFileTmp(cnt).NfUscNoEst & "." & Opzioni.EstUsc & vbNewLine
            Apri = ParGPSBabel(cnt, Var(tmpFile).Valore & "\gpsbabel.exe")
            objDOS.ExecuteCommand "ExecAndCapture", Apri, "gpsbabel"
            
            Xp_ProgressBar.value = cnt
            DoEvents
            If Cancel1 = True Then GoTo PremutoCancella
        Next
        
    End If

    ' Mi assicuro che il programma sia chiuso
    objDOS.DosClose ("gpsbabel")
    
    ' Cancello il file MakeAscBat
    If FileExists(Var(MakeAscBat).Valore) = True Then Kill Var(MakeAscBat).Valore
    
    lblStato.Caption = "Verifica dei file........"
    DoEvents

    ' Scrivo nel file.log
    For cnt = 0 To UBound(ArrayFileTmp)
        If FileExists(ArrayFileTmp(cnt).PfUscNoEst & "." & Opzioni.EstUsc) = True Then
            WriteLog ("= Creato file in tmpFile: " & ArrayFileTmp(cnt).NfUscNoEst & "." & Opzioni.EstUsc)
        Else
            WriteLog ("* Errore creazione file in tmpFile: " & ArrayFileTmp(cnt).NfUscNoEst & "." & Opzioni.EstUsc)
            ContoErrori = ContoErrori + 1
        End If
        Xp_ProgressBar.value = cnt
        DoEvents
        If Cancel1 = True Then GoTo PremutoCancella
    Next

    Exit Sub

PremutoCancella:
    lblStato.Caption = ""
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione ConvertiGPSBabel."
    Exit Sub

End Sub

Private Function ParGPSBabel(cnt As Long, Optional gpsbexe As String = "gpsbabel.exe")
    ' Prepara il testo da passare a GPSBabel per l'elaborazione del file
    '
    ' Chr(34) è la doppia virgoletta.
    ' Il parametro -p "" serve per indicare a GPSBabel di non cercare il file gpsbabel.ini
    Dim Apri As String
    
    ParGPSBabel = ""
    
    '...\gpsbabel.exe -p "" -w -c MS-ANSI -i xcsv,style=csvPoiGPS.style -f "...\POI.csv" -c MS-ANSI -o tomtom -F "...\POI.ov2"
    '...\gpsbabel.exe -p "" -w -c MS-ANSI -i tomtom -f "...\POI.ov2" -c MS-ANSI -o tomtom -F "...\POI.ov2"
           Apri = gpsbexe & " -p " & Chr(34) & Chr(34) & " -w -c MS-ANSI -i "
    Apri = Apri & Opzioni.FileIng & " -f "
    Apri = Apri & Chr(34) & ArrayFileTmp(cnt).PfIngEst & Chr(34) & " -c MS-ANSI -o "
    Apri = Apri & Opzioni.FileUsc & " -F "
    Apri = Apri & Chr(34) & ArrayFileTmp(cnt).PfUscNoEst & "." & Opzioni.EstUsc & Chr(34)
    
    If Var(DebugMode).Valore = 1 Then WriteLog "ConvertiGPSBabel: " & Apri, "Debug"

    ParGPSBabel = Apri

End Function

Private Sub Converti_OV2_ASC()
    Dim EsitoAperturaProgramma As Boolean
    Dim fileTmp As String
    Dim cnt As Long
    Dim fNum As Integer
    Dim sComandLine As String
    Dim tmpNomeFile As String

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

'---' Conversione .ov2 > .asc --------------------------------------------------------------------------------
    ' Dimensiono la ProgressBar in base al numero di righe dell'Array
    Xp_ProgressBar.ProgressLook = XP_Blue
    Xp_ProgressBar.Max = UBound(ArrayFileTmp) + 1
    Xp_ProgressBar.value = 0
    cnt = 0
    
    ' Apro il file in modalità for append
    fNum = FreeFile
    fileTmp = Var(MakeAscBat).Valore
    Open fileTmp For Append As fNum
    ' ......Scrivo i dati nel file
    ' cambio la directory di lavoro al dos
    Print #fNum, "cd " & Var(tmpFile).Nome
    ' inserisco il ciclo
    Print #fNum, "for %%1 in (*.ov2) do dumpov2.exe %%1"
    ' Chiudo il file
    Close fNum
    
    lblStato.Caption = "Conversione dei file .ov2 in .asc in corso........attendere"
    DoEvents
    
    ' Avvio il file MakeAscBat per la creazione dei file .asc
    If cekAltMetodo = vbChecked Then
        ' Utilizzo il veccio metodo....
        EsitoAperturaProgramma = eseguiEattendi(Var(MakeAscBat).Valore, , Var(tmpFile).Valore, vbNormalFocus)
    Else
        For cnt = 0 To UBound(ArrayFileTmp)
            txtOutputs.Visible = True
            If Opzioni.EstUsc = "asc" Then
                ' Se il file contiene degli spazi aggiungo le virgolette
                If InStr(1, ArrayFileTmp(cnt).NfUscNoEst, " ", vbTextCompare) <> 0 Then
                    tmpNomeFile = Chr(34) & ArrayFileTmp(cnt).NfUscNoEst & "." & Opzioni.EstIng & Chr(34)
                Else
                    tmpNomeFile = ArrayFileTmp(cnt).NfUscNoEst & "." & Opzioni.EstIng
                End If
                sComandLine = Var(tmpFile).Valore & "\dumpov2.exe " & tmpNomeFile
                txtOutputs.Text = txtOutputs.Text & vbNewLine & sComandLine & vbNewLine
                objDOS.ExecuteCommand "ShellExecuteCapture", sComandLine, , Var(tmpFile).Valore
            Else
                ' Se il file contiene degli spazi aggiungo le virgolette
                If InStr(1, ArrayFileTmp(cnt).NfIngNoEst, " ", vbTextCompare) <> 0 Then
                    tmpNomeFile = Chr(34) & ArrayFileTmp(cnt).NfIngNoEst & "." & Opzioni.EstIng & Chr(34)
                Else
                    tmpNomeFile = ArrayFileTmp(cnt).NfIngNoEst & "." & Opzioni.EstIng
                End If
                sComandLine = Var(tmpFile).Valore & "\dumpov2.exe " & tmpNomeFile
                txtOutputs.Text = txtOutputs.Text & vbNewLine & sComandLine & vbNewLine
                objDOS.ExecuteCommand "ShellExecuteCapture", sComandLine, , Var(tmpFile).Valore
            End If
            Xp_ProgressBar.value = cnt
            DoEvents
            If Cancel1 = True Then GoTo PremutoCancella
        Next
        
    End If
    
    ' Mi assicuro che il programma sia chiuso
    objDOS.DosClose ("dumpov2")
    
    ' Cancello il file MakeAscBat
    If FileExists(Var(MakeAscBat).Valore) = True Then Kill Var(MakeAscBat).Valore
    
    lblStato.Caption = "Verifica dei file........"
    DoEvents

    ' Scrivo nel file.log
    For cnt = 0 To UBound(ArrayFileTmp)
        If FileExists(ArrayFileTmp(cnt).PfUscNoEst & ".asc") = True Then
            If Opzioni.EstUsc = "asc" Then
                WriteLog ("= Creato file in tmpFile: " & ArrayFileTmp(cnt).PfUscNoEst & ".asc")
            Else
                WriteLog ("+ Creato file in tmpFile: " & ArrayFileTmp(cnt).PfUscNoEst & ".asc")
            End If
        Else
            WriteLog ("* Errore creazione file in tmpFile: " & ArrayFileTmp(cnt).PfUscNoEst & ".asc")
            ContoErrori = ContoErrori + 1
        End If
        Xp_ProgressBar.value = cnt
        DoEvents
        If Cancel1 = True Then GoTo PremutoCancella
    Next

    Exit Sub

PremutoCancella:
    lblStato.Caption = ""
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione Converti_OV2_ASC."
    Exit Sub

End Sub

Private Sub CancellaFile(EstensioneFile)
    Dim cnt As Long
    Dim Pfile As String
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
'---' Cancellazione dei file --------------------------------------------------------------------------------
    ' Dimensiono la ProgressBar in base al numero di righe dell'Array
    Xp_ProgressBar.ProgressLook = XP_Red
    Xp_ProgressBar.Max = UBound(ArrayFileTmp) + 1
    Xp_ProgressBar.value = 0
    cnt = 0
    
    lblStato.Caption = "Cancellazione dei dei file temporanei........"
    DoEvents
    
    EstensioneFile = LCase(Right$(EstensioneFile, 3))
    
    ' Cancello i file non più necessari
    For cnt = 0 To UBound(ArrayFileTmp)
        Pfile = ArrayFileTmp(cnt).PfIngNoEst & "." & EstensioneFile
        
        If FileExists(Pfile) = True Then
            Kill Pfile
            WriteLog ("- Cancellato file in tmpFile: " & ArrayFileTmp(cnt).NfIngNoEst & "." & EstensioneFile)
        Else
            WriteLog ("* Errore cancellazione file in tmpFile: " & ArrayFileTmp(cnt).NfIngNoEst & "." & EstensioneFile)
            ContoErrori = ContoErrori + 1
            lblStato.Caption = "Errore cancellazione file in tmpFile: " & ArrayFileTmp(cnt).NfIngNoEst & "." & EstensioneFile
            DoEvents
        End If
        Xp_ProgressBar.value = cnt
        DoEvents
        If Cancel1 = True Then GoTo PremutoCancella
    Next
                
    lblStato.Caption = ""
    DoEvents
    
    Exit Sub

PremutoCancella:
    lblStato.Caption = ""
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione CancellaFile. File: " & ArrayFileTmp(cnt).NfIngNoEst & "." & EstensioneFile & "."
    Exit Sub

End Sub

Private Sub Converti_ASC_OV2()
    Dim EsitoAperturaProgramma As Boolean
    Dim fileTmp As String
    Dim cnt As Long
    Dim fNum As Integer
    Dim tmpNomeFile As String
    Dim sCommandLine As String

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

'---' Conversione .asc > .ov2 --------------------------------------------------------------------------------
    Xp_ProgressBar.ProgressLook = XP_Orange
    Xp_ProgressBar.Max = UBound(ArrayFileTmp) + 1
    Xp_ProgressBar.value = 0
    cnt = 0

    lblStato.Caption = "Conversione dei file .asc in .ov2 in corso........attendere"
    DoEvents
    
    ' Avvio il file MakeOv2Bat per la creazione dei file .asc
    If cekAltMetodo = vbChecked Then
        ' Utilizzo il veccio metodo....

        ' Apro il file in modalità for append
        fNum = FreeFile
        fileTmp = Var(MakeOv2Bat).Valore
        Open fileTmp For Append As fNum
        ' ......Scrivo i dati nel file
        ' cambio la directory di lavoro al dos
        Print #fNum, "cd " & Var(tmpFile).Nome
        ' inserisco il ciclo
        Print #fNum, "for %%1 in (*.asc) do makeov2.exe %%1"
        ' Chiudo il file
        Close fNum

        EsitoAperturaProgramma = eseguiEattendi(Var(MakeOv2Bat).Valore, , Var(tmpFile).Valore, vbNormalFocus)
        DoEvents
        
    Else
        txtOutputs.Visible = True
        For cnt = 0 To UBound(ArrayFileTmp)
            ' Se il file contiene degli spazi aggiungo le virgolette
            If InStr(1, ArrayFileTmp(cnt).NfIngNoEst, " ", vbTextCompare) <> 0 Then
                tmpNomeFile = Chr(34) & ArrayFileTmp(cnt).NfIngNoEst & ".asc" & Chr(34)
            Else
                tmpNomeFile = ArrayFileTmp(cnt).NfIngNoEst & ".asc"
            End If
            
            sCommandLine = Var(tmpFile).Valore & "\makeov2.exe " & tmpNomeFile
            txtOutputs.Text = txtOutputs.Text & vbNewLine & sCommandLine & vbNewLine
            objDOS.ExecuteCommand "ShellExecuteCapture", sCommandLine, , Var(tmpFile).Valore
            
            Xp_ProgressBar.value = cnt
            DoEvents
            If Cancel1 = True Then GoTo PremutoCancella
        Next
        
    End If
    
    ' Mi assicuro che il programma sia chiuso
    objDOS.DosClose ("makeov2")

    ' Cancello il file MakeOv2Bat
    If FileExists(Var(MakeOv2Bat).Valore) = True Then Kill Var(MakeOv2Bat).Valore

    lblStato.Caption = "Verifica dei file........"
    DoEvents

    ' Scrivo nel file.log
    For cnt = 0 To UBound(ArrayFileTmp)
        If FileExists(ArrayFileTmp(cnt).PfUscNoEst & "." & Opzioni.EstUsc) = True Then
            WriteLog ("= Creato file in tmpFile: " & ArrayFileTmp(cnt).NfUscNoEst & "." & Opzioni.EstUsc)
            lblStato.Caption = "Creato file in tmpFile: " & ArrayFileTmp(cnt).NfUscNoEst & "." & Opzioni.EstUsc
        Else
            WriteLog ("* Errore creazione file in tmpFile: " & ArrayFileTmp(cnt).NfUscNoEst & "." & Opzioni.EstUsc)
            ContoErrori = ContoErrori + 1
            lblStato.Caption = "Errore creazione file in tmpFile: " & ArrayFileTmp(cnt).NfUscNoEst & "." & Opzioni.EstUsc
        End If
        Xp_ProgressBar.value = cnt
        DoEvents
        If Cancel1 = True Then GoTo PremutoCancella
    Next
    
    lblStato.Caption = ""
    
    Exit Sub

PremutoCancella:
    lblStato.Caption = ""
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione Converti_ASC_OV2."
    Exit Sub

End Sub

Private Sub objDOS_ReceiveOutputs(CommandOutputs As String)
    txtOutputs.Text = txtOutputs.Text & CommandOutputs
End Sub

Private Sub txtOutputs_Change()
    ' Accorcio il numero di caratteri ad un massino di....
    txtOutputs.Text = Right$(txtOutputs.Text, 10000)
    ' Imposto la posizione di scrittura dei nuovi caratteri
    txtOutputs.SelStart = Len(txtOutputs.Text)
End Sub

Private Sub txtOutputs_DblClick()
    txtOutputs.Visible = False
End Sub

Private Sub CopiaFileElaborati(Cartella)
    Dim target_folder As folder
    Dim cnt As Long

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    If Cartella = "" Then
        Screen.MousePointer = vbArrow
        MsgBox "Cartella di destinazione non impostata. I file non sono stati copiati.", vbExclamation, App.ProductName
        Exit Sub
    End If
        
    ' Dimensiono la ProgressBar in base al numero di righe dell'Array
    Xp_ProgressBar.ProgressLook = XP_Blue
    Xp_ProgressBar.Max = UBound(ArrayFileTmp) + 1
    Xp_ProgressBar.value = 0
    cnt = 0

    Screen.MousePointer = vbHourglass
    Xp_ProgressBar.Visible = True
    lblStato.Caption = "Copia dei file creati in corso........................."
    DoEvents
    
    cmdCancella1.Enabled = True

    txtOutputs.Visible = True
    DoEvents
    
    For cnt = 0 To UBound(ArrayFileTmp)
        txtOutputs.Text = txtOutputs.Text & "Copia del file: " & ArrayFileTmp(cnt).NfUscNoEstFinale & "." & Opzioni.EstUsc & vbNewLine
        DoEvents
        ' Copio i file nella directory indicata
        CopiaFile ArrayFileTmp(cnt).PfUscNoEst & "." & Opzioni.EstUsc, Cartella & "\" & ArrayFileTmp(cnt).NfUscNoEstFinale & "." & Opzioni.EstUsc, True
        
        ' Se il file di uscita è ov2 copio anche la bitmap
        Select Case Opzioni.EstUsc
            Case Is = "ov2"
                CopiaFile Var(tmpBMPfile).Valore & "\" & ArrayFileTmp(cnt).NfUscNoEst & ".bmp", Cartella & "\" & ArrayFileTmp(cnt).NfUscNoEstFinale & ".bmp", True
        End Select
        
        Xp_ProgressBar.value = cnt
        DoEvents
        If Cancel1 = True Then GoTo PremutoCancella

        ' Scrivo nel file.log
        If FileExists(Cartella & "\" & ArrayFileTmp(cnt).NfUscNoEstFinale & "." & Opzioni.EstUsc) = True Then
            ' Scrivo nel file.log
            WriteLog ("< Copiato file: " & ArrayFileTmp(cnt).NfUscNoEstFinale & "." & Opzioni.EstUsc)
            DoEvents
        Else
            ' Scrivo nel file.log
            WriteLog ("* Errore copia file: " & ArrayFileTmp(cnt).NfUscNoEstFinale & "." & Opzioni.EstUsc)
            ContoErrori = ContoErrori + 1
            DoEvents
        End If
    Next

    ' Imposto tutti i file .bmp come file nascosti per evitare che vengano visualizzati nella galleria dei cellulari nokia
    Call NascondiFile("*.bmp", Cartella)
    
    txtOutputs.Text = ""
    txtOutputs.Visible = False
    DoEvents
    
    Xp_ProgressBar.Visible = False
    lblStato.Caption = "File copiati in: " & Cartella & vbNewLine & _
                        "Adesso puoi copiare i file della cartella nella cartella della mappa!"
                        
    Shell "explorer.exe " & Cartella, vbNormalFocus

    cmdCancella1.Enabled = False
    
    ' Apro la finestra per visualizzare il file .log
    If ContoErrori <> 0 Then GoTo ApriFormLog
    
    If cekVisualizzaLogFile = 1 Then
ApriFormLog:
        frmLogFile.Show
        frmLogFile.SetFocus
    End If

    Screen.MousePointer = vbArrow
    
    Exit Sub

PremutoCancella:
    lblStato.Caption = ""
    cmdCancella1.Enabled = False
    Screen.MousePointer = vbArrow
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione CopiaFileDaElaborare."
    cmdCancella1.Enabled = False
    Exit Sub

End Sub

Public Sub cmdInserisciFile_Click()
    Dim FPaths() As String
    Dim IncludiSubDirectory As Boolean
    
    If cekIncludiSubDirectory = 1 Then
        IncludiSubDirectory = True
    Else
        IncludiSubDirectory = False
    End If
    
    Cancel = False
    Me.MousePointer = vbHourglass
    lblStato.Caption = ""
    cmdCancella.Enabled = True
    cmdInserisciFile.Enabled = False
    
    ' Cerco i file e li copio nella ListBox
    FindFiles Dir1.path, FPaths(), "*" & Estensione, LBHS, IncludiSubDirectory
    
    If Cancel Then Exit Sub
    
    cmdCancella.Enabled = False
    cmdInserisciFile.Enabled = True
    
    ' Cancello i riferimenti ai file doppi nella ListBox
    LBHS.KillDuplicati True
    LBHS.RefreshHScroll
    If LBHS.ListCount > 0 Then
        cmdAvviaRemake.Enabled = True
        optCopiaIn(0).Enabled = True
        optCopiaIn(1).Enabled = True
        lblStato.Caption = LBHS.ListCount & " file " & Estensione & " in elenco"
    End If
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cmdCancella_Click()
    
    Cancel = True
    lblStato.Caption = ""
    cmdInserisciFile.Enabled = True
    LBHS.Clear
    If Cancel = True Then cmdCancella.Enabled = False
    
End Sub

Private Sub cmdCancella1_Click()
    
    Cancel1 = True
    
    If Cancel1 = True Then
        cmdCancella1.Enabled = False
        DoEvents
        lblStato.Caption = "Annullamento in corso.............attendere"
        DoEvents
        cmdAvviaRemake.Enabled = False
        optCopiaIn(0).Enabled = False
        optCopiaIn(1).Enabled = False
        cmdCancella1.Enabled = False
        Xp_ProgressBar.Visible = False
        LBHS.Clear
        txtOutputs.Text = ""
        txtOutputs.Visible = False
        ' Mi assicuro che il programma sia chiuso
        objDOS.DosClose ("dumpov2")
        objDOS.DosClose ("makeov2")
        MousePointer = 0
        DoEvents
    End If
    
End Sub

Private Sub File1_DblClick()

    LBHS.AddItem File1.path & "\" & File1.filename, True, True, True
    LBHS.RefreshHScroll
    If LBHS.ListCount > 0 Then
        cmdAvviaRemake.Enabled = True
        optCopiaIn(0).Enabled = True
        optCopiaIn(1).Enabled = True
        lblStato.Caption = LBHS.ListCount & " file " & Estensione & " in elenco"
    End If
    
End Sub

Private Sub Dir1_Change()
    File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
    Dim i As Long

    On Error GoTo Errore
    
    Dir1.path = Drive1.Drive
    
Errore:
    If Err.Number = 68 Then
        Dim Drv
        Dim aa As Integer
        Dim bb As Integer
        aa = 97
        bb = 65
        
        For i = 0 To 25
            If Drive1.Drive = Chr(aa) & ":" Then
                Drv = Chr(bb)
            Else
                aa = aa + 1
                bb = bb + 1
            End If
        Next i
        
        If MsgBox(Drv & ":\" & " Non è accessible." & vbCrLf & vbCrLf & "Questo drive non è pronto", vbRetryCancel Or vbCritical, App.ProductName) = vbRetry Then
        
        Resume
            Dir1.path = Drive1.Drive
        Resume Next
        Else
            Drive1.Drive = Dir1.path
        End If
    End If

End Sub

Private Sub FindFiles(strPARENT As String, strPaths() As String, FindFile As String, ListBox As Object, IncludiSubDirectory As Boolean)
    Dim i As Integer
    Dim G As Integer
    Dim FileCount() As Integer
    Dim lngFNAMEScntr As Long

    On Error GoTo ErrorH
    
    'The result in is used to fill a listbox.
    If FindFile = "" Then Exit Sub
    
    Dim lngTopIndex As Long
    Dim lngPathIndex As Long
    Dim strNextPath As String
    Dim UStart As Boolean
    Dim UEnd As Boolean
    Dim Found As Boolean
    
    ReDim fType(2, 0)
    
    Dim filename() As String
    Dim FileNo As Single
    Dim UFlag() As Boolean
    Dim Matches As Integer
    ReDim filename(1)
    
    FileNo = 1
    FileNo = InStr(FileNo, FindFile, ";", vbTextCompare)
    filename(0) = Mid(FindFile, 1, IIf(FileNo > 0, FileNo - 1, Len(FindFile)))
    
    Do While FileNo <> 0
        ReDim Preserve filename(UBound(filename) + 1)
        filename(UBound(filename) - 1) = Mid(FindFile, FileNo + 1, IIf(FileNo > 0, _
        IIf(InStr(FileNo + 1, FindFile, ";", vbTextCompare) > 0, InStr(FileNo + 1, _
        FindFile, ";", vbTextCompare) - (FileNo + 1), Len(FindFile) - FileNo), FileNo - 1))
        FileNo = InStr(FileNo + 1, FindFile, ";", vbTextCompare)
    Loop
    
    ReDim UFlag(3, UBound(filename))
    ReDim FileCount(UBound(filename))
    G = 1
    For i = 0 To UBound(filename) - 1
        UFlag(0, i) = False
        UFlag(1, i) = False
        UFlag(2, i) = False
        Do
            G = InStr(G, filename(i), "*")
            If G = 1 Then
                UFlag(0, i) = True
                filename(i) = Right(filename(i), (Len(filename(i)) - 1))
            End If
            If G = Len(filename(i)) Then
                UFlag(1, i) = True
                filename(i) = Left(filename(i), Len(filename(i)) - 1)
            End If
            If G <> 0 Or G = Len(filename(i)) Then G = G + 1
        Loop Until G = 0
        G = 1
    Next
    
    'Remove *s form the search string
    ' "seed" the loop
    lngTopIndex = 0
    lngPathIndex = 0
    lngFNAMEScntr = 0
    ReDim strPaths(0)
    strPaths(0) = IFBACKSLASH(strPARENT)
    
    Do
        If IncludiSubDirectory = True Then 'SubFolders
            'Creates a folders object containing the subfolders and files
            Set objFolders = objFSO.GetFolder(strPaths(lngPathIndex)).SubFolders
            ' Add subfolders, if any, to folder array
            For Each objFolder In objFolders
                'Increment the folder counter
                lngTopIndex = lngTopIndex + 1
                'Create an additional element in the array
                ReDim Preserve strPaths(lngTopIndex)
                'Store the previous path and the current folder to the path array
                strPaths(lngTopIndex) = strPaths(lngPathIndex) & objFolder.Name & "\"
            Next
        End If
        
        Set objFiles = objFSO.GetFolder(strPaths(lngPathIndex)).Files
        
        lblStato.Caption = strPaths(lngPathIndex)
        DoEvents
        
        ' Add filenames, if any, to array
        For Each objFile In objFiles
            Found = False
            'No Wildcards looks for the exact file name
            For i = 0 To UBound(filename) - 1
                'This next section determines if there is a wild card indicator before
                'and or after the string to locate
                If Not UFlag(0, i) And Not UFlag(1, i) And Not UFlag(2, i) Then
                    If objFSO.FileExists(strPaths(lngPathIndex) & filename(i)) Then
                        ListBox.AddItem strPaths(lngPathIndex) & filename(i), False, False, False
                        UFlag(2, i) = True
                        FileCount(i) = FileCount(i) + 1
                    End If
               End If
                If Len(objFile.Name) <= Len(filename(i)) Then Exit For
                'If the wild card indicator is before and after all finds are located
                If UFlag(0, i) And UFlag(1, i) Then
                    If InStr(1, UCase(objFile.Name), UCase(filename(i))) > 0 Then Found = True
                End If
                'Identifies only those filenames with the Wildcard at the start of the string
                If UFlag(0, i) And Not UFlag(1, i) Then
                    If InStr((Len(objFile.Name) - Len(filename(i))), UCase(objFile.Name), UCase(filename(i))) > 0 Then Found = True
                End If
                'Identifies only those file names with a Wildcard at the end of the string
                If UFlag(1, i) And Not UFlag(0, i) Then
                    If InStr(1, Left(UCase(objFile.Name), Len(filename(i))), UCase(filename(i))) > 0 Then Found = True
                End If
                If Found Then
                    ListBox.AddItem objFile.path, False, False, False
                    FileCount(i) = FileCount(i) + 1
                    Exit For
                End If
                If Cancel Then
                    frmRemakeov2.MousePointer = 0
                    Exit Sub
                End If
            Next
        Next
          
        ' Point to next entry in subfolder array
          lngPathIndex = lngPathIndex + 1
          
        ' If there are no more subfolders, exit
    Loop Until lngPathIndex > lngTopIndex
        
    Exit Sub

ErrorH:
    'Stop
    Resume Next
    
End Sub

Public Function IFBACKSLASH(strX As String) As String
' Funzione per fixing the DOS path of a root directory

     IFBACKSLASH = IIf(Right(strX, 1) = "\", strX, strX & "\")

End Function

Private Sub Form_Unload(Cancel As Integer)
    
    ' Clear all the objects from memory
    Set objFSO = Nothing
    Set objFiles = Nothing
    Set objFile = Nothing
    Set objFolders = Nothing
    Set objFolder = Nothing
    
    If FormIsLoad("frmDownload") = True Then
        frmDownload.Visible = True
        frmDownload.SetFocus
    
    Else
        frmMain.Visible = True
        frmMain.WindowState = vbNormal
        frmMain.ZOrder
        frmMain.SetFocus
    End If
    
End Sub

Private Sub List1_DblClick()
            
    PercorsoFileList = LBHS.List(LBHS.ListIndex)
    frmCheckOV2.Show
    
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Cancella la riga nella ListBox
    
    If KeyCode = 46 Then
        LBHS.RemoveItem LBHS.ListIndex, True
        If LBHS.ListCount > 0 Then
            cmdAvviaRemake.Enabled = True
            optCopiaIn(0).Enabled = True
            optCopiaIn(1).Enabled = True
            lblStato.Caption = LBHS.ListCount & " file " & Estensione & " in elenco"
        Else
            cmdAvviaRemake.Enabled = False
            optCopiaIn(0).Enabled = False
            optCopiaIn(1).Enabled = False
            lblStato.Caption = ""
        End If
    End If


End Sub

Private Sub List1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    If data.GetFormat(vbCFFiles) Then
        For i = 1 To data.Files.count
            If Right$(data.Files(i), 4) = Estensione Then
                Call LBHS.AddItem(data.Files(i), True, True, False)
            Else
                MsgBox ("           Formato del file non supportato!" & vbNewLine & vbNewLine & "Con queste impostazioni puoi caricare i file " & Estensione)
            End If
        Next
    End If
    
    LBHS.RefreshHScroll
    If LBHS.ListCount > 0 Then
        cmdAvviaRemake.Enabled = True
        optCopiaIn(0).Enabled = True
        optCopiaIn(1).Enabled = True
        lblStato.Caption = LBHS.ListCount & " file " & Estensione & " in elenco"
    End If

End Sub

Private Sub List1_OLESetData(data As DataObject, DataFormat As Integer)
    With Me.List1
        Call data.Files.Add(.List(.ListIndex))
    End With
End Sub

Private Sub List1_OLEStartDrag(data As DataObject, AllowedEffects As Long)
    Call data.SetData(, vbCFFiles)
End Sub
