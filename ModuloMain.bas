Attribute VB_Name = "ModuloMain"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

'-------------------------------------------------------------------------------------
' Dichiaro la funzione che serve per fare il Beep
Public Declare Function Beep& Lib "kernel32" (ByVal Freq&, ByVal Duration&)
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
' Per la funzione VerificaStatoInternet
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
' Per la funzione MouseMove che serve per spostare il mouse in un determinato punto di una finestra
'Type POINTAPI
'    x As Long
'    y As Long
'End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
' per la funzione MousePosizioneFinestra
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'GetCursorPos da la posizione del cursore del mouse, rispetto all’angolo superiore sinistro dello schermo.
'Se si desidera conoscere la posizione del mouse, relativamente al form in uso si deve convertire i valori con:
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'-------------------------------------------------------------------------------------

' Per bloccare l'aggiornamento di una finestra---------------------------------------
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
' Prevent the ListView control from updating on screen -
' this is to hide the changes being made to the listitems
' and also to speed up the sort
''LockWindowUpdate .hWnd

' Unlock the list window so that the OCX can update it
''LockWindowUpdate 0&
'-------------------------------------------------------------------------------------

' Per la funzione DisableClose---------------------------------------
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Boolean) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
' --------------------------------------------------------------------

'-------------------------------------------------------------------------------------
' Dichiaro la funzione che serve mantenere una form in primo piano
' Imposta alcuni valori costanti (da WIN32API.TXT).
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'USO per Primo Piano
'Call SetWindowPos(Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, &H3)
'USO per Secondo Piano
'Call SetWindowPos(Form1.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, &H3)
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
' Dichiaro la funzione che serve per lanciare un file o un'applicazione
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5

'-------------------------------------------------------------------------------------
' Dichiaro le funzioni che servono per ottenere una pausa nel programma
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long

'-------------------------------------------------------------------------------------
' Per la funzione WriteLog imposta come dimensione massima del file .log  2 Mb
Private Const MaxLogSize = 2000000
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
' Per verificare se una particolare applicazione e avviata
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT = &H80000000
Private Const MAX_PATH As Integer = 260

Private Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
   th32ModuleID As Long
   cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
   szExeFile As String * MAX_PATH
End Type

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'-------------------------------------------------------------------------------------


'-------------------------------------------------------------------------------------
' L'array che contiene il nome del file con la relativa estensione e percorso
Public Type NomeFile
    ' File ingresso
    PfIngNoEst As String
    PfIngEst As String
    NfIngNoEst As String
    NfIngEst As String
    ' File uscita
    PfUscNoEst As String
    NfUscNoEst As String
    NfUscNoEstFinale As String ' il nome del file trattato con la funziona ElaboraNomeFile........
    ' File ListBox
    LSTnome As String 'solo nome file originale
    LSTdir As String  'solo percorso file originale
End Type
Public ArrayFileTmp() As NomeFile
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
' Per le impostazioni delle operazioni da eseguire
Public Type Opz
    Programma As String
    FileIng As String
    EstIng As String
    FileUsc As String
    EstUsc As String
End Type
Public Opzioni As Opz
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
' Per GpsBabel
Public Type Capabilities
    Tipo As String
    FileFormat As String
    FormatName As String
    rWaypoints As Boolean
    wWaypoints As Boolean
    rTracks As Boolean
    wTracks As Boolean
    rRoutes As Boolean
    wRoutes As Boolean
    Estensione As String
End Type
Public arrayGpsBabelCap() As Capabilities
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
' Per la calibrazione della mappa immagine
Public Type tyInfMappa
    BackColor As Variant
    MapCheck As Long
End Type
Public InfMappa As tyInfMappa
'-------------------------------------------------------------------------------------

Private classeRES As clsResourceFile
Public clsManifestFile As clsManifest

Public Licenza As String            ' Contiene il testo della licenza
Public Versione As String           ' Contiene la versione del programma
Public SecondaIstanza As Boolean    ' Indica se il programma è stato aperto mentre un'altra istanza del programma era già stata aperta
Public arrCampiRmkFile() As String  ' Array con le impostazioni del file .rmk
Public CommandLineFile As String
Public Cancel1 As Boolean
Public ContoErrori As Integer
Public Estensione As String
Public SetupDescrizione As String   ' Le impostazioni per la form SetupDercrizione recuperate dal file .rmk
Public PercorsoFileList As String   ' Il file completo di percorso selezionato nella ListBox
Public XmlFileConfig As String      ' Il percorso completo di nome del file FileConfig.xml
Public DebSep As String             ' Il separatore da inserire nel file Debug.log
Public PrimoAvvio As Boolean        ' Indica se è il primo avvio del programma. Serve per far visualizzare la licenza
Public HomePage As String           ' La home page del programma

Sub Main()
    Dim CommandTmp As String
    Dim nCampi As Integer
    Dim ret
    Dim strTmp As String

    XmlFileConfig = App.path & "\FileConfig.xml"
    Estensione = ".ov2"
    DebSep = "°|°"
    SecondaIstanza = False
    HomePage = "http://remakeov2.poigps.com"
    Versione = App.ProductName & " - Versione: " & App.major & "." & App.minor & " build " & App.Revision & " beta" & vbNewLine & "Freeware by standreani"
    Licenza = "RemakeOv2 - Licenza Freeware" & vbNewLine
    Licenza = Licenza & vbNewLine & "   L'autore non riconosce alcun tipo di garanzia per il prodotto software ed ogni tipo di documentazione ad esso relativa che sono forniti ""così come sono"", senza alcuna garanzia di qualsiasi tipo, sia espressa che implicita, e senza alcuna limitazione."
    Licenza = Licenza & vbNewLine & vbNewLine & "   L'intero rischio derivante dall'uso o dalle prestazioni del prodotto software rimane a carico dell'utente."
    Licenza = Licenza & vbNewLine & vbNewLine & "   Nella misura massima consentita dalla legge in vigore, in nessun caso l'autore o i suoi fornitori saranno responsabili per eventuali danni di qualsiasi genere, accidentali o indiretti derivanti dall'uso o dall'incapacità di utilizzare il prodotto software ovvero dalla fornitura o mancata fornitura del servizio di supporto tecnico."

    ' Leggo i dati del file .xml
    VarXml.LeggiXML
    
    ' 0 = gestione errori attivata  -  1 = gestione errori disattivata
    If Var(GestioneErrori).Valore < 0 Or Var(GestioneErrori).Valore >= 1 Then
        ret = MsgBox("AVVISO: la gestione interna degli errori è stata disattivata!  " & vbNewLine & "Vuoi lasciarla disattivata?   ", vbInformation + vbYesNo, App.ProductName)
        If ret = vbNo Then
            lVar(GestioneErrori) = 0
        End If
    End If
    Set clsManifestFile = New clsManifest

    ' Controllo se esiste la cartella e quindi apro la finestra con la licenza
    If FileExists(Var(tmpFile).Valore) = False Then
        PrimoAvvio = True
        frmAbout.ApriForm , "Hai letto tutta la licenza di utilizzo del programma?" & vbNewLine & "Accetti la licenza?" & vbNewLine & vbNewLine & "- Premendo Si accetti la licenza." & vbNewLine & "- Premendo No non accetti la licenza ed il programma verrà chiuso." & vbNewLine & "- Premendo Annulla puoi tornare a leggere la licenza."
    End If
    
    ' Scrivo il numero di versione del programma nel file per il sito web
    If FileExists(App.path & "\RemakeOv2.ver") = True Then
        strTmp = LeggiFile(App.path & "\RemakeOv2.ver")
        If Right$(strTmp, 2) = vbCrLf Then strTmp = Left$(strTmp, Len(strTmp) - 2)
        ' Recupero il testo contenuto nel file dopo la prima riga
        strTmp = Trim$("News:" & Right$(strTmp, Len(strTmp) - InStr(1, strTmp, "News:", vbTextCompare) - 4))
        CreaFile App.path & "\RemakeOv2.ver", GetProgVers & vbNewLine & strTmp
    End If
    
    ' Controllo che non ci sia già un'altra istanza del programma avviata
    Call LimitaAvvio(3, frmRemakeov2, App.ProductName & " è già aperto!" & vbNewLine & "Non puoi aprire due istanze del programma.")

    frmSplash.Show
    
    ReDim arrCampiRmkFile(0)
    ReDim arrayGpsBabelCap(0)
    
    ' Creo la struttura delle cartelle temporanee
    Call CartelleTmp(True)

    ' Verifico se ci sono da fare adattamenti per le versioni più vecchie del programma
    Call AdattamentiPrecedentiVersioni
    
    ' Sostituisco i caratteri " con un carattere nullo
    CommandTmp = Replace(Command$, Chr(34), "", , , vbTextCompare)
    If CommandTmp <> "" Then ' Se il programma è stato avviato tramite riga di comando....
        CommandLineFile = GetLongFileName(CommandTmp$)
        
        If Dir(CommandLineFile) = "" Then
            CommandLineFile = CommandTmp
        End If
        
        If Dir(CommandLineFile) = "" Then
            CommandLineFile = ""
            frmMain.Show
        Else
            Estensione = Right$(CommandLineFile, 4)
            ' MsgBox CommandLineFile
            Load frmMain
            
            If Estensione = ".ov2" Or Estensione = ".asc" Or Estensione = ".kml" Or Estensione = ".gpx" Then
                frmWeb.Show
            Else
                nCampi = CaricaCampiFileRmk(CommandLineFile)
                If arrCampiRmkFile(0) = "FileDatiPOI" Then
                    frmWeb.Show
                ElseIf arrCampiRmkFile(0) = "FileImpostazioniDownloadWeb" Then
                    frmDownload.Show
                End If
            End If
        End If
    Else
        CommandLineFile = ""
        frmMain.Show
    End If
    
    Call DisableClose(frmMain)
       
    Unload frmSplash

End Sub

Public Sub ChiudiProgramma()
    
    lVar(VersioneProgramma) = App.major & "." & App.minor & "." & App.Revision
    
    ' Chiudo tutti i form
    ChiudiTuttiForm
    
    Set clsManifestFile = Nothing

End Sub

Public Sub AdattamentiPrecedentiVersioni()
    Dim sPrecVers As String
    Dim arrTmp
    Dim Msg As String
    Dim bMessaggio As Boolean
    Const nrg As String = vbNewLine
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    sPrecVers = Var(VersioneProgramma).Valore
    If sPrecVers = "" Or FileExists(XmlFileConfig) = False Then Exit Sub
    
    Msg = "ATTENZIONE!" & nrg & nrg & "A causa della modifica di alcune impostazioni del programma e' necessario:" & nrg
    
    If sPrecVers <= "1.4.898" Then
        ' Formato stringa GpsBabel_In_Out: PosListBoxIngresso,PosListBoxUscita
        lVar(GpsBabel_In_Out) = "71,70"
        Msg = Msg & vbNewLine & "- adattamento a GPSBabel 1.3.4 - premere il tasto Ripara e converti file o poi il tasto Operazioni e verifcare che i valori nelle due liste siano impostati su ov2 oppure sul formato desiderato."
        bMessaggio = True
    End If
    
    'If sPrecVers < GetProgVers Then
    '    If FileExists(App.path & "/andrMap.ocx") = False Then Kill App.path & "/andrMap.ocx"
    'End If
    
    DoEvents
    
    If bMessaggio = True Then MsgBox Msg, vbInformation, App.ProductName
    
    Exit Sub

Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione AdattamentiPrecedentiVersioni."

End Sub

Private Sub LimitaAvvio(Tipo As Integer, Mio As Object, Messaggio As String)
    ' Controlla se c'è già una istanza attiva
    ' se tipo = 0 default avvia una successiva istanza
    ' se tipo = 1 termina la nuova istanza con un messaggio
    ' se tipo = 2 passa il controllo alla prima
    ' se tipo = 3 apre solo una...........
    ' N.B. bisogna dichiarare LimitaAvvio nel form_load principale
    '
    ' Call LimitaAvvio(2, Me, "")
    Dim ret
    Dim sTitle As String

    On Local Error GoTo Errore
    
    If Tipo = 0 Then Exit Sub

    If Tipo = 1 Then
        If App.PrevInstance = True Then
            If Messaggio <> "" Then ret = MsgBox(Messaggio, vbExclamation, App.ProductName)
                End
        End If
        Exit Sub
    End If
            
    If Tipo = 2 Then
        If App.PrevInstance = True Then
            sTitle = Mio.Caption
            Mio.Caption = Hex$(Mio.hwnd)
            AppActivate sTitle
        End
        End If
        Exit Sub
    End If

    If Tipo = 3 Then
        If App.PrevInstance = True Then
            SecondaIstanza = True
            If Messaggio <> "" And Command = "" Then
                ret = MsgBox(Messaggio, vbExclamation, App.ProductName)
                End
            End If
        End If
        Exit Sub
    End If

Errore:
    End
    
End Sub

Public Sub CartelleTmp(Crea As Boolean)
    Dim res As Boolean
    Dim strTmp As String

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    ' Inizializzo la classe per i file di risorse
    Set classeRES = New clsResourceFile

    If Crea = True Then
        ' Cancello il file.log
        If FileExists(App.path & "\Remakeov2.Log") = True Then Kill (App.path & "\Remakeov2.Log")
        
        ' Creo la cartella tmpFile cancellando quella esistente
        Call CreateFolder(Var(tmpFile).Valore, True)
        ' Creo la cartella tmpBMPfile cancellando quella esistente
        Call CreateFolder(Var(tmpBMPfile).Valore, True)
        ' Creo la cartella tmpDownFile non cancellando quella esistente
        Call CreateFolder(Var(tmpDownFile).Valore, False)
        ' Creo la cartella PoiScaricati non cancellando quella esistente
        Call CreateFolder(Var(PoiScaricati).Valore, False)
        ' Creo la cartella CartellaMappe non cancellando quella esistente
        Call CreateFolder(Var(CartellaMappe).Valore, False)
        ' Creo la cartella CartellaScript non cancellando quella esistente
        Call CreateFolder(Var(CartellaScript).Valore, False)
        DoEvents
        
        If FileExists(Var(tmpFile).Valore & "/dumpov2.exe") = False Then res = classeRES.SalvaInCartella(1, "24", Var(tmpFile).Valore, "24.txt")
        
        ' Copio i file necessari nella cartella tmpFile dal file delle risorse
        If FileExists(Var(tmpFile).Valore & "/dumpov2.exe") = False Then res = classeRES.SalvaInCartella(101, "DUMPOV2", Var(tmpFile).Valore, "dumpov2.exe")
        If FileExists(Var(tmpFile).Valore & "/makeov2.exe") = False Then res = classeRES.SalvaInCartella(102, "MAKEOV2", Var(tmpFile).Valore, "makeov2.exe")
        If FileExists(Var(tmpFile).Valore & "/gpsbabel.exe") = False Then res = classeRES.SalvaInCartella(103, "GPSBABEL", Var(tmpFile).Valore, "gpsbabel.exe")
        If FileExists(Var(tmpFile).Valore & "/libexpat.dll") = False Then res = classeRES.SalvaInCartella(104, "GPSBABEL", Var(tmpFile).Valore, "libexpat.dll")
        If FileExists(App.path & "/vbuzip10.dll") = False Then res = classeRES.SalvaInCartella(106, "ZIP", App.path, "vbuzip10.dll")
        If FileExists(App.path & "/GpTabXP.ocx") = False Then res = classeRES.SalvaInCartella(107, "COMPONENTI", App.path, "GpTabXP.ocx")
        If FileExists(App.path & "/andrMap.ocx") = False Then res = classeRES.SalvaInCartella(108, "COMPONENTI", App.path, "andrMap.ocx")
        If FileExists(Var(CartellaMappe).Valore & "/Mondo.gif") = False Then res = classeRES.SalvaInCartella(109, "BITMAP", Var(CartellaMappe).Valore, "Mondo.gif")
        If FileExists(Var(CartellaMappe).Valore & "/Mondo.gif.inf") = False Then
            strTmp = "Cal1Lon=-180" & vbCrLf & "Cal1Lat=90" & vbCrLf & "Cal1X=0" & vbCrLf & "Cal1Y=0" & vbCrLf & "Cal2Lon=180" & vbCrLf & "Cal2Lat=-90" & vbCrLf & "Cal2X=720" & vbCrLf & "Cal2Y=360" & vbCrLf & "BackColor=clBtnFace" & vbCrLf & "MapCheck=19336" & vbCrLf
            CreaFile Var(CartellaMappe).Valore & "/Mondo.gif.inf", strTmp
        End If
        If FileExists(Var(tmpFile).Valore & "/csvPoiGPS.style") = False Then res = classeRES.SalvaInCartella(111, "GPSBABEL", Var(tmpFile).Valore, "csvPoiGPS.style")
        DoEvents
        
         ' Se il file Regioni.csv non esiste lo creo
        If FileExists(Var(RegioniCsv).Valore) = False Then
            strTmp = "Abruzzo"
            strTmp = strTmp & vbNewLine & "Basilicata"
            strTmp = strTmp & vbNewLine & "Calabria"
            strTmp = strTmp & vbNewLine & "Campania"
            strTmp = strTmp & vbNewLine & "Emilia-Romagna"
            strTmp = strTmp & vbNewLine & "Friuli-Venezia Giulia"
            strTmp = strTmp & vbNewLine & "Lazio"
            strTmp = strTmp & vbNewLine & "Liguria"
            strTmp = strTmp & vbNewLine & "Lombardia"
            strTmp = strTmp & vbNewLine & "Marche"
            strTmp = strTmp & vbNewLine & "Molise"
            strTmp = strTmp & vbNewLine & "Piemonte"
            strTmp = strTmp & vbNewLine & "Puglia"
            strTmp = strTmp & vbNewLine & "Sardegna"
            strTmp = strTmp & vbNewLine & "Sicilia"
            strTmp = strTmp & vbNewLine & "Toscana"
            strTmp = strTmp & vbNewLine & "Trentino-Alto Adige"
            strTmp = strTmp & vbNewLine & "Umbria"
            strTmp = strTmp & vbNewLine & "Valle d'Aosta"
            strTmp = strTmp & vbNewLine & "Veneto"
            CreaFile Var(RegioniCsv).Valore, strTmp
        End If
       
        ' Se il file Province.csv non esiste lo creo
        If FileExists(Var(ProvinceCsv).Valore) = False Then
            strTmp = "AG"
            strTmp = strTmp & vbNewLine & "AL"
            strTmp = strTmp & vbNewLine & "AN"
            strTmp = strTmp & vbNewLine & "AO"
            strTmp = strTmp & vbNewLine & "AP"
            strTmp = strTmp & vbNewLine & "AQ"
            strTmp = strTmp & vbNewLine & "AR"
            strTmp = strTmp & vbNewLine & "AT"
            strTmp = strTmp & vbNewLine & "AV"
            strTmp = strTmp & vbNewLine & "BA"
            strTmp = strTmp & vbNewLine & "BG"
            strTmp = strTmp & vbNewLine & "BI"
            strTmp = strTmp & vbNewLine & "BL"
            strTmp = strTmp & vbNewLine & "BN"
            strTmp = strTmp & vbNewLine & "BO"
            strTmp = strTmp & vbNewLine & "BR"
            strTmp = strTmp & vbNewLine & "BS"
            strTmp = strTmp & vbNewLine & "BZ"
            strTmp = strTmp & vbNewLine & "CA"
            strTmp = strTmp & vbNewLine & "CB"
            strTmp = strTmp & vbNewLine & "CE"
            strTmp = strTmp & vbNewLine & "CH"
            strTmp = strTmp & vbNewLine & "CL"
            strTmp = strTmp & vbNewLine & "CN"
            strTmp = strTmp & vbNewLine & "CO"
            strTmp = strTmp & vbNewLine & "CR"
            strTmp = strTmp & vbNewLine & "CS"
            strTmp = strTmp & vbNewLine & "CT"
            strTmp = strTmp & vbNewLine & "CZ"
            strTmp = strTmp & vbNewLine & "EN"
            strTmp = strTmp & vbNewLine & "FC"
            strTmp = strTmp & vbNewLine & "FE"
            strTmp = strTmp & vbNewLine & "FG"
            strTmp = strTmp & vbNewLine & "FI"
            strTmp = strTmp & vbNewLine & "FR"
            strTmp = strTmp & vbNewLine & "GE"
            strTmp = strTmp & vbNewLine & "GO"
            strTmp = strTmp & vbNewLine & "GR"
            strTmp = strTmp & vbNewLine & "IM"
            strTmp = strTmp & vbNewLine & "IS"
            strTmp = strTmp & vbNewLine & "KR"
            strTmp = strTmp & vbNewLine & "LC"
            strTmp = strTmp & vbNewLine & "LE"
            strTmp = strTmp & vbNewLine & "LI"
            strTmp = strTmp & vbNewLine & "LO"
            strTmp = strTmp & vbNewLine & "LT"
            strTmp = strTmp & vbNewLine & "LU"
            strTmp = strTmp & vbNewLine & "MC"
            strTmp = strTmp & vbNewLine & "ME"
            strTmp = strTmp & vbNewLine & "MI"
            strTmp = strTmp & vbNewLine & "MN"
            strTmp = strTmp & vbNewLine & "MO"
            strTmp = strTmp & vbNewLine & "MS"
            strTmp = strTmp & vbNewLine & "MT"
            strTmp = strTmp & vbNewLine & "NA"
            strTmp = strTmp & vbNewLine & "NO"
            strTmp = strTmp & vbNewLine & "NU"
            strTmp = strTmp & vbNewLine & "OR"
            strTmp = strTmp & vbNewLine & "PA"
            strTmp = strTmp & vbNewLine & "PC"
            strTmp = strTmp & vbNewLine & "PD"
            strTmp = strTmp & vbNewLine & "PE"
            strTmp = strTmp & vbNewLine & "PG"
            strTmp = strTmp & vbNewLine & "PI"
            strTmp = strTmp & vbNewLine & "PN"
            strTmp = strTmp & vbNewLine & "PO"
            strTmp = strTmp & vbNewLine & "PR"
            strTmp = strTmp & vbNewLine & "PT"
            strTmp = strTmp & vbNewLine & "PU"
            strTmp = strTmp & vbNewLine & "PV"
            strTmp = strTmp & vbNewLine & "PZ"
            strTmp = strTmp & vbNewLine & "RA"
            strTmp = strTmp & vbNewLine & "RC"
            strTmp = strTmp & vbNewLine & "RE"
            strTmp = strTmp & vbNewLine & "RG"
            strTmp = strTmp & vbNewLine & "RI"
            strTmp = strTmp & vbNewLine & "RM"
            strTmp = strTmp & vbNewLine & "RN"
            strTmp = strTmp & vbNewLine & "RO"
            strTmp = strTmp & vbNewLine & "SA"
            strTmp = strTmp & vbNewLine & "SI"
            strTmp = strTmp & vbNewLine & "SO"
            strTmp = strTmp & vbNewLine & "SP"
            strTmp = strTmp & vbNewLine & "SR"
            strTmp = strTmp & vbNewLine & "SS"
            strTmp = strTmp & vbNewLine & "SV"
            strTmp = strTmp & vbNewLine & "TA"
            strTmp = strTmp & vbNewLine & "TE"
            strTmp = strTmp & vbNewLine & "TN"
            strTmp = strTmp & vbNewLine & "TO"
            strTmp = strTmp & vbNewLine & "TP"
            strTmp = strTmp & vbNewLine & "TR"
            strTmp = strTmp & vbNewLine & "TS"
            strTmp = strTmp & vbNewLine & "TV"
            strTmp = strTmp & vbNewLine & "UD"
            strTmp = strTmp & vbNewLine & "VA"
            strTmp = strTmp & vbNewLine & "VB"
            strTmp = strTmp & vbNewLine & "VC"
            strTmp = strTmp & vbNewLine & "VE"
            strTmp = strTmp & vbNewLine & "VI"
            strTmp = strTmp & vbNewLine & "VR"
            strTmp = strTmp & vbNewLine & "VT"
            strTmp = strTmp & vbNewLine & "VV"
            CreaFile Var(ProvinceCsv).Valore, strTmp
        End If
        
        If FileExists(Var(TipoPdiCsv).Valore) = False Then
            strTmp = "PDI Italia" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=2"
            strTmp = strTmp & vbNewLine & "PDI Austria" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=23"
            strTmp = strTmp & vbNewLine & "PDI Benelux" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=26"
            strTmp = strTmp & vbNewLine & "PDI Danimarca" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=27"
            strTmp = strTmp & vbNewLine & "PDI Francia" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=4"
            strTmp = strTmp & vbNewLine & "PDI Germania" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=5"
            strTmp = strTmp & vbNewLine & "PDI Inghilterra" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=3"
            strTmp = strTmp & vbNewLine & "PDI per Disabili" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=30"
            strTmp = strTmp & vbNewLine & "PDI Spagna" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=18"
            strTmp = strTmp & vbNewLine & "PDI Svizzera" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=28"
            strTmp = strTmp & vbNewLine & "PDI Ungheria" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=29"
            strTmp = strTmp & vbNewLine & "PDI Utenti" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=17"
            strTmp = strTmp & vbNewLine & "PDI Destinator" & vbTab & "http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=34"
            CreaFile Var(TipoPdiCsv).Valore, strTmp
        End If

        If FileExists(Var(Sostituisci).Valore) = False Then
            strTmp = " |_"
            strTmp = strTmp & ""
            CreaFile Var(Sostituisci).Valore, strTmp
        End If

    Else
        ' Cancello i file
        If FileExists(Var(tmpFile).Valore & "\dumpov2.exe") = True Then Kill (Var(tmpFile).Valore & "\dumpov2.exe")
        If FileExists(Var(tmpFile).Valore & "\makeov2.exe") = True Then Kill (Var(tmpFile).Valore & "\makeov2.exe")
        If FileExists(Var(tmpFile).Valore & "\gpsbabel.exe") = True Then Kill (Var(tmpFile).Valore & "\gpsbabel.exe")
        If FileExists(Var(tmpFile).Valore & "\libexpat.dll") = True Then Kill (Var(tmpFile).Valore & "\libexpat.dll")
    End If
    
    ' Se non serve cancelo il file
    If Var(DebugMode).Valore = 0 And FileExists(App.path & "\Debug.log") = True Then Kill (App.path & "\Debug.log")
    
    Exit Sub

Errore:
    GestErr Err, "Errore nella funzione CartelleTmp. Variabile = " & Crea
    Exit Sub

End Sub

Public Sub CreaRemakeBat()
' Questa routine crea il file Remake.bat da utilizzare per la correzione manuale

'del *.asc
'for %%1 in (*.ov2) do dumpov2.exe %%1
'del *.ov2
'for %%1 in (*.asc) do makeov2.exe %%1
'del *.bmp
'del *.asc

End Sub

Public Sub WaitMicroSecondi(Pausa As Long)
    ' In questo modo si utilizza una funzione api per attendere un certo tempo
    ' Perché utilizzare questa procedura?
    ' Prima di tutto è precisa al millesimo di secondo, secondo rispetta l'idea di
    ' multitasking, cioè non occupa risorse durante l'attesa.
    ' Avete mai provato a vere cosa succede alle risorse (uso CPU) durante altre procedure di attesa? fatelo e rimarrete sconcertati.
    
    ' Pausa = microsecondi di pausa
    ' 1000000 dovrebbe corrispondere ad un secondo
    
    Dim curStart As LARGE_INTEGER
    Dim curEnd As LARGE_INTEGER
    Dim curFreq As LARGE_INTEGER
    Dim WaitTop As Long
    
    On Error Resume Next
    
    QueryPerformanceFrequency curFreq
    QueryPerformanceCounter curStart
    
    WaitTop = curStart.lowpart + Int(Pausa * 3.38) '* Int(CDbl(curFreq.lowpart) / 894950)
    
    Do
        QueryPerformanceCounter curEnd
    Loop Until curEnd.lowpart >= WaitTop
      
End Sub

Public Sub WriteLog(ByVal Stringa1 As String, Optional ByVal NomeFile As String = "")

   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Dim sLogFile As String, sLogPath As String, iLogSize As Long
   Dim fso, f
    
   On Error GoTo errhandler
    
   If NomeFile = "" Then NomeFile = App.EXEName
   
   'Set the path and filename of the log
   sLogPath = App.path & "\" & NomeFile
   sLogFile = sLogPath & ".log"
   
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile(sLogFile, ForAppending, True)
   
   'Get the size of the log to check if it's getting unwieldly
   iLogSize = GetLogSize(sLogFile)
   
   If iLogSize > MaxLogSize Then
        'If too big, back it up to to retain some sort of history
        fso.CopyFile sLogFile, (sLogPath & ".old"), True
        Set f = Nothing
        fso.DeleteFile sLogFile
        'And start with a clean log-file
        Set f = fso.OpenTextFile(sLogFile, ForAppending, True)
   End If
    
   'Append the log-entry to the file together with time and date
   f.WriteLine Now() & vbTab & Stringa1
   
errhandler:
    Exit Sub
    
End Sub

Public Function GetLogSize(FileSpec As String) As Long
    ' Returns the size of a file in bytes.
    ' If the file does not exist, it returns -1.

    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If (fso.FileExists(FileSpec)) Then
         Set f = fso.GetFile(FileSpec)
         GetLogSize = f.Size
    Else
         GetLogSize = -1
    End If
    
End Function

Public Function PrimaMaiuscola(ByVal vData As String) As String
   Dim i As Integer         '  Contatore di lettere

   vData = Trim$(vData)     '  Stringa tutto maiuscolo

   If Len(vData) < 1 Then   ' Calcola la lunghezza della stringa
      PrimaMaiuscola = vData
      Exit Function
   End If

   vData = UCase$(Left$(vData, 1)) & LCase$(Right$(vData, Len(vData) - 1))

   For i = 2 To Len(vData)
      If (Mid$(vData, i, 1) = " ") Or (Mid$(vData, i, 1) = "-") And (i + 1 <= Len(vData)) Then
         Mid$(vData, i + 1, 1) = UCase$(Mid$(vData, i + 1, 1))
      End If
   Next i

   PrimaMaiuscola = Trim$(vData)
   
End Function

Public Function CancIsRunning(AppName As String) As Boolean
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim r As Long
    
    'Takes a snapshot of the processes and the heaps, modules, and threads used by the processes
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)

    'set the length of our ProcessEntry-type
    uProcess.dwSize = Len(uProcess)

    'Retrieve information about the first process encountered in our system snapshot
    r = Process32First(hSnapShot, uProcess)
    
    Do While r
        If AppName = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0)) Then
            CancIsRunning = True
            Exit Do
        Else
            'Retrieve information about the next process recorded in our system snapshot
            r = Process32Next(hSnapShot, uProcess)
            CancIsRunning = False
        End If
    Loop
   
    'close our snapshot handle
    CloseHandle hSnapShot
    
End Function

Public Sub ChiudiTuttiForm()
' Purpose: To Unload all forms once the program finishes
   Dim f As Integer
   
   f = Forms.count
   Do While f > 0
       Unload Forms(f - 1)
       If f = Forms.count Then Exit Do
       f = f - 1
   Loop
   
End Sub

Public Function GetLongFileName(sShortName As String) As String
 'In questo caso non serve una funzione di Windows, come per il procedimento inverso, ma basta una semplice routine Visual Basic:
 'Per recuperare il nome lungo del file, si utilizza la funzione Dir, la quale restituisce il nome del file che corrisponde a un attributo o a un tipo di file specificato; di conseguenza, se il valore restituito dalla funzione è una stringa vuota, significa che il file non esiste.

    Dim sLongName As String, sTemp As String
     Dim iSlashPos As Integer
    
    'Aggiunge un backlash ("\") al nome corto del file.
    sShortName = sShortName & "\"

   'Inizia la ricerca dal quarto carattere per saltare la lettera di unità (es.: "C:\").
    iSlashPos = InStr(4, sShortName, "\")

    While iSlashPos
        sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal Or vbHidden Or vbSystem Or vbDirectory)
        If sTemp = "" Then
            'Errore 52 - Nome o numero di file non valido
             GetLongFileName = ""
            Exit Function
        End If

    sLongName = sLongName & "\" & sTemp
    iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
    Wend

    'Inserisce la lettera di unità.
    GetLongFileName = UCase$(Left$(sShortName, 2)) & sLongName

End Function

Public Sub SetOnTop(Form As Form, Optional Attiva = True)
    ' Form in primo piano
    Dim handle As Long
    handle = Form.hwnd
    
    SetWindowPos handle, HWND_TOPMOST, Form.Left / 15, Form.Top / 15, Form.Width / 15, Form.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub

Public Function ContieneNumero(Stringa As String) As Boolean
    ' Restituisce True se nella stringa c'è almento un numero
    Dim cnt As Integer
    
    For cnt = 1 To Len(Stringa)
        If IsNumeric(Left$(Stringa, cnt)) Then
            ContieneNumero = True
            Exit Function
        End If
    Next
    
    ContieneNumero = False
    
End Function

Public Function TrimDUP(TextIN, Optional TrimChar = " ") As String
    'Remove Duplicate Spaces or the Duplicate Character

    On Error GoTo LocalError

    TrimChar = CStr(TrimChar)
    TrimDUP = CStr(TextIN)
    TrimDUP = Replace(TrimDUP, TrimChar, vbNullChar)
    While InStr(TrimDUP, String(2, vbNullChar)) > 0
        TrimDUP = Replace(TrimDUP, String(2, vbNullChar), vbNullChar)
    Wend

    ' Delete Leading and Trailing
    If Left(TrimDUP, 1) = vbNullChar Then TrimDUP = Right(TrimDUP, Len(TrimDUP) - 1)
    If Right(TrimDUP, 1) = vbNullChar Then TrimDUP = Left(TrimDUP, Len(TrimDUP) - 1)

LocalError:
    TrimDUP = Replace(TrimDUP, vbNullChar, TrimChar, , , vbTextCompare)
    
End Function

Public Function FormIsLoad(ByVal sForm As String) As Boolean
    ' Verifica se un Form è o no stato caricato
    ' Con questa funzione è possibile, a run time,
    ' capire se un certo Form è o no stato caricato
    Dim i As Long
    Dim f As Form

    FormIsLoad = False
    For i = 0 To Forms.count - 1
        If LCase$(Forms.Item(i).Name) = LCase$(sForm) Then
           FormIsLoad = True
           Exit Function
        End If
    Next

End Function

Public Sub DisableClose(ByRef frm As Form, Optional Disable As Boolean = True)
    'Setting Disable to False disables the 'X', otherwise, it's reset
    Dim hMenu As Long
    Dim nCount As Long
    
    If Disable Then
        hMenu = GetSystemMenu(frm.hwnd, False)
        nCount = GetMenuItemCount(hMenu)
        Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or MF_BYPOSITION)
        Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or MF_BYPOSITION)
        DrawMenuBar frm.hwnd
    Else
        GetSystemMenu frm.hwnd, True
        DrawMenuBar frm.hwnd
    End If

End Sub

Public Function MostraPosizione(ListView1 As ListView, SitoWeb As String, Riga As Long) As String
    Dim Latitudine As String
    Dim Longitudine As String
    Dim Cap As String
    Dim Indirizzo As String
    Dim Civico As String
    Dim Città As String
    
    If ControllaCella(ListView1, Riga, 8) = False Then GoTo Esci
    If ControllaCella(ListView1, Riga, 9) = False Then GoTo Esci
    
    Latitudine = ListView1.ListItems.Item(Riga).ListSubItems.Item(8).Text
    Latitudine = Replace(Latitudine, ",", ".")
    Longitudine = ListView1.ListItems.Item(Riga).ListSubItems.Item(9).Text
    Longitudine = Replace(Longitudine, ",", ".")
    Cap = ListView1.ListItems.Item(Riga).ListSubItems.Item(3).Text
    Indirizzo = ListView1.ListItems.Item(Riga).ListSubItems.Item(2).Text
    Civico = ""
    Indirizzo = Replace(Longitudine, " ", "+")
    Città = ListView1.ListItems.Item(Riga).ListSubItems.Item(4).Text
    Città = Replace(Longitudine, " ", "+")

    If LCase(SitoWeb) = "multimap" Then
        MostraPosizione = "http://www.multimap.com/map/browse.cgi?" & "lat=" & Latitudine & "&lon=" & Longitudine & "&scale=5000&icon=x"
    ElseIf LCase(SitoWeb) = "mapquest" Then
        MostraPosizione = "http://www.mapquest.com/maps/map.adp?latlongtype=decimal&latitude=" & Latitudine & "&longitude=" & Longitudine & "&zoom=9"
    ElseIf LCase(SitoWeb) = "map24" Then
        MostraPosizione = "http://www.it.map24.com/"
    ElseIf LCase(SitoWeb) = "www16.mappy" Then
        MostraPosizione = "http://www16.mappy.com/sid0D9k5LSC7f3Mn21w/AFGM?recherche=0&posl=poi&show_poi=1&ids=0&poix=0&poiy=0&poi_rr=0.5&poi_rx=0.6&poi_ry=0.5&csl=m1&fsl=m1&gsl=m1&msl=m1&temp_no_prop=0&comment=&xsl=1&out=2&wnm1=" & Indirizzo & "&wcm1=&nom1=&tcm1=%3Ba10m1%3D130644&tnm1=" & Città & "&pcm1=" & Cap & "&ccm1=380"
    ElseIf LCase(SitoWeb) = "googlemaps" Then
        MostraPosizione = "http://maps.google.it/maps?f=q&hl=it&q=" & Latitudine & "N+" & Longitudine & "E&ie=UTF8&om=1&z=12"
    End If
    
    Exit Function
    
Esci:
    MostraPosizione = ""
    Screen.MousePointer = vbDefault
    Exit Function
    
End Function

Public Function MousePosizioneFinestra(miaForm As Form) As POINTAPI
    Dim Posizione As POINTAPI
    Dim PuntoX As Integer
    Dim PuntoY As Integer
    'per ricavare la posizione x e y
    'PuntoX = MousePosizioneFinestra.x 'coordinata del punto x relativa al form corrente
    'PuntoY = MousePosizioneFinestra.y 'coordinata del punto y relativa al form corrente

    ' Restituisce la posizione x,y relativamente allo schermo
    GetCursorPos Posizione
    
     ' Converte la posizione x,y relativamente al form specificata (.hWnd)
    ScreenToClient miaForm.hwnd, Posizione

    MousePosizioneFinestra = Posizione
    
End Function

Public Sub MouseMove(miaForm As Form, X As Single, Y As Single)
    Dim pt As POINTAPI
    Dim hwnd As Long
    
    hwnd = miaForm.hwnd
    
    pt.X = X
    pt.Y = Y
    ClientToScreen hwnd, pt
    SetCursorPos pt.X, pt.Y
    
End Sub

Public Function TempoTrascorso(dtaDataIniziale As Date, dtaDataFinale As Date, Optional RisultatoInSecondi As Boolean = False) As Variant
    Dim lGiorni As Integer
    Dim lOre As Integer
    Dim lMinuti As Integer
    Dim lSecondi As Integer
    Dim dblSec As Double
    ' Esempio:
    ' Dim dtaTempo As Date
    '
    ' dtaTempo = Now
    ' [...]
    ' strTempo = TempoTrascorso(dtaTempo, Now)

    dblSec = DateDiff("s", dtaDataIniziale, dtaDataFinale)
    lSecondi = dblSec Mod 60
    lMinuti = (dblSec \ 60) Mod 60
    lOre = ((dblSec \ 60) \ 60) Mod 24
    lGiorni = Int(((dblSec \ 60) \ 60) \ 24)
    
    If RisultatoInSecondi = False Then
        TempoTrascorso = Format$(lGiorni, "00") & "." & Format$(lOre, "00") & ":" & Format$(lMinuti, "00") & ":" & Format$(lSecondi, "00")
    Else
        TempoTrascorso = dblSec
    End If
   
End Function

Public Function StimaTempoRestante(SecondiTrascorsi As Double, RecElaborati As Long, RecTotali As Long) As String
    
    If RecElaborati <> 0 And RecTotali <> 0 Then
       StimaTempoRestante = ConvertiSecInGiorni((SecondiTrascorsi / CDbl(RecElaborati) * CDbl(RecTotali)) - SecondiTrascorsi)
    Else
        StimaTempoRestante = "00.00.00.00"
    End If

End Function

Public Function ConvertiSecInGiorni(Secondi As Double) As String
    Dim lGiorni As Integer
    Dim lOre As Integer
    Dim lMinuti As Integer
    Dim lSecondi As Integer
    Dim dblSec As Double
    
    dblSec = Secondi
    lSecondi = dblSec Mod 60
    lMinuti = (dblSec \ 60) Mod 60
    lOre = ((dblSec \ 60) \ 60) Mod 24
    lGiorni = Int(((dblSec \ 60) \ 60) \ 24)

    ConvertiSecInGiorni = Format$(lGiorni, "00") & "." & Format$(lOre, "00") & ":" & Format$(lMinuti, "00") & ":" & Format$(lSecondi, "00")
    
End Function

Public Sub CaricaImmagine(ByVal picImmagine As PictureBox, ByVal PosizioneImmagine As String, Optional ByVal Estensione As String = "")
    On Error Resume Next
    
    If PosizioneImmagine = "" Then Exit Sub
    
    If Estensione <> "" Then
        PosizioneImmagine = Left$(PosizioneImmagine, Len(PosizioneImmagine) - 3) & Estensione
    End If
    
    If FileExists(PosizioneImmagine) = True Then
        picImmagine.Picture = LoadPicture(PosizioneImmagine)
    Else
        picImmagine.Picture = LoadPicture()
    End If

End Sub

Public Function SplitOne(ByVal Testo As String, ByVal Separatore As String, ByVal PosRecord As Long) As String
    ' Splitta una stringa di testo e restituisce solo il valore di PosRecord
    Dim vTmp
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    If Testo = "" Then
        SplitOne = ""
        Exit Function
    End If
    
    vTmp = Split(Testo, Separatore)
    If PosRecord <= UBound(vTmp) Then
        SplitOne = vTmp(PosRecord)
    Else
        SplitOne = ""
    End If
    
    Exit Function
    
Errore:
    GestErr Err, "Errore nella funzione SplitOne. Testo: " & Testo
    
End Function

Public Function GetValue(ByVal sTesto As String, ByVal sChiave As String, Optional ByVal sDelimitatore As String = "=") As String
    ' Dato sTesto restituisce il Valore dopo la chiave separato da delimitatore........
    Dim cnt As Integer
    Dim arrRiga As Variant
    Dim arrTesto As Variant
    
    GetValue = ""
    
    arrTesto = Split(sTesto, vbCrLf, , vbTextCompare)
    
    For cnt = 0 To UBound(arrTesto)
        If arrTesto(cnt) = "" Then Exit For
        arrRiga = Split(arrTesto(cnt), sDelimitatore, , vbTextCompare)
        If LCase(arrRiga(0)) = LCase(sChiave) Then
            GetValue = Replace(arrRiga(1), ".", ",")
            Exit For
        End If
    Next
    
End Function

Public Function GetProgVers() As String
    
    GetProgVers = App.major & "." & App.minor & "." & App.Revision

End Function

Public Function VerificaStatoInternet() As Boolean

    VerificaStatoInternet = InternetGetConnectedState(0&, 0&)

End Function

Public Function ConfermaChiusuraForm(Optional ByVal Messaggio As String = "") As Integer
    
    If Var(ConfermaUscitaForm).Valore > 0 Then
        
        Select Case Messaggio
            Case Is = "0", ""
            Messaggio = "Vuoi chiudere questa finestra adesso?    " & vbNewLine & "I dati non salvati andranno persi."
            
            Case Is = "1"
            Messaggio = "Vuoi uscire dal programma adesso?      "
                
        End Select
        
        If MsgBox(Messaggio, vbInformation + vbYesNo, App.ProductName) = vbNo Then
            ConfermaChiusuraForm = 1
        Else
            ConfermaChiusuraForm = 0
        End If
    End If

End Function
