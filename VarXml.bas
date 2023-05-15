Attribute VB_Name = "VarXml"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

Public Type tyVarXml ' Il tipo dati utilizzato nella classe clsVariabiliXml
    Indice As enNomVarXml
    Nome As String
    Valore As Variant
    Opzioni As String
    Sezione As String
    SubSezione As String
    Predefinito As String
    Descrizione As String
End Type

Public mtrVarXml() As tyVarXml

Public Enum enNomVarXml
    Zero = 0                    ' Valore Zero
    tmpFile = 1                 ' La cartella temporanea dei file .ov2
    tmpBMPfile = 2              ' La cartella temporanea dei file .bmp
    tmpDownFile = 3             ' La cartella temporanea dei file scaricati
    PoiScaricati = 4            ' La cartella che contiene i POI scaricati
    CartellaMappe = 5           ' La cartella che contiene le mappe
    CartellaScript = 44         ' La cartella che contiene gli Script
    '
    MakeAscBat = 6              ' Il percorso completo di nome del file .bat per la creazione dei file .asc
    MakeOv2Bat = 7              ' Il percorso completo di nome del file .bat per la creazione dei file .asc
    RegioniCsv = 8              ' Il percorso completo di nome del file .csv con l'elenco delle regioni d'Italia
    ProvinceCsv = 9             ' Il percorso completo di nome del file .csv con l'elenco delle province d'Italia
    TipoPdiCsv = 10             ' Il percorso completo di nome del file .csv con l'elenco di tutti i tipi di PDI
    CartellePersonali = 11      ' Elenco di cartelle impostate dagli utenti
    Sostituisci = 45            ' Elenco di corrispondenze dei caratteri da sostituire
    '
    '
    GestioneErrori = 12         ' Attiva / Disattiva la gestione degli errori
    DebugMode = 13              ' Scrittura del file Debug.log - 0 Disattivato
    UsaManifestFile = 14        ' Determina se utilizzare il file manifest esterno (0 - 1)
    CampiRmkFile = 15           ' L'indicatore nei file .rmk.....
    SetupDescPred = 16          ' Le impostazioni di default per la form SetupDercrizione
    CommaSep = 17               ' Il separatore utilizzato per i file .csv
    SelezMultipla = 18          ' Per la selezione multipla delle ListView
    xElaboraNomeFile = 19       ' Determina se il nome dei file scaricati e trattati vengano elaborati oppure no
    ScaricaBMP = 20             ' Determina se scaricare il file .bmp
    EndBeep = 21                ' Determina se deve essere emesso il Beep alla fine dell'operazione (0 - 1)
    BMPnascoste = 22            ' Determina se i file .bmp devono essere nascosti dopo la copia nella cartella finale (questo serve per l'utilizzo dei poi nei cellulari Nokia per non intasare la galleria immagini)
    VisualizzaLog = 23          ' Visualizza sempre il file log
    LimiteLon = 24              ' Il limite minimo e massimo della longitudine
    LimiteLat = 25              ' Il limite minimo e massimo della latitudine
    AutoVerificaCoordinate = 26 ' Imposta se eseguire la verifica automatica delle coordinate
    ConfermaUscitaForm = 27     ' Messaggio di conferma all'uscita delle form
    '
    PoiGpsXmlWeb = 28           ' La pagina web del sito dove risiede l'elenco dei file in xml
    PoiGpsWebWeb = 29           ' La pagina web del sito che contiene il link a tutti i file
    '
    SalvaOgni = 30              ' Imposta ogni quante righe deve essere effettuato il backup automatico
    SalvataggioBackup = 31      ' Imposta il salvataggio automatico
    '
    ProxySet = 32               ' Le impostazioni di connessione del Proxy
    '
    PoiGpsXmlUserName = 33
    PoiGpsXmlPsw = 34
    '
    GpsBabel_In_Out = 35        ' Il formato di ingresso e quello di uscita di GPSbabel (indicare i numeri delle ListBox separati da virgola)
    '
    PhoneFileExe = 36           ' Il percorso completo di nome del file PhoneFile.exe
    PhoneFileTele = 37          ' Elenco dei telefoni separato dal carattere |
    PhoneFileWeb = 38           ' Pagina Web del programma
    PhoneFileMapDir = 39             ' Il testo da cercare nei nomi delle cartelle che contengono la mappa
    '
    UltimoFilePOI = 40          ' Il percorso completo di nome dell'ultimo file POI aperto
    UltimoFileWEB = 41          ' Il percorso completo di nome dell'ultimo file delle impostazioni di download aperto
    VersioneProgramma = 42      ' Il numero memorizzzato nel file .xml che contiene il numero della versione del programma avviata l'ultima volta
    MainLastPos = 43            ' L'ultima posizione della form Main (x,y)
    
End Enum

Public Function LeggiXML(Optional ByVal Chiave As enNomVarXml = Zero, Optional ByVal ScriviValore As Variant = "") As tyVarXml
    Dim strTmp As String
    Dim NomeVar As String
    Dim SezioneVar As String
    Dim SubSezioneVar As String
    Dim Opzioni As String
    Dim Descrizione As String
    Dim cnt As Integer
    Dim OrigChiave As enNomVarXml
    
    'If Var(GestioneErrori).Valore = 0 Then On Error Resume Next

    ReDim Preserve mtrVarXml(45) ' Il numero delle variabili da memorizzare
    OrigChiave = Chiave
    
    NomeVar = "Zero"
    If Chiave = Zero Or Chiave = Zero Then
        mtrVarXml(Chiave).Indice = Zero
        mtrVarXml(Chiave).Nome = NomeVar
        mtrVarXml(Chiave).Valore = "Zero"
        mtrVarXml(Chiave).Opzioni = ""
        mtrVarXml(Chiave).Sezione = ""
        mtrVarXml(Chiave).SubSezione = ""
        mtrVarXml(Chiave).Predefinito = ""
        mtrVarXml(Chiave).Descrizione = ""
    End If

    If OrigChiave = Zero Then
        Chiave = 1
    End If

    ' Inizio lettura ----------------------------------------------------------------------------------------------------------------------------
    NomeVar = "tmpFile"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = tmpFile Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.path & "\tmpFile", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "tmpBMPfile"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = tmpBMPfile Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.path & "\tmpBMPfile", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "tmpDownFile"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = tmpDownFile Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.path & "\tmpDownFile", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "PoiScaricati"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = PoiScaricati Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.path & "\PoiScaricati", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "CartellaMappe"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = " Cartella che contiene le mappe utilizzate dal programma "
    If Chiave = Zero Or Chiave = CartellaMappe Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.path & "\mappe", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    
    ' Cartella Script
    
    
    
    
    '
    '
    NomeVar = "MakeAscBat"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = MakeAscBat Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, Var(tmpFile).Valore & "\MakeAsc.bat", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "MakeOv2Bat"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = MakeOv2Bat Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, Var(tmpFile).Valore & "\MakeOv2.bat", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "RegioniCsv"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = RegioniCsv Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.path & "\Regioni.csv", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "ProvinceCsv"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = ProvinceCsv Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.path & "\Province.csv", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "TipoPdiCsv"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = TipoPdiCsv Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.path & "\TipoPdi.csv", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "CartellePersonali"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = "Qua puoi inserire un elenco delle cartelle personali che il programma potrà utilizzare"
    If Chiave = Zero Or Chiave = CartellePersonali Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    '
    '
    NomeVar = "GestioneErrori"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "0||1"
    Descrizione = ""
    If Chiave = Zero Or Chiave = GestioneErrori Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "0", ScriviValore, Opzioni
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "DebugMode"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "0||1"
    Descrizione = ""
    If Chiave = Zero Or Chiave = DebugMode Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "0", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "UsaManifestFile"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "0||1"
    Descrizione = ""
    If Chiave = Zero Or Chiave = UsaManifestFile Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "1", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "CampiRmkFile"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = CampiRmkFile Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "@RemakeOv2@", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "SetupDescPred"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = SetupDescPred Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, ",,,,,),,,>,,[,],(,,,,1,0,0,0,1,0,0,1,,,,1,0", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "CommaSep"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = CommaSep Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, ";", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "SelezMultipla"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "Vero||Falso"
    Descrizione = ""
    If Chiave = Zero Or Chiave = SelezMultipla Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "Falso", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "xElaboraNomeFile"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "Vero||Falso"
    Descrizione = ""
    If Chiave = Zero Or Chiave = xElaboraNomeFile Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "Falso", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "ScaricaBMP"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "Vero||Falso"
    Descrizione = ""
    If Chiave = Zero Or Chiave = ScaricaBMP Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "Vero", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "EndBeep"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "0||1"
    Descrizione = ""
    If Chiave = Zero Or Chiave = EndBeep Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "1", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "BMPnascoste"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "0||1"
    Descrizione = ""
    If Chiave = Zero Or Chiave = BMPnascoste Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "0", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "VisualizzaLog"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "0||1"
    Descrizione = "Se 0 il file log viene visualizzato solo in caso di errori."
    If Chiave = Zero Or Chiave = VisualizzaLog Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "0", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "LimiteLon"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = LimiteLon Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "6,60|18,52", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "LimiteLat"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = LimiteLat Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "35,50|47,09", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "AutoVerificaCoordinate"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "0||1||2"
    Descrizione = "0 = Disattivato" & vbNewLine & "1 = Solo in caso di errori" & vbNewLine & "2 = Attivato sempre"
    If Chiave = Zero Or Chiave = AutoVerificaCoordinate Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "1", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "ConfermaUscitaForm"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Generali"
    Opzioni = "0||1"
    Descrizione = "0 = Disattivata" & vbNewLine & "1 = Attivata"
    If Chiave = Zero Or Chiave = ConfermaUscitaForm Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "1", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    '
    '
    NomeVar = "PoiGpsXmlWeb"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Web"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = PoiGpsXmlWeb Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "http://www.poigps.com/poi.xml", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "PoiGpsWebWeb"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Web"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = PoiGpsWebWeb Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "http://www.poigps.com/modules.php?name=Downloads&d_op=getit&lid=", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    '
    '
    NomeVar = "SalvaOgni"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Salvataggio"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = SalvaOgni Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "40", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "SalvataggioBackup"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Salvataggio"
    Opzioni = "Vero||Falso"
    Descrizione = ""
    If Chiave = Zero Or Chiave = SalvataggioBackup Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "Falso", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    '
    '
    NomeVar = "ProxySet"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Proxy"
    Opzioni = "0"
    Descrizione = ""
    If Chiave = Zero Or Chiave = ProxySet Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "0", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    '
    '
    NomeVar = "PoiGpsXmlUserName"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Utente"
    Opzioni = "poigpsuser"
    Descrizione = ""
    If Chiave = Zero Or Chiave = PoiGpsXmlUserName Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "poigpsuser", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "PoiGpsxmlPsw"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Utente"
    Opzioni = "99554396671831"
    Descrizione = ""
    If Chiave = Zero Or Chiave = PoiGpsXmlPsw Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "99554396671831", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    '
    '
    NomeVar = "GpsBabel_In_Out"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "GpsBabel_In_Out"
    Opzioni = ""
    Descrizione = "Impostazioni dei file in entrata / uscita di GpsBabel"
    If Chiave = Zero Or Chiave = GpsBabel_In_Out Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "79,70", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    '
    '
    NomeVar = "PhoneFileExe"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "PhoneFile"
    Opzioni = "°file°°°||*.exe"
    Descrizione = ""
    If Chiave = Zero Or Chiave = PhoneFileExe Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "C:\Programmi\PhoneFile\PhoneFile.exe", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "PhoneFileTele"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "PhoneFile"
    Opzioni = "°cancel°"
    Descrizione = ""
    If Chiave = Zero Or Chiave = PhoneFileTele Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "PhoneFileWeb"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "PhoneFile"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = PhoneFileWeb Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "http://www.devlex.com/Products/PhoneFile.shtml", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "PhoneFileMapDir"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "PhoneFile"
    Opzioni = ""
    Descrizione = "Il testo da cercare nei nomi delle cartelle che contengono la mappa"
    If Chiave = Zero Or Chiave = PhoneFileMapDir Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "Map", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    '
    '
    NomeVar = "UltimoFilePOI"
    SezioneVar = "Varie"
    SubSezioneVar = "Generali"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = UltimoFilePOI Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "UltimoFileWEB"
    SezioneVar = "Varie"
    SubSezioneVar = "Generali"
    Opzioni = ""
    Descrizione = ""
    If Chiave = Zero Or Chiave = UltimoFileWEB Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "VersioneProgramma"
    SezioneVar = "Varie"
    SubSezioneVar = "Generali"
    Opzioni = ""
    Descrizione = "Alla chiusura viene memorizzata la versione del programma"
    If Chiave = Zero Or Chiave = VersioneProgramma Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.major & "." & App.minor & "." & App.Revision, ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    NomeVar = "MainLastPos"
    SezioneVar = "Varie"
    SubSezioneVar = "Generali"
    Opzioni = "°cancel°"
    Descrizione = "L'ultima posizione della finestra principale del programma"
    If Chiave = Zero Or Chiave = MainLastPos Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, "", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    '
    '
    '
    NomeVar = "CartellaScript"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = " Cartella che contiene gli script utilizzati dal programma "
    If Chiave = Zero Or Chiave = CartellaScript Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.path & "\Scripts", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1
    '
    '
    '
    '
    '
    NomeVar = "Sostituisci"
    SezioneVar = "Impostazioni"
    SubSezioneVar = "Directory"
    Opzioni = ""
    Descrizione = "File contenente le corrispondenze dei caratteri che verranno trattati se è attiva la funzione xElaboraNomeFile"
    If Chiave = Zero Or Chiave = Sostituisci Then
        intLeggiXML SezioneVar, SubSezioneVar, NomeVar, Chiave, XmlFileConfig, App.path & "\Sostituisci.csv", ScriviValore, Opzioni, Descrizione
        If OrigChiave <> Zero Then GoTo Esci
    End If
    If OrigChiave = Zero Then Chiave = Chiave + 1


    
Esci:
    If Chiave <> Zero Then
        LeggiXML.Indice = mtrVarXml(OrigChiave).Indice
        LeggiXML.Nome = mtrVarXml(OrigChiave).Nome
        LeggiXML.Valore = mtrVarXml(OrigChiave).Valore
        LeggiXML.Opzioni = mtrVarXml(OrigChiave).Opzioni
    End If
    
End Function

Private Sub intLeggiXML(ByVal Sezione As String, ByVal SubSezione As String, ByVal NomeVar As String, ByVal Chiave As enNomVarXml, ByVal PosFileXml, ByVal Predefinito As String, Optional ByVal ScriviValore As String = "", Optional ByVal Opzioni As String, Optional ByVal Descrizione As String)
    Dim strValore As String
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    If ScriviValore <> "" Then
        ScriviChiaveXML Sezione, SubSezione, NomeVar, ScriviValore, PosFileXml, False
    End If
    
    mtrVarXml(Chiave).Indice = Chiave
    mtrVarXml(Chiave).Nome = NomeVar
    
    strValore = LeggiChiaveXML(Sezione, SubSezione, NomeVar, PosFileXml, Predefinito)
    
    Select Case LCase(strValore)
        Case "false", "true", "falso", "vero"
            mtrVarXml(Chiave).Valore = CBool(strValore)
        
        Case (IsNumeric(strValore)) = True
            If (strValore <= 2147483647) And (InStr(1, strValore, ",", vbTextCompare) = 0) And (InStr(1, strValore, ".", vbTextCompare) = 0) Then
                mtrVarXml(Chiave).Valore = CLng(strValore)
            Else
                mtrVarXml(Chiave).Valore = CStr(strValore)
            End If
        
        Case Else
            mtrVarXml(Chiave).Valore = CStr(strValore)
    End Select
    
    mtrVarXml(Chiave).Opzioni = Opzioni
    mtrVarXml(Chiave).Sezione = Sezione
    mtrVarXml(Chiave).SubSezione = SubSezione
    mtrVarXml(Chiave).Predefinito = Predefinito
    mtrVarXml(Chiave).Descrizione = Descrizione
    
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione intLeggiXML." & vbNewLine & Sezione & " -*- " & SubSezione & " -*- " & NomeVar
    
End Sub

Public Property Let lVar(ByVal Chiave As enNomVarXml, Valore As Variant)
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    LeggiXML Chiave, Valore
    Exit Property
Errore:
    GestErr Err, "Errore nella funzione PropertyLetVar del modulo VarXml."
End Property

Public Property Get Var(ByVal Chiave As enNomVarXml) As tyVarXml
    On Error GoTo Errore
    Var.Indice = mtrVarXml(Chiave).Indice
    Var.Nome = mtrVarXml(Chiave).Nome
    Var.Valore = mtrVarXml(Chiave).Valore
    Var.Opzioni = mtrVarXml(Chiave).Opzioni
    Exit Property
Errore:
    GestErr Err, "Errore nella funzione PropertyGetVar del modulo VarXml."
End Property

