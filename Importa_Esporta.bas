Attribute VB_Name = "Importa_Esporta"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

Public Type Ov2FileTy
    aTy1PoiLatitude As Double
    bTy1PoiLongitude As Double
    cTy1Poi3Latitude As Double
    dTy1Poi3Longitude As Double
    eTy1Distanza As Double
    fTy2PoiLatitude As Double
    gTy2PoiLongitude As Double
    hTy2descrizione As String
    iDescrizione As String
    iIndirizzo As String
    iCap As String
    iCitta As String
    iProvincia As String
    iTelefono As String
    iCategoria As String
    iNumProgressivoRecord As Long
    fMarcato As Boolean
End Type

Private Type PoiRecOld
    type As Byte
    Length As Long
    longitude As Long
    latitude As Long
End Type

Private Type PoiRec
    type As Byte
    Length As Long
End Type

Private Type PoiRec1
    longitude1 As Long
    latitude1 As Long
    longitude2 As Long
    latitude2 As Long
End Type

Private Type PoiRec2
    longitude As Long
    latitude As Long
End Type

Const pi = 3.141592654
Const AxisMajor = 6378137
Const AxisMinor = 6356752.3142
Const Elevation = 200

Public ArrayOv2PoiRec() As Ov2FileTy
Public NomeFileAperto As String 'Contiene il nome del file aperto, è utilizzata da altre funzioni per sapere quale è il nome dell'ultimo file aperto
Public PatchNomeFileAperto As String 'Contiene il nome e la posizione del file aperto, è utilizzata da altre funzioni per sapere quale è il nome dell'ultimo file aperto

Dim clsCmd As New clsCommonDialog

Public Function ExportaDati(FormHandle As Long, Optional PatchNomeFile As String = "", Optional NomeFile As String = "Export", Optional ByVal sEstensioni As String = "", Optional ByVal MessaggioRisultati As Boolean = False, Optional ByVal DefaultExtension As String = "*.rmk", Optional ByVal sCampiRmkFile As String = "", Optional ListView As ListView, Optional ByVal SaveCheck As Boolean = False) As Long
    ' Restituisce:
    '   -1 in caso di errore oppure
    '   un numero che contiene il numero dei record esportati
    '
    ' sEstensioni = un stringa contenente le estensioni separate dal carattere |
    '
    Dim FilterIndex As Long
    Dim Filename As String
    Dim arrTmp
    Dim NumeroRecordExport As Long: NumeroRecordExport = 0
    Dim cnt As Integer
    Dim SepEst As String: SepEst = Chr(0)
    Dim RMK As String: RMK = "RemakeOv2 File (*.rmk)" & SepEst & "*.rmk" & SepEst
    Dim ov2 As String: ov2 = "TomTom POI OV2 (*.ov2)" & SepEst & "*.ov2" & SepEst
    Dim CSV As String: CSV = "Comma Separated Value Files (*.csv)" & SepEst & "*.csv" & SepEst
    Dim ASC As String: ASC = "ASCII Format Files (*.asc)" & SepEst & "*.asc" & SepEst
    Dim KML As String: KML = "Google Earth Files (*.kml)" & SepEst & "*.kml" & SepEst
    Dim GPX As String: GPX = "GPS Exchange Format Files (*.gpx)" & SepEst & "*.gpx" & SepEst

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    If PatchNomeFile = "" Then
        If NomeFileAperto = "" Then NomeFileAperto = "Export"
        ' se il nome del file ha l'estensione.....
        If Mid$(NomeFile, Len(NomeFile) - 3, 1) = "." Then
            Select Case LCase$(Mid$(NomeFile, Len(NomeFile) - 3, 4))
                Case Is = ".rmk"
                    FilterIndex = 1
                Case Is = ".ov2"
                    FilterIndex = 2
                Case Is = ".csv"
                    FilterIndex = 3
                Case Is = ".asc"
                    FilterIndex = 4
                Case Is = ".gpx"
                    FilterIndex = 5
            End Select
            NomeFile = Left$(NomeFile, Len(NomeFile) - 4)
        End If

        If sEstensioni = "" Then
            sEstensioni = RMK & ov2 & CSV & ASC & KML & GPX & SepEst
        Else
            arrTmp = Split(sEstensioni, "|")
            sEstensioni = ""
            For cnt = 0 To UBound(arrTmp)
                Select Case UCase(arrTmp(cnt))
                    Case Is = "OV2"
                        sEstensioni = sEstensioni & ov2
                    Case Is = "RMK"
                        sEstensioni = sEstensioni & RMK
                    Case Is = "CSV"
                        sEstensioni = sEstensioni & CSV
                    Case Is = "ASC"
                        sEstensioni = sEstensioni & ASC
                    Case Is = "KML"
                        sEstensioni = sEstensioni & KML
                    Case Is = "GPX"
                        sEstensioni = sEstensioni & GPX
                End Select
            Next
            sEstensioni = sEstensioni & SepEst
        End If

        ' The filter entries must be seperated by nulls and terminated with two nulls
        clsCmd.Filter = sEstensioni
        If FilterIndex = Null Then FilterIndex = 2
        clsCmd.FilterIndex = FilterIndex  ' Imposta la posizione predefinita della combo con il file da aprire
        clsCmd.Filename = NomeFile
        clsCmd.DefaultExtension = DefaultExtension
        clsCmd.DialogTitle = "Salva con nome"
        clsCmd.hwnd = FormHandle             ' The dialog will not be modal unless a hWnd is specified.
        'clsCmd.InitDir = App.Path
        clsCmd.Flags = StandardFlag.SaveFile    ' not multiselect
        clsCmd.ShowSave
        If clsCmd.CancelPressed = False Then
            PatchNomeFile = clsCmd.Filename
            Filename = clsCmd.Filename
        Else
            ExportaDati = 0
            Exit Function
        End If
    Else
        Filename = PatchNomeFile
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents

    If PatchNomeFile <> "" Then
        PatchNomeFileAperto = Filename
        NomeFileAperto = FileNameFromPath(Filename)
        Select Case Right$(PatchNomeFile, 4)
            Case ".gpx"
                ArrayOv2PoiRecDaListView ListView, ArrayOv2PoiRec
                NumeroRecordExport = ExportGPXb(PatchNomeFile, MessaggioRisultati)
            Case ".asc"
                ArrayOv2PoiRecDaListView ListView, ArrayOv2PoiRec
                NumeroRecordExport = ExportASC(PatchNomeFile, MessaggioRisultati)
            Case ".ov2"
                ArrayOv2PoiRecDaListView ListView, ArrayOv2PoiRec
                NumeroRecordExport = ExportOV2code(PatchNomeFile, MessaggioRisultati)
            Case ".csv"
                Call CaricaArrayDaListView(ListView, SaveCheck)
                NumeroRecordExport = SaveAsCSV(arrListView, PatchNomeFile, , Var(CommaSep).Valore, MessaggioRisultati, ".csv", 1)
            Case ".rmk"
                Call CaricaArrayDaListView(ListView, SaveCheck)
                NumeroRecordExport = SaveAsCSV(arrListView, PatchNomeFile, sCampiRmkFile, "|", MessaggioRisultati, ".rmk", 1)
            Case Else
                MsgBox "Formato estensione del file non supportato: " & NomeFileAperto, vbInformation, App.ProductName
        End Select
        ExportaDati = NumeroRecordExport
    Else
        ExportaDati = 0
    End If

    Screen.MousePointer = vbDefault

    Exit Function
    
Errore:
    ExportaDati = -1
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione ExportaDati."

End Function

Private Function ExportGPXa(Filename As String, Optional Messaggio As Boolean = False) As Integer
    Dim cntRecord As Long
    Dim NumRows As Long
    Dim R As Long
    Dim Data
    
    Open Filename For Output As #1
    
    Data = ""
    Data = Data & "<?xml version='1.0' encoding='ISO-8859-1' standalone='yes'?> " & vbCrLf
    Data = Data & "<gpx version='1.1' "
    Data = Data & "creator='SkipperOV2' "
    Data = Data & "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' "
    Data = Data & "xmlns='http://www.topografix.com/GPX/1/1' "
    Data = Data & "xsi:schemaLocation='http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd'> " & vbCrLf
    Data = Data & "<metadata>" & vbCrLf
    Data = Data & "<name>" & "TovaGPS" & "</name>" & vbCrLf
    Data = Data & "<desc>" & "TovaGPS" & "</desc>" & vbCrLf
    Data = Data & "<link href='http://www.garmin.com'>" & vbCrLf
    Data = Data & "<text>Garmin International</text>" & vbCrLf
    Data = Data & "</link>" & vbCrLf
    Data = Data & "</metadata>" & vbCrLf
    Print #1, Data
    
    R = 0
    cntRecord = 0
    
    While Trim$(ArrayOv2PoiRec(R).aTy1PoiLatitude) <> 0 Or Trim$(ArrayOv2PoiRec(R).fTy2PoiLatitude) <> 0
        If Trim$(ArrayOv2PoiRec(R).aTy1PoiLatitude) <> "" Then
            Print #1, "<trk><trkseg>"
            Data = "<trkpt lat='" & Replace(ArrayOv2PoiRec(R).aTy1PoiLatitude, ",", ".")
            Data = Data & "' lon='" & Replace(ArrayOv2PoiRec(R).bTy1PoiLongitude, ",", ".") & "' >"
            Data = Data & "<type>Trackpoint</type></trkpt>"
            
            Print #1, Data
            Data = "<trkpt lat='" & Replace(ArrayOv2PoiRec(R).cTy1Poi3Latitude, ",", ".")
            Data = Data & "' lon='" & Replace(ArrayOv2PoiRec(R).bTy1PoiLongitude, ",", ".") & "' >"
            Data = Data & "<type>Trackpoint</type></trkpt>"
            
            Print #1, Data
            Data = "<trkpt lat='" & Replace(ArrayOv2PoiRec(R).cTy1Poi3Latitude, ",", ".")
            Data = Data & "' lon='" & Replace(ArrayOv2PoiRec(R).dTy1Poi3Longitude, ",", ".") & "' >"
            Data = Data & "<type>Trackpoint</type></trkpt>"
            
            Print #1, Data
            Data = "<trkpt lat='" & Replace(ArrayOv2PoiRec(R).aTy1PoiLatitude, ",", ".")
            Data = Data & "' lon='" & Replace(ArrayOv2PoiRec(R).dTy1Poi3Longitude, ",", ".") & "' >"
            Data = Data & "<type>Trackpoint</type></trkpt>"
            
            Print #1, Data
            Data = "<trkpt lat='" & Replace(ArrayOv2PoiRec(R).aTy1PoiLatitude, ",", ".")
            Data = Data & "' lon='" & Replace(ArrayOv2PoiRec(R).bTy1PoiLongitude, ",", ".") & "' >"
            Data = Data & "<type>Trackpoint</type></trkpt>"
            
            Print #1, Data
            
            Print #1, "</trkseg></trk>"
            cntRecord = cntRecord + 1
        End If
        R = R + 1
    Wend
    
    Print #1, "<extensions>"
    Print #1, "</extensions>"
    Print #1, "</gpx>"
    
    Close #1
    
    If Messaggio = True Then MsgBox (" " & cntRecord & " record scritti ")
    ExportGPXa = cntRecord
    
End Function

Public Function ExportGPXb(Filename, Optional Messaggio As Boolean = False) As Long
Attribute ExportGPXb.VB_Description = "Salva File GPX"
Attribute ExportGPXb.VB_ProcData.VB_Invoke_Func = "x\n14"
    Dim cntRecord As Long
    Dim NumRows As Long
    Dim i As Long
    Dim Data
    
    Open Filename For Output As #1
    Data = ""
    Data = Data & "<?xml version='1.0' encoding='ISO-8859-1' standalone='yes'?> " & vbCrLf
    Data = Data & "<gpx version='1.1' "
    Data = Data & "creator='TovaGPS' "
    Data = Data & "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' "
    Data = Data & "xmlns='http://www.topografix.com/GPX/1/1' "
    Data = Data & "xsi:schemaLocation='http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd'> " & vbCrLf
    Data = Data & "<metadata>" & vbCrLf
    Data = Data & "<name>" & "TovaGPS" & "</name>" & vbCrLf
    Data = Data & "<desc>" & "TovaGPS" & "</desc>" & vbCrLf
    Data = Data & "<link href='http://www.garmin.com'>" & vbCrLf
    Data = Data & "<text>Garmin International</text>" & vbCrLf
    Data = Data & "</link>" & vbCrLf
    Data = Data & "</metadata>" & vbCrLf
    Print #1, Data
    
    i = 0
    cntRecord = 0

    For i = 0 To UBound(ArrayOv2PoiRec)
        If Trim$(ArrayOv2PoiRec(i).fTy2PoiLatitude) <> 0 And Trim$(ArrayOv2PoiRec(i).gTy2PoiLongitude) <> 0 Then
            Data = "<wpt lat='" & Replace(ArrayOv2PoiRec(i).fTy2PoiLatitude, ",", ".")
            Data = Data & "' lon='" & Replace(ArrayOv2PoiRec(i).gTy2PoiLongitude, ",", ".") & "' >"
            Print #1, Data
            
            If Trim$(ArrayOv2PoiRec(i).hTy2descrizione) <> "" Then
                Data = Trim$(ArrayOv2PoiRec(i).hTy2descrizione)
            Else
                Data = ArrayOv2PoiRec(i).hTy2descrizione & " ( "
                Data = Data & ArrayOv2PoiRec(i).fTy2PoiLatitude & " - "
                Data = Data & ArrayOv2PoiRec(i).gTy2PoiLongitude & " ) "
            End If
            
            Data = Replace(Data, "&", "&amp;")
            Data = Replace(Data, "'", "&apos;")
            Data = Replace(Data, "<", "&lt;")
            Data = Replace(Data, ">", "&gt;")
            Data = "<name>" & Data & "</name></wpt>"
            Print #1, Data
            cntRecord = cntRecord + 1
        End If
    Next
    
    Print #1, "<extensions>"
    Print #1, "</extensions>"
    Print #1, "</gpx>"
    Close #1
    
    If Messaggio = True Then MsgBox (" " & cntRecord & " record scritti ")
    ExportGPXb = cntRecord
    
End Function

Public Function ExportASC(Filename, Optional Messaggio As Boolean = False) As Long
    ' Ok funziona
    ' Mancherebbe solo questa parte:
    '  _ deleted bytes ignored, _ bytes not processed
    ' Readable locations in "
    Dim cntRecord As Long
    Dim i As Long
    Dim lData
    Dim Data As String
    
    Open Filename For Output As #1
    
    Data = "; Readable locations in " & App.path & vbNewLine
    Data = Data & ";" & vbNewLine
    Data = Data & "; Longitude,    Latitude, " & Chr(34) & "Name" & Chr(34) & vbNewLine
    Data = Data & "; ========== ============ ==================================================" & vbNewLine
    Print #1, Data
    
    i = 0
    cntRecord = 0
    
    For i = 0 To UBound(ArrayOv2PoiRec)
        If Trim$(ArrayOv2PoiRec(i).fTy2PoiLatitude) <> 0 And Trim$(ArrayOv2PoiRec(i).gTy2PoiLongitude) <> 0 Then
        
            If Trim$(ArrayOv2PoiRec(i).hTy2descrizione) <> "" Then
                Data = Trim$(ArrayOv2PoiRec(i).hTy2descrizione)
            Else
                Data = ArrayOv2PoiRec(i).hTy2descrizione & " ( "
                Data = Data & ArrayOv2PoiRec(i).fTy2PoiLatitude & " - "
                Data = Data & ArrayOv2PoiRec(i).gTy2PoiLongitude & " ) "
            End If
            
            Data = Replace(Data, Chr(34), "'")
            lData = "  " & Replace(ArrayOv2PoiRec(i).gTy2PoiLongitude, ",", ".") & " ,"
            lData = lData & "   " & Replace(ArrayOv2PoiRec(i).fTy2PoiLatitude, ",", ".") & " , "
            lData = lData & """" & Data & """"
            Print #1, lData
            cntRecord = cntRecord + 1
        End If
    Next
    
    Data = vbNewLine
    Data = Data & "; " & cntRecord & " record written, _ deleted bytes ignored, _ bytes not processed"
    Print #1, Data
    
    Close #1
    
    If Messaggio = True Then MsgBox (" " & cntRecord & " record scritti ")
    ExportASC = cntRecord
    
End Function

Public Function ExportOV2code(Filename, Optional Messaggio As Boolean = False) As Long
    Dim cntRecord As Long
    Dim ii As Integer
    Dim Data As String
    Dim Poi As PoiRecOld
    Dim i As Long
    Dim bc As Byte
    Dim strMsg As String
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo errExportOV2code
    
    Open Filename For Output As #1
    Close #1
    
    Open Filename For Binary As #1
    
    i = 0
    cntRecord = 0
    
    For i = 0 To UBound(ArrayOv2PoiRec)
        If Trim$(ArrayOv2PoiRec(i).fTy2PoiLatitude) <> 0 And Trim$(ArrayOv2PoiRec(i).gTy2PoiLongitude) <> 0 Then
            Poi.type = 2
            Poi.longitude = Round(ArrayOv2PoiRec(i).gTy2PoiLongitude * 100000, 0)
            Poi.latitude = Round(ArrayOv2PoiRec(i).fTy2PoiLatitude * 100000, 0)

            If Trim$(ArrayOv2PoiRec(i).hTy2descrizione) <> "" Then
                Data = Trim$(ArrayOv2PoiRec(i).hTy2descrizione)
            Else
                Data = ArrayOv2PoiRec(i).hTy2descrizione & " ( "
                Data = Data & ArrayOv2PoiRec(i).fTy2PoiLatitude & " - "
                Data = Data & ArrayOv2PoiRec(i).gTy2PoiLongitude & " ) "
            End If

            Poi.Length = Len(Data) + 14
            Put #1, , Poi

            For ii = 1 To Len(Data)
                bc = ASC(Mid(Data, ii, 1))
                Put #1, , bc
            Next

            bc = 0
            Put #1, , bc
            cntRecord = cntRecord + 1
        End If
    Next
    Close #1

    If Messaggio = True Then
        If cntRecord = 0 Then
            ' Cancello il file perchè non contiene dati
            If FileExists(Filename) = True Then Kill Filename
            strMsg = "Non e stato possibile creare il file." & vbNewLine & "In questo tipo di file vengono esportate soltanto le righe che contengono le coordinate."
        Else
            strMsg = cntRecord & " record scritti "
        End If
        MsgBox strMsg, vbInformation, App.ProductName
    End If
    
    ExportOV2code = cntRecord
    
    Exit Function
    
errExportOV2code:
    Close #1
    ExportOV2code = -1
    GestErr Err, "Errore nella funzione ExportOV2code." & vbNewLine & "L'errore potrebbe essere nella riga: " & i + 1 & " della lista."
    
End Function

Public Sub ArrayOv2PoiRecInListView(ListView1 As ListView, ArrayOv2PoiRec1() As Ov2FileTy, Optional CancellaListView As Boolean = False)
    Dim cntRecord As Long
    Dim itmX As Variant
    Dim i As Long
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo errArrayOv2PoiRecInListView
    
    cntRecord = 0
   
    If CancellaListView = True Then
        ListView1.ListItems.Clear
        ListView1.Sorted = False
        'Così si può cancellare anche l'intestazione
        'ListView1.ColumnHeaders.Clear
        DoEvents
    End If
    
    For i = 0 To UBound(ArrayOv2PoiRec1) ' Scorro tutte le righe dell'array
        If ArrayOv2PoiRec1(i).hTy2descrizione <> "" Then
            If ArrayOv2PoiRec1(i).iNumProgressivoRecord <> 0 Then
                Set itmX = ListView1.ListItems.Add(, , Format(ArrayOv2PoiRec1(i).iNumProgressivoRecord, "00000"))
            Else
                Set itmX = ListView1.ListItems.Add(, , Format(cntRecord + 1, "00000"))
            End If
            itmX.SubItems(1) = ArrayOv2PoiRec1(i).iDescrizione
            itmX.SubItems(2) = ArrayOv2PoiRec1(i).iIndirizzo
            itmX.SubItems(3) = ArrayOv2PoiRec1(i).iCap
            itmX.SubItems(4) = ArrayOv2PoiRec1(i).iCitta
            itmX.SubItems(5) = ArrayOv2PoiRec1(i).iProvincia
            itmX.SubItems(6) = ArrayOv2PoiRec1(i).iTelefono
            itmX.SubItems(7) = ArrayOv2PoiRec1(i).iCategoria
            itmX.SubItems(8) = ArrayOv2PoiRec1(i).fTy2PoiLatitude
            itmX.SubItems(9) = ArrayOv2PoiRec1(i).gTy2PoiLongitude
            itmX.SubItems(11) = ArrayOv2PoiRec1(i).hTy2descrizione
            cntRecord = cntRecord + 1
        End If
    Next i
    
    Call AutoSizeUltimaColonna(ListView1)
    
    Exit Sub

errArrayOv2PoiRecInListView:
    GestErr Err, "Errore nella funzione ArrayOv2PoiRecInListView."

End Sub

Public Sub ArrayOv2PoiRecDaListView(ListView As ListView, ArrayOv2PoiRec1() As Ov2FileTy)
    Dim i As Long
    Dim numb As Variant

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo errArrayOv2PoiRecDaListView
    
    ' Cancello l'array
    ReDim ArrayOv2PoiRec1(ListView.ListItems.Count - 1)

    For i = 0 To ListView.ListItems.Count - 1
        ArrayOv2PoiRec1(i).aTy1PoiLatitude = 0
        ArrayOv2PoiRec1(i).bTy1PoiLongitude = 0
        ArrayOv2PoiRec1(i).cTy1Poi3Latitude = 0
        ArrayOv2PoiRec1(i).dTy1Poi3Longitude = 0
        ArrayOv2PoiRec1(i).eTy1Distanza = 0
        
        numb = Replace(Trim(ListView.ListItems(i + 1).SubItems(8)), ".", ",")
        If IsNumeric(numb) = True Then
            ArrayOv2PoiRec1(i).fTy2PoiLatitude = CDbl(numb)
        Else
            ArrayOv2PoiRec1(i).fTy2PoiLatitude = 0
        End If
        
        numb = Replace(Trim(ListView.ListItems(i + 1).SubItems(9)), ".", ",")
        If IsNumeric(numb) = True Then
            ArrayOv2PoiRec1(i).gTy2PoiLongitude = CDbl(numb)
        Else
            ArrayOv2PoiRec1(i).gTy2PoiLongitude = 0
        End If
        
        If Trim(ListView.ListItems(i + 1).SubItems(11)) = "" Then
            ArrayOv2PoiRec1(i).hTy2descrizione = Trim(ListView.ListItems(i + 1).SubItems(1))
        Else
            ArrayOv2PoiRec1(i).hTy2descrizione = Trim(ListView.ListItems(i + 1).SubItems(11))
        End If
        
        ArrayOv2PoiRec1(i).iNumProgressivoRecord = Trim(ListView.ListItems(i + 1))
        ArrayOv2PoiRec1(i).iDescrizione = Trim(ListView.ListItems(i + 1).SubItems(1))
        ArrayOv2PoiRec1(i).iIndirizzo = Trim(ListView.ListItems(i + 1).SubItems(2))
        ArrayOv2PoiRec1(i).iCap = Trim(ListView.ListItems(i + 1).SubItems(3))
        ArrayOv2PoiRec1(i).iCitta = Trim(ListView.ListItems(i + 1).SubItems(4))
        ArrayOv2PoiRec1(i).iProvincia = Trim(ListView.ListItems(i + 1).SubItems(5))
        ArrayOv2PoiRec1(i).iTelefono = Trim(ListView.ListItems(i + 1).SubItems(6))
        ArrayOv2PoiRec1(i).iCategoria = Trim(ListView.ListItems(i + 1).SubItems(7))
        
        ArrayOv2PoiRec1(i).fMarcato = False
        
    Next
    
    Exit Sub

errArrayOv2PoiRecDaListView:
    GestErr Err, "Errore nella funzione ArrayOv2PoiRecDaListView."
    
End Sub

Public Function CaricaCampiFileRmk(Filename As String) As Integer
    ' Carica i campi inseriti nel file .rmk sull'array
    ' Restituisce il numero di righe occupate dai campi inseriti all'inizio del file .rmk
    Dim cnt As Long
    Dim sLine As String
    
    ReDim arrCampiRmkFile(0)
    
    Open Filename For Input As #7
    While Not (EOF(7))
        Line Input #7, sLine
        If Left$(sLine, Len(Var(CampiRmkFile).Valore)) = Var(CampiRmkFile).Valore Then
            arrCampiRmkFile() = Split(Mid$(sLine, Len(Var(CampiRmkFile).Valore) + 1), ";", , vbTextCompare)
            cnt = cnt + 1
            CaricaCampiFileRmk = cnt
        End If
    Wend
    Close #7
   
End Function

Public Function ImportaDati(FormHandle As Long, Optional ByVal ListView As ListView, Optional ByRef Filename As String = "", Optional ByVal sEstensioni As String = "", Optional ByVal SplitDescrizione As Boolean = False, Optional ByVal CancellaListView As Boolean = True, Optional ByVal InitDir As String = "", Optional ByVal DefaultExtension As String = "*.rmk", Optional ByVal SaveCheck As Boolean = False, Optional ByVal TipoFileRmk As String = "", Optional ByVal CalcolaDistanza As Boolean = False, Optional ByVal AccodaFile As Integer = 0) As Boolean
    ' Restituisce True se il file è stato caricato
    '
    ' sEstensioni = un stringa contenente le estensioni separate dal carattere |
    '
    Dim nCampi As Integer
    Dim NomeFileDefault As String
    Dim FilterIndex As Long
    Dim arrTmp
    Dim cnt As Integer
    Dim tmpSetupDescrizione As String ' Salva momentaneamente il SetupDescrizione del file .rmk
    Dim SepEst As String: SepEst = Chr(0)
    Dim RMK As String: RMK = "RemakeOv2 File (*.rmk)" & SepEst & "*.rmk" & SepEst
    Dim ov2 As String: ov2 = "TomTom POI OV2 (*.ov2)" & SepEst & "*.ov2" & SepEst
    Dim CSV As String: CSV = "Comma Separated Value Files (*.csv)" & SepEst & "*.csv" & SepEst
    Dim ASC As String: ASC = "ASCII Format Files (*.asc)" & SepEst & "*.asc" & SepEst
    Dim KML As String: KML = "Google Earth Files (*.kml)" & SepEst & "*.kml" & SepEst
    Dim GPX As String: GPX = "GPS Exchange Format Files (*.gpx)" & SepEst & "*.gpx" & SepEst
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    If sEstensioni = "" Then
        sEstensioni = RMK & ov2 & CSV & ASC & KML & GPX & SepEst
    Else
        arrTmp = Split(sEstensioni, "|")
        sEstensioni = ""
        For cnt = 0 To UBound(arrTmp)
            Select Case UCase(arrTmp(cnt))
                Case Is = "RMK"
                    sEstensioni = sEstensioni & RMK
                Case Is = "OV2"
                    sEstensioni = sEstensioni & ov2
                Case Is = "CSV"
                    sEstensioni = sEstensioni & CSV
                Case Is = "ASC"
                    sEstensioni = sEstensioni & ASC
                Case Is = "KML"
                    sEstensioni = sEstensioni & KML
                Case Is = "GPX"
                    sEstensioni = sEstensioni & GPX
            End Select
        Next
        sEstensioni = sEstensioni & SepEst
    End If
    
    If InitDir = "" Then
        InitDir = App.path
    Else
        NomeFileDefault = FileNameFromPath(InitDir, True)
    End If
    
    If Left$(DefaultExtension, 2) <> "*." Then DefaultExtension = "*.rmk"
    
    '''''' NON FUNZIONA SE VIENE UTILIZZATO sEstensioni
    Select Case LCase$(DefaultExtension)
        Case Is = "*.rmk"
            FilterIndex = 1
        Case Is = "*.ov2"
            FilterIndex = 2
        Case Is = "*.csv"
            FilterIndex = 3
        Case Is = "*.asc"
            FilterIndex = 4
        Case Is = "*.kml"
            FilterIndex = 5
        Case Is = "*.gpx"
            FilterIndex = 6
    End Select

    If Filename = "" Then
ApriAncora:
        ' The filter entries must be seperated by nulls and terminated with two nulls
        clsCmd.Filter = sEstensioni
        clsCmd.FilterIndex = FilterIndex ' Imposta la posizione predefinita della combo con il file da aprire
        clsCmd.Filename = NomeFileDefault
        clsCmd.DefaultExtension = DefaultExtension
        clsCmd.DialogTitle = "Apri file"
        clsCmd.hwnd = FormHandle      ' the dialog will not be modal unless a hWnd is specified.
        If InitDir <> "" Then clsCmd.InitDir = InitDir
        clsCmd.Flags = StandardFlag.OpenFile
        clsCmd.ShowOpen
        If clsCmd.CancelPressed = False Then
            Filename = clsCmd.Filename
        Else
            ImportaDati = False
            NomeFileAperto = ""
            Exit Function
        End If
    End If

    Screen.MousePointer = vbHourglass
    DoEvents
   
    If Dir(Filename) <> "" Then
        
        If AccodaFile = 0 Then
            PatchNomeFileAperto = Filename
            NomeFileAperto = FileNameFromPath(Filename)
            ' Cancello la variabile
            SetupDescrizione = ""
        Else
            If LCase$(Right$(PatchNomeFileAperto, 4)) = ".rmk" Then
                Load frmSetupDescrizione
                tmpSetupDescrizione = frmSetupDescrizione.PreparaSetupDescrizione(0)
                Unload frmSetupDescrizione
            End If
        End If
        
        Select Case LCase$(Right$(Filename, 4))
        
            Case ".lnk"
                InitDir = GetLinkInfo(Filename)
                NomeFileAperto = ""
                GoTo ApriAncora
                
            Case ".kml"
                ImportKML Filename, SplitDescrizione
                If Not ListView Is Nothing Then Call ArrayOv2PoiRecInListView(ListView, ArrayOv2PoiRec, CancellaListView)
                ImportaDati = True
                
            Case ".gpx"
                ImportGPX Filename, SplitDescrizione
                If Not ListView Is Nothing Then Call ArrayOv2PoiRecInListView(ListView, ArrayOv2PoiRec, CancellaListView)
                ImportaDati = True
                
            Case ".asc"
                ImportASC Filename, SplitDescrizione
                Call ArrayOv2PoiRecInListView(ListView, ArrayOv2PoiRec, CancellaListView)
                ImportaDati = True
                
            Case ".ov2"
                ImportOV2 Filename, SplitDescrizione, CalcolaDistanza
                If Not ListView Is Nothing Then Call ArrayOv2PoiRecInListView(ListView, ArrayOv2PoiRec, CancellaListView)
                ImportaDati = True
                
            Case ".rmk"
                If Not ListView Is Nothing Then
                    ' Carico i campi inseriti nel file .rmk sull'array
                    nCampi = CaricaCampiFileRmk(Filename)
                    If (arrCampiRmkFile(0) = "FileDatiPOI") And (TipoFileRmk = "FileDatiPOI" Or TipoFileRmk = "") Then
                        arrListView() = ImportCSVinArray(Filename, "|", False, nCampi)
                        If arrListViewVuoto = False Then Call CaricaListViewDaArray(ListView, CancellaListView)
                        ImportaDati = True
                        
                    ElseIf (arrCampiRmkFile(0) = "FileImpostazioniDownloadWeb") And (TipoFileRmk = "FileImpostazioniDownloadWeb" Or TipoFileRmk = "") Then
                        arrListView() = ImportCSVinArray(Filename, "|", False, nCampi)
                        If arrListViewVuoto = False Then Call CaricaListViewDaArray(ListView, CancellaListView, SaveCheck)
                        ImportaDati = True
                    End If
                Else
                    ImportaDati = False
                End If
                
            Case ".csv"
                If Not ListView Is Nothing Then
                    ' Imposto le variabili
                    Set frmImporta.ListViewDest = ListView
                    frmImporta.CancellaListView = CancellaListView
                    frmImporta.sFilename = Filename
                    ' Apro la form
                    frmImporta.Show
                    ImportaDati = True
                Else
                    ImportaDati = False
                End If
                
                
            Case Else
                MsgBox ("Formato del file non supportato!" & vbNewLine & "Puoi caricare i file .rmk .ov2 .csv .kml .gpx .asc")
                ImportaDati = False
        End Select
        
        ' Se il file viene aggiunto ad un file già aperto leggo le impostazioni del primo file aperto
        If (AccodaFile = 1) And (PatchNomeFileAperto <> "") And (LCase$(Right$(PatchNomeFileAperto, 4)) = ".rmk") Then
            CaricaCampiFileRmk (PatchNomeFileAperto)
            SetupDescrizione = tmpSetupDescrizione
        End If

    Else
        NomeFileAperto = ""
        ImportaDati = False
    End If
    
Esci:
    Screen.MousePointer = vbDefault
    Exit Function
    
Errore:
    GestErr Err, "Errore nella funzione ImportaDati."
    GoTo Esci
    
End Function

Public Sub ImportKML(Filename, Optional SplitDescrizione As Boolean = False)
    Dim cntRowArray As Long
    Dim vPoint
    Dim sDescr As String
    Dim Latitudine As String
    Dim Longitudine As String
    Dim xmlDoc As Object
    Dim xmlList As Object
    Dim xmlNode As Object
    
    If Filename <> "" Then
        Set xmlDoc = CreateObject("MSXML.DOMDocument")
        xmlDoc.async = False
        xmlDoc.Load (Filename)
        
        cntRowArray = 0
        ReDim ArrayOv2PoiRec(0)
        If Not xmlDoc Is Nothing Then
            Set xmlList = xmlDoc.getElementsByTagName("Placemark")
                For Each xmlNode In xmlList
                    ReDim Preserve ArrayOv2PoiRec(cntRowArray)
                    sDescr = xmlNode.firstChild().Text
                    vPoint = Split(xmlNode.lastChild().Text, ",")
                    Longitudine = vPoint(0)
                    Latitudine = vPoint(1)
                    ArrayOv2PoiRec(cntRowArray).aTy1PoiLatitude = 0
                    ArrayOv2PoiRec(cntRowArray).bTy1PoiLongitude = 0
                    ArrayOv2PoiRec(cntRowArray).cTy1Poi3Latitude = 0
                    ArrayOv2PoiRec(cntRowArray).dTy1Poi3Longitude = 0
                    ArrayOv2PoiRec(cntRowArray).eTy1Distanza = 0
                    ArrayOv2PoiRec(cntRowArray).fTy2PoiLatitude = Latitudine
                    ArrayOv2PoiRec(cntRowArray).gTy2PoiLongitude = Longitudine
                    If SplitDescrizione = True Then SplitDescr sDescr, cntRowArray
                    ArrayOv2PoiRec(cntRowArray).hTy2descrizione = PrimaMaiuscola(sDescr)
                    cntRowArray = cntRowArray + 1
                Next
        End If
        Set xmlDoc = Nothing
    End If
    
End Sub

Public Sub ImportGPX(Filename, Optional SplitDescrizione As Boolean = False)
    Dim cntRowArray As Long
    Dim sDescr As String
    Dim Latitudine As String
    Dim Longitudine As String
    Dim xmlDoc As Object
    Dim xmlList As Object
    Dim xmlNode As Object

    If Filename <> "" Then
        cntRowArray = 0
        ReDim ArrayOv2PoiRec(0)
        Set xmlDoc = CreateObject("MSXML.DOMDocument")
        xmlDoc.async = False
        xmlDoc.Load (Filename)
        If Not xmlDoc Is Nothing Then
            Set xmlList = xmlDoc.getElementsByTagName("wpt")
                For Each xmlNode In xmlList
                    ReDim Preserve ArrayOv2PoiRec(cntRowArray)
                    sDescr = xmlNode.firstChild().Text
                    Latitudine = xmlNode.Attributes(0).Text
                    Longitudine = xmlNode.Attributes(1).Text
                    ArrayOv2PoiRec(cntRowArray).aTy1PoiLatitude = 0
                    ArrayOv2PoiRec(cntRowArray).bTy1PoiLongitude = 0
                    ArrayOv2PoiRec(cntRowArray).cTy1Poi3Latitude = 0
                    ArrayOv2PoiRec(cntRowArray).dTy1Poi3Longitude = 0
                    ArrayOv2PoiRec(cntRowArray).eTy1Distanza = 0
                    ArrayOv2PoiRec(cntRowArray).fTy2PoiLatitude = Latitudine
                    ArrayOv2PoiRec(cntRowArray).gTy2PoiLongitude = Longitudine
                    If SplitDescrizione = True Then SplitDescr sDescr, cntRowArray
                    ArrayOv2PoiRec(cntRowArray).hTy2descrizione = PrimaMaiuscola(sDescr)
                    cntRowArray = cntRowArray + 1
                Next
        End If
        Set xmlDoc = Nothing
    End If
    
End Sub

Public Sub ImportASC(Filename, Optional SplitDescrizione As Boolean = False)
    Dim cntRowArray As Long
    Dim sDescr As String
    Dim Data
    Dim vField
    
    If Filename <> "" Then
        ReDim ArrayOv2PoiRec(0)
        cntRowArray = 0
        Open Filename For Input As #1
        Do Until EOF(1)
            ReDim Preserve ArrayOv2PoiRec(cntRowArray)
            Line Input #1, Data
            If Left$(Data, 1) <> ";" And Data <> "" Then
                 vField = Split(Data, ",")
                 sDescr = Replace(vField(2), Chr(34), "")
                 ArrayOv2PoiRec(cntRowArray).aTy1PoiLatitude = 0
                 ArrayOv2PoiRec(cntRowArray).bTy1PoiLongitude = 0
                 ArrayOv2PoiRec(cntRowArray).cTy1Poi3Latitude = 0
                 ArrayOv2PoiRec(cntRowArray).dTy1Poi3Longitude = 0
                 ArrayOv2PoiRec(cntRowArray).eTy1Distanza = 0
                 ArrayOv2PoiRec(cntRowArray).fTy2PoiLatitude = Replace(Trim(vField(0)), ".", ",")
                 ArrayOv2PoiRec(cntRowArray).gTy2PoiLongitude = Replace(Trim(vField(1)), ".", ",")
                 If SplitDescrizione = True Then SplitDescr sDescr, cntRowArray
                 ArrayOv2PoiRec(cntRowArray).hTy2descrizione = PrimaMaiuscola(sDescr)
                 cntRowArray = cntRowArray + 1
             End If
        Loop
        Close #1
    End If
    
End Sub

Public Sub ImportOV2(Filename, Optional SplitDescrizione As Boolean = False, Optional CalcolaDistanza As Boolean = False)
    Dim Comodo As String
    Dim descr
    Dim cntRowArray As Long
    Dim Poi As PoiRec
    Dim Poi1 As PoiRec1
    Dim Poi2 As PoiRec2
    Dim i As Long
    Dim bc As Byte
    Dim ToSkip As Long
    
    If Filename <> "" Then
        ReDim ArrayOv2PoiRec(0)
        cntRowArray = 0
        Open Filename For Binary As #1
        Do Until EOF(1)
            ReDim Preserve ArrayOv2PoiRec(cntRowArray)
            Get #1, , Poi
            Select Case Poi.type
            
            Case 1 ' Tipo 1
                Get #1, , Poi1
                ArrayOv2PoiRec(cntRowArray).aTy1PoiLatitude = Trim(Poi1.latitude1 / 100000)
                ArrayOv2PoiRec(cntRowArray).bTy1PoiLongitude = Trim(Poi1.longitude1 / 100000)
                ArrayOv2PoiRec(cntRowArray).cTy1Poi3Latitude = Trim(Poi1.latitude2 / 100000)
                ArrayOv2PoiRec(cntRowArray).dTy1Poi3Longitude = Trim(Poi1.longitude2 / 100000)
                If CalcolaDistanza = True Then ArrayOv2PoiRec(cntRowArray).eTy1Distanza = CalculateDistance(Poi1.latitude1 / 100000, Poi1.longitude1 / 100000, Poi1.latitude2 / 100000, Poi1.longitude2 / 100000)
                cntRowArray = cntRowArray + 1
            
            Case 2 ' Tipo 2
                Get #1, , Poi2
                Comodo = ""
                For i = 1 To Poi.Length - 14
                    Get #1, , bc
                    Comodo = Comodo & Chr(bc)
                Next
                Get #1, , bc
                ArrayOv2PoiRec(cntRowArray).fTy2PoiLatitude = Trim(Poi2.latitude / 100000)
                ArrayOv2PoiRec(cntRowArray).gTy2PoiLongitude = Trim(Poi2.longitude / 100000)
                
                ArrayOv2PoiRec(cntRowArray).hTy2descrizione = PrimaMaiuscola(Comodo)
                
                If SplitDescrizione = True Then
                    SplitDescr Comodo, cntRowArray
                End If
                
                cntRowArray = cntRowArray + 1

            Case 3 ' Tipo 3
                Get #1, , Poi2
                Comodo = ""
                For i = 1 To Poi.Length - 13
                    Get #1, , bc
                    Comodo = Comodo & Chr(bc)
                Next
                
                descr = Split(Comodo, Chr(0))
                Comodo = descr(0)
                
                ArrayOv2PoiRec(cntRowArray).fTy2PoiLatitude = Trim(Poi2.latitude / 100000)
                ArrayOv2PoiRec(cntRowArray).gTy2PoiLongitude = Trim(Poi2.longitude / 100000)
                
                If SplitDescrizione = True Then
                    SplitDescr Comodo, cntRowArray
                End If
                
                ArrayOv2PoiRec(cntRowArray).hTy2descrizione = PrimaMaiuscola(Comodo)
                cntRowArray = cntRowArray + 1

            Case Else
                ToSkip = Poi.Length - 5
                For i = 1 To ToSkip
                    Get #1, , bc
                Next
            
            End Select
            
            DoEvents
            
        Loop
        Close #1
    End If
    
End Sub

Public Sub SplitDescr(ByRef Comodo As String, ByRef cntRowArray As Long)
    Dim sDescr As String
    Dim sAddres As String
    Dim sCivNum As String
    Dim sCity As String
    Dim sCap As String
    Dim sProv As String
    Dim sTel As String
    Dim pos1 As Integer
    Dim pos2 As Integer
    
    sDescr = ""
    sAddres = ""
    sCivNum = ""
    sCap = ""
    sProv = ""
    sCity = ""
    sTel = ""

    ' Trovo descrizione, numero civico e indirizzo-------------------------------
    pos1 = InStr(1, Comodo, "[")
    If pos1 > 0 Then
        sDescr = Mid$(Comodo, 1, pos1 - 1)
        pos2 = InStr(pos1 + 1, Comodo, "]")
        
        If pos2 > pos1 Then
            sCivNum = Mid$(Comodo, pos1 + 1, pos2 - (pos1 + 1))
        End If
        
        pos1 = InStr(pos2 + 1, Comodo, "(")
        
        If pos1 > 0 Then
            sAddres = Mid$(Comodo, pos2 + 1, pos1 - (pos2 + 1))
        End If
    End If '----------------------------------------------------------------------
    
    ' Trovo Cap, Città e Provincia------------------------------------------------
    pos1 = InStr(1, Comodo, "(")
    If pos1 > 0 Then
        If sDescr = "" Then
        sDescr = Mid$(Comodo, 1, pos1 - 1)
        End If
        
        pos2 = InStr(pos1 + 1, Comodo, ")")
        
        If pos2 > pos1 Then
            sCity = Mid$(Comodo, pos1 + 1, pos2 - (pos1 + 1))
                If IsNumeric(Left(sCity, 5)) Then
                    sCap = Mid$(sCity, 1, 5)
                    sCity = Mid$(sCity, 7)
                End If
            pos1 = InStr(1, Comodo, ">")
            If pos1 - pos2 = 3 Then
                sProv = Mid$(Comodo, pos2 + 1, 2)
            End If
        End If
    End If '----------------------------------------------------------------------
    
    ' Trovo il numero di telefono-------------------------------------------------
    pos1 = InStr(1, Comodo, ">")
    If pos1 > 0 Then
        If sDescr = "" Then
            sDescr = Mid$(Comodo, 1, pos1 - 1)
        End If
        
        sTel = Mid$(Comodo, pos1)
        Comodo = ""
        
        ' Tolgo i caratteri che non sono numeri
        For pos1 = 1 To Len(sTel)
            If Mid$(sTel, pos1, 1) >= "0" And Mid$(sTel, pos1, 1) <= "9" Then
                Comodo = Comodo & Mid$(sTel, pos1, 1)
            End If
        Next
        If Mid$(sTel, 2, 1) = "+" Then Comodo = "+" & Comodo
        
        sTel = Comodo
    End If '----------------------------------------------------------------------
    
    If sDescr = "" Then
        sDescr = Comodo
    End If
    
    ArrayOv2PoiRec(cntRowArray).iDescrizione = PrimaMaiuscola(sDescr)
    ArrayOv2PoiRec(cntRowArray).iIndirizzo = PrimaMaiuscola(sAddres) & " " & sCivNum
    ArrayOv2PoiRec(cntRowArray).iCap = sCap
    ArrayOv2PoiRec(cntRowArray).iCitta = PrimaMaiuscola(sCity)
    ArrayOv2PoiRec(cntRowArray).iProvincia = UCase(sProv)
    ArrayOv2PoiRec(cntRowArray).iTelefono = sTel
    ArrayOv2PoiRec(cntRowArray).iCategoria = ""
    
End Sub

Public Function CalculateDistance(Lat1 As Variant, Long1 As Variant, Lat2 As Variant, Long2 As Variant) As Variant
    Dim TrueAngle1 As Variant
    Dim TrueAngle2 As Variant
    Dim Radius1 As Variant
    Dim Radius2 As Variant
    Dim XCoordinate1 As Variant
    Dim YCoordinate1 As Variant
    Dim XCoordinate2 As Variant
    Dim YCoordinate2 As Variant
    Dim X As Variant
    Dim Y As Variant
    Dim Meters As Variant
    Dim Feet As Variant

    If Lat1 = "" Then Lat1 = 0
    If Long1 = "" Then Long1 = 0
    If Lat2 = "" Then Lat2 = 0
    If Long2 = "" Then Long2 = 0
    
    TrueAngle1 = (Atn((AxisMinor ^ 2) / (AxisMajor ^ 2) * Tan(Lat1 * pi / 180))) * 180 / pi
    TrueAngle2 = (Atn((AxisMinor ^ 2) / (AxisMajor ^ 2) * Tan(Lat2 * pi / 180))) * 180 / pi

    Radius1 = (1 / ((Cos(TrueAngle1 * pi / 180)) ^ 2 / AxisMajor ^ 2 + (Sin(TrueAngle1 * pi / 180)) ^ 2 / AxisMinor ^ 2)) ^ 0.5 + Elevation
    Radius2 = (1 / ((Cos(TrueAngle2 * pi / 180)) ^ 2 / AxisMajor ^ 2 + (Sin(TrueAngle2 * pi / 180)) ^ 2 / AxisMinor ^ 2)) ^ 0.5 + Elevation

    XCoordinate1 = Radius1 * Cos(TrueAngle1 * pi / 180)
    YCoordinate1 = Radius1 * Sin(TrueAngle1 * pi / 180)
    XCoordinate2 = Radius2 * Cos(TrueAngle2 * pi / 180)
    YCoordinate2 = Radius2 * Sin(TrueAngle2 * pi / 180)

    X = ((XCoordinate1 - XCoordinate2) ^ 2 + (YCoordinate1 - YCoordinate2) ^ 2) ^ 0.5
    Y = 2 * pi * ((((XCoordinate1 + XCoordinate2) / 2)) / 360) * (Long1 - Long2)

    Meters = ((X) ^ 2 + (Y) ^ 2) ^ 0.5

    CalculateDistance = Meters
    
End Function

Public Sub VerificaCoordinate(ListView1 As ListView, ColonnaLat As Integer, ColonnaLon As Integer, MinLat As Double, MaxLat As Double, MinLon As Double, MaxLon As Double, Optional ByVal MessaggioFinale As Integer = 2, Optional ByVal EvidenziaErrori As Boolean = True)
    '
    ' MessaggioFinale:
    ' 0 = Disattivato
    ' 1 = Solo in caso di errori
    ' 2 = Attivato sempre
    '
    Dim cnt As Long
    Dim Errori As Integer
    Dim Correzioni As Integer
    Dim vTmp As Variant
    Dim Lat As Double
    Dim Lon As Double
    Dim Msg As String
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    If ListView1.ListItems.Count = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    DoEvents

    For cnt = 1 To ListView1.ListItems.Count
        vTmp = ListView1.ListItems(cnt).SubItems(ColonnaLat)
        ' Cambio il punto con la virgola
        If InStr(1, vTmp, ".", vbTextCompare) <> 0 Then
            vTmp = Replace$(vTmp, ".", ",")
            ListView1.ListItems(cnt).SubItems(ColonnaLat) = vTmp
            Correzioni = Correzioni + 1
        End If
        If (vTmp) = "" Then vTmp = 0
        Lat = vTmp
        
        vTmp = ListView1.ListItems(cnt).SubItems(ColonnaLon)
        ' Cambio il punto con la virgola
        If InStr(1, vTmp, ".", vbTextCompare) <> 0 Then
            vTmp = Replace$(vTmp, ".", ",")
            ListView1.ListItems(cnt).SubItems(ColonnaLon) = vTmp
            Correzioni = Correzioni + 1
        End If
        If vTmp = "" Then vTmp = 0
        Lon = vTmp
        
        ' Attivo il ceck della riga
        If EvidenziaErrori = True Then
            If ((Lat >= MinLat And Lat <= MaxLat) = True) And ((Lon >= MinLon And Lon <= MaxLon) = True) Then
                ListView1.ListItems.Item(cnt).Checked = False
            Else
                ListView1.ListItems.Item(cnt).Checked = True
                Errori = Errori + 1
            End If
        End If
    Next
    
    cnt = cnt - 1
    Call ControllaCheck(ListView1)
    
    If MessaggioFinale = 2 Or Errori > 0 Then
        Msg = "Verifica terminata.                                        " & vbNewLine & vbNewLine
        If Errori > 0 Or Correzioni > 0 Then
            Msg = Msg & "Correzioni: " & Correzioni & " su " & cnt * 2 & " - Errori: " & Errori & " su " & cnt
        Else
            Msg = Msg & "Nessun problema trovato."
        End If
        
        MsgBox Msg, vbInformation, App.ProductName
    End If

    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione VerificaCoordinate alla riga: " & cnt
    
End Sub
