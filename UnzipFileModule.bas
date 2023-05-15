Attribute VB_Name = "UnzipFileModule"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

Private Type tyZipFile
    SizeFile As Long
    SizeFileCompresso As Long
    FileCriptato As Boolean
    NomeFile As String
    Data As String
    Cartella As String
End Type

Private arrZipFile() As tyZipFile

Private m_cUnzip As cUnzip
Private m_cExtractToMRU As cMRU
Private m_cZipMRU As cMRU
Private m_sBaseKey As String

Public Sub AvviaOperazioniUnZip()
    Dim ListaFile() As String: ReDim ListaFile(0)
    Dim cnt As Long
    
    On Error GoTo Errore

    ' Set up unzipping object
    Set m_cUnzip = New cUnzip
    ' Set up Extract To MRU:
    Set m_cExtractToMRU = New cMRU
    ' Set up Zip FIles MRU:
    Set m_cZipMRU = New cMRU

    ListaFile = FileList(Var(tmpDownFile).Valore & "\*.zip")
    
    For cnt = 0 To UBound(ListaFile)
        Call ApriZipFile(Var(tmpDownFile).Valore & "\" & ListaFile(cnt), "ov2")
        Call ApriZipFile(Var(tmpDownFile).Valore & "\" & ListaFile(cnt), "bmp")
    Next
    
Exit_Label:
    Exit Sub

Errore:
    MsgBox "Errore nella funzione AvviaOperazioniUnZip. Err: " & CStr(Err.Number) & " - Desc: " & Err.Description & " - " & Err.Source, vbCritical, App.ProductName
    
End Sub

Private Function ApriZipFile(ByVal sFile As String, Optional Estensione As String) As Boolean
'Public Function pOpen(ByVal sFIle As String) As Boolean
    Dim I As Long
    Dim EstTrovata As Boolean
    
    EstTrovata = False

    ' Cancello l'array
    ReDim arrZipFile(0)
    
    ' Imposto la directory
    m_cUnzip.ZipFile = sFile
    m_cUnzip.Directory
    
    ' Carico i dati nell'array
    For I = 1 To m_cUnzip.FileCount
        ReDim Preserve arrZipFile(I - 1)
        arrZipFile(I - 1).FileCriptato = m_cUnzip.FileEncrypted(I)
        arrZipFile(I - 1).NomeFile = m_cUnzip.Filename(I)
        arrZipFile(I - 1).SizeFile = m_cUnzip.FileSize(I)
        arrZipFile(I - 1).Data = Format$(m_cUnzip.FileDate(I), "short date") & " " & Format$(m_cUnzip.FileDate(I), "short time")
        arrZipFile(I - 1).SizeFileCompresso = m_cUnzip.FilePackedSize(I)
        arrZipFile(I - 1).Cartella = m_cUnzip.FileDirectory(I)
        If Estensione = Right$(m_cUnzip.Filename(I), 3) Then EstTrovata = True
    Next
    
    ' Estraggo i file.....
    If EstTrovata = True Then
        Call EstraiZipFile(Var(PoiScaricati).Valore, Estensione)
    End If

End Function

Private Sub EstraiZipFile(Optional ByVal sFolder As String = "", Optional ByVal Estensione As String = "")
    Dim cnt As Long
    Dim bSel As Boolean
    Dim EstensioneFileCor As String
    Dim NumFileDaDecomprimere As Integer
    
    If sFolder = "" Then sFolder = App.Path
    Estensione = LCase$(Estensione)
    NumFileDaDecomprimere = 0

    For cnt = 0 To UBound(arrZipFile)
        If Estensione = "" Then
            ' Seleziono tutti i file contenuti nel file .zip
            m_cUnzip.FileSelected(cnt + 1) = True
            NumFileDaDecomprimere = NumFileDaDecomprimere + 1
            bSel = True
            
        Else
            EstensioneFileCor = LCase$(Right$(arrZipFile(cnt).NomeFile, 3))
            
            ' Seleziono solo i file con l'estensione cercata
            If Estensione = EstensioneFileCor Then
                m_cUnzip.FileSelected(cnt + 1) = True
                NumFileDaDecomprimere = NumFileDaDecomprimere + 1
                
                If Estensione = "bmp" Then
                    If Var(ScaricaBMP).Valore = False And (FileExists(Var(PoiScaricati).Valore & "/" & ElaboraNomeFile(arrZipFile(cnt).NomeFile, False)) = True) Then
                        ' Se il file che si vuole scaricare è .bmp e questo esiste già, non lo scarico
                        m_cUnzip.FileSelected(cnt + 1) = False
                        NumFileDaDecomprimere = NumFileDaDecomprimere - 1
                    End If
                End If
                
            Else
                m_cUnzip.FileSelected(cnt + 1) = False
            End If
            
        End If
    Next
   
    ' Se non c'è la cartella oppure non ci sono file da decomprimere non avvio l'operazione
    If (sFolder <> "") And NumFileDaDecomprimere >= 1 Then
       m_cExtractToMRU.Add sFolder
       m_cUnzip.UnzipFolder = sFolder
       ' Estraggo solo se il file nel file .zip è più nuovo di quello esistente
       m_cUnzip.ExtractOnlyNewer = False
       ' Sovrascrivo i file esistenti
       m_cUnzip.OverwriteExisting = True
       ' Estraggo i file
       m_cUnzip.Unzip
    End If
   
End Sub
