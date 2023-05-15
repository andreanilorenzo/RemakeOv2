Attribute VB_Name = "FileFolderModule"
Option Explicit

'-------------------------------------------------------------------------------------
' Per la funzione ShellDelete
Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40
Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
'-------------------------------------------------------------------------------------

Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Sub NascondiFile(Estensione, Cartella)
    ' Imposta l'attributo dei file in Cartella su nascosto
    Dim sFileSpec As String
    Dim sFilename As String
    Dim varAttributes As VbFileAttribute
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    If Var(BMPnascoste).Valore = 0 Then
        varAttributes = vbNormal + vbArchive
    ElseIf Var(BMPnascoste).Valore = 1 Then
        varAttributes = vbHidden + vbArchive
    End If
    
    sFileSpec = Cartella & "\" & Estensione
    sFilename = Dir(sFileSpec)

    Do ' scorro tutti i file
        If sFilename = "" Then Exit Do
        SetAttr Cartella & "\" & sFilename, varAttributes
        sFilename = Dir
    Loop

    Exit Sub

Errore:
    GestErr Err, "Errore nella funzione NascondiFile."

End Sub

Public Function ElaboraNomeFileInCartella(ByVal Cartella As String, Optional ByVal Estensione As String = "*.*", Optional Ripristina As Boolean = False) As Boolean
        
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
        
    If Var(xElaboraNomeFile).Valore = False Then
        Exit Function
    End If
    
    Dim ListaFile() As String: ReDim ListaFile(0)
    Dim cnt As Long
    
    ' Carico l'arrai con tutti i file trovati nella cartella
    ListaFile = FileList(Cartella & "\" & Estensione)
    
    If ListaFile(0) <> "" Then
        For cnt = 0 To UBound(ListaFile)
            Call RinominaFile(Cartella & "\" & ListaFile(cnt), Cartella & "\" & ElaboraNomeFile(ListaFile(cnt), Ripristina))
        Next
        DoEvents
    End If
    
    Exit Function
    
Errore:
    GestErr Err, "Errore nella funzione ElaboraNomeFileInCartella." & vbNewLine & "(" & Cartella & "-" & Estensione & "-" & Ripristina & ")"
    ElaboraNomeFileInCartella = False
    Exit Function
    
End Function

Public Function ElaboraNomeFile(ByVal StringaTesto As String, Optional ByVal Ripristina As Boolean = False) As String
    Dim cnt As Integer
    Dim cntCar As Integer
    Dim car As String
    Dim strResult As String
    Dim strTmp As String
    Dim Splitted() As String
    Dim Caratteri() As String
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    If Var(DebugMode).Valore = 1 Then
        WriteLog "ElaboraNomeFile: " & StringaTesto & DebSep & Ripristina & DebSep & "xElaboraNomeFile: " & Var(xElaboraNomeFile).Valore, "Debug"
    End If
    
    If Var(xElaboraNomeFile).Valore = False Then
        strResult = StringaTesto
        GoTo Esci
    End If

    ReDim Caratteri(1, 1)
    
    If Ripristina = False Then
        Caratteri(0, 0) = " "
        Caratteri(0, 1) = "_"
    Else
        Caratteri(0, 0) = "_"
        Caratteri(0, 1) = " "
    End If
    
    StringaTesto = Replace$(StringaTesto, Caratteri(0, 0), Caratteri(0, 1))
    
    strResult = ""
    cntCar = 1
    For cnt = 1 To Len(StringaTesto) Step 1
        car = Mid$(StringaTesto, cntCar, 1)
        
        Select Case car
            
            Case Is = "_"
                strResult = strResult & "_" & UCase$(Mid$(StringaTesto, cntCar + 1, 1))
                cntCar = cntCar + 2
            
            Case Is = "-"
                strResult = strResult & "-" & UCase$(Mid$(StringaTesto, cntCar + 1, 1))
                cntCar = cntCar + 2
            
            Case Is = " "
                ' Se è un carattere solo lo lascio minuscolo
                If Mid$(StringaTesto, cntCar + 2, 1) = " " Then
                    strResult = strResult & " " & LCase$(Mid$(StringaTesto, cntCar + 1, 1))
                Else
                    strResult = strResult & " " & UCase$(Mid$(StringaTesto, cntCar + 1, 1))
                End If
                cntCar = cntCar + 2
            
            Case Else
                strResult = strResult & car
                cntCar = cntCar + 1
                
        End Select
    Next

Esci:
    ElaboraNomeFile = strResult
    Exit Function
    
Errore:
    GestErr Err, "Errore nella funzione ElaboraNomeFile." & vbNewLine & "(" & StringaTesto & "-" & Ripristina & ")"
    ElaboraNomeFile = StringaTesto
    Exit Function

End Function

Public Function ElaboraNome(StringaTesto As String) As String
    ' Per adesso non è utilizzata, ma potrebbe servire.........
    Dim cnt As Integer
    Dim LenStr As Long
    Dim car As String
    Dim strResult As String
    Dim strTmp As String
    Dim Splitted() As String
    
    LenStr = Len(StringaTesto)
    
    For cnt = 1 To LenStr Step 1
        car = Mid$(StringaTesto, cnt, 1)
        Select Case car
            Case Is = "-", "_"
                strResult = strResult & " "
            Case Else
                strResult = strResult & car
        End Select
    Next
    
    ' Divido la stringa in più stringhe con delimitatore lo spazio
    Splitted = Split(strResult)
    strResult = ""
    
    For cnt = 0 To UBound(Splitted)
        ' La prima lettera maiuscola
        strTmp = PrimaMaiuscola(Trim$(Splitted(cnt)))
        If strTmp <> "" Then strResult = strResult & " " & strTmp
    Next
    
    ' Tolgo gli eventuali spazi di troppo
    strResult = Trim$(strResult)
    
    ' Sostituisco gli spazi nel nome del file con il carattere "_"
    strResult = Replace(strResult, " ", "_", 1, -1, vbTextCompare)
    
    ElaboraNome = strResult
    
End Function

Public Function RinominaFile(ByVal Origine As String, ByVal Destinazione As String, Optional ByVal EliminaEsistente As Boolean = True) As Boolean
    
    If LCase$(Origine) = LCase$(Destinazione) Then
        RinominaFile = False
        Exit Function
    End If
    
    If EliminaEsistente = True And FileExists(Destinazione) = True Then
        Kill Destinazione
    End If
    
    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFile Origine, Destinazione
    Set fso = Nothing
    RinominaFile = True
    
End Function

Public Function GetLinkInfo(ByVal strLinkName As String) As String
   ' Add a reference (VB Menu -> Project -> References)
   ' to the "Windows Scripting Host Object Model".
   Dim wshShell As New IWshShell_Class
   Dim wshLink As IWshShortcut_Class

   If strLinkName = "" Then
    GetLinkInfo = ""
    Exit Function
   End If
   
   Set wshLink = wshShell.CreateShortcut(strLinkName)
   'Debug.Print "Arguments: " & wshLink.Arguments
   'Debug.Print "Description: " & wshLink.Description
   'GetLinkInfo = wshLink.FullName
   'Debug.Print "HotKey: "; wshLink.HotKey
   'Debug.Print "IconLocation: " & wshLink.IconLocation
   GetLinkInfo = wshLink.TargetPath
   'Debug.Print "WindowStyle: " & wshLink.WindowStyle
   'Debug.Print "WorkingDirectory: " & wshLink.WorkingDirectory

End Function

Public Function FileList(ByVal mask As String) As String()
    ' This function takes a directory name or a file/directory mask
    ' string (e.g., "C:\*.txt") and returns a string array containing
    ' all the files in the directory or meeting the criteria.
    Dim sWkg As String
    Dim sAns() As String
    Dim lCtr As Long
    
    ReDim sAns(0) As String
    sWkg = Dir(mask, vbNormal)
    
    Do While Len(sWkg)
    
        If sAns(0) = "" Then
            sAns(0) = sWkg
        Else
            lCtr = UBound(sAns) + 1
            ReDim Preserve sAns(lCtr) As String
            sAns(lCtr) = sWkg
        End If
        sWkg = Dir
    Loop
    
    FileList = sAns

End Function

Public Function GetFileInFolder(ByRef ArrayFile() As String, ByVal FolderSpec As String, sEstensioni As String) As Boolean
    ' FolderSpec = la cartella da esaminare
    ' sEstensioni = Stringa con le estensioni
    '
    ' Esempio:
    ' xxx = GetFileInFolder("C:\", ".exe|.doc|.txt"
    '
    Dim OFS As New FileSystemObject
    Dim oFolder As folder
    Dim oFile As File
    Dim oFileCollection
    Dim tempBoolean
    Dim SubFolders As folder
    Dim strFile As String
    Dim arrEstens As Variant
    Dim cnt As Integer
    Dim arrTmp As Variant
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    arrEstens = Split(LCase(sEstensioni), "|", , vbTextCompare)
    
    If OFS.FolderExists(FolderSpec) Then
        Set oFolder = OFS.GetFolder(FolderSpec)
        
        On Error Resume Next
        ' Scorro tutti i file nella cartella
        For Each oFile In oFolder.Files
            ' Scorro tutto l'array con le estensioni
            For cnt = 0 To UBound(arrEstens)
                If LCase(EstensioneFromFile(oFile.Name)) = LCase(arrEstens(cnt)) Or arrEstens(cnt) = ".*" Then
                    strFile = oFile.Name & vbTab & strFile
                    Exit For
                End If
            Next
        Next
    End If
    
    If Replace$(strFile, vbTab, "") <> "" Then
        arrTmp = Split(strFile, vbTab, , vbTextCompare)
        ' Tolgo l'ultimo valore dell'array che è vuoto
        ReDim Preserve arrTmp(UBound(arrTmp) - 1)
        ' Assegno il risultato alla funzione
        ArrayFile = arrTmp
        GetFileInFolder = True
    Else
        ArrayFile = Split("")
        GetFileInFolder = False
    End If
    
    Exit Function
    
Errore:
    ArrayFile = Split("")
    GetFileInFolder = False
    GestErr Err, "Errore nella funzione GetFileInFolder."
    
End Function

Public Function DeleteAllFiles(ByVal FolderSpec As String, Optional ByVal DeleteSubFolder As Boolean = False) As Boolean
    ' Requires a reference to the Scripting Runtime.
    '
    ' Deletes all files in folder specified by parameter FolderSpec.
    ' Does not delete subfolders or files within subfolders
    '
    ' Returns True if sucessful, false otherwise
    '
    'Requires a reference the Microsoft Scripting Runtime
    '
    'EXAMPLE: DeleteAllFiles "C:\Test"
    
    Dim OFS As New FileSystemObject
    Dim oFolder As folder
    Dim oFile As File
    Dim oFileCollection
    Dim tempBoolean
    Dim SubFolders As folder
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    If OFS.FolderExists(FolderSpec) Then
        Set oFolder = OFS.GetFolder(FolderSpec)
        On Error Resume Next
        
        For Each oFile In oFolder.Files
            oFile.Delete True 'setting force to true deletes read-only file
        Next
        
        If DeleteSubFolder = True Then
            ' cancello tutte le sottocartelle
            Set oFolder = OFS.GetFolder(FolderSpec)
            Set oFileCollection = oFolder.SubFolders
            For Each SubFolders In oFileCollection
                tempBoolean = DeleteAllFiles(FolderSpec)
            Next
        End If

        DeleteAllFiles = oFolder.Files.count = 0
    End If
    
    Exit Function
    
Errore:
    GestErr Err, "Errore nella funzione DeleteAllFiles."
    
End Function

Public Function FileExists(ByVal sPath As String) As Boolean

  'Determines if a file exists. This function
  'tests the validity of the file and path. It
  'works only on the local file system or on a
  'remote drive that has been mounted to a drive
  'letter.
  '
  'It will return False for remote file paths
  'that begin with the UNC names \\server
  'or \\server\share. It will also return False
  'if a mounted remote drive is out of service.
  '
  'Requires Version 4.71 and later of Shlwapi.dll
  
   FileExists = CBool(PathFileExists(sPath) = 1)
   
End Function

Public Function FileNameFromPath(ByVal path As String, Optional Controlla As Boolean = False) As String
' Trova il nome di un file passanbogli in percorso
    Dim strTmp As String
    
    ' Se il file non esiste esco dalla funzione
    If Controlla = True Then
        If FileExists(path) = False Then
            FileNameFromPath = ""
            Exit Function
        End If
    End If
    
    If path = "" Then
        FileNameFromPath = ""
        Exit Function
    End If
    
    strTmp = StrReverse(Split(StrReverse(path), "\")(0))
    
    If Len(strTmp) = Len(path) Then
        strTmp = StrReverse(Split(StrReverse(path), "/")(0))
        FileNameFromPath = strTmp
    Else
        FileNameFromPath = strTmp
    End If

End Function

Public Function EstensioneFromFile(ByVal sFile As String) As String
    ' sFile = il nome o il percorso completo del file
    Dim Estensione As String
    
    EstensioneFromFile = LCase$(Right$(sFile, 4))
    
    If Left$(EstensioneFromFile, 1) = "." Then
        Exit Function
    Else
        EstensioneFromFile = ""
    End If
    
End Function

Public Function DirectoryFromFile(ByVal fullpath As String, Optional ByVal NoBarraFinale As Boolean = True) As String
    ' INPUT: File FullPath
    ' RETURNS:  Directory only
    '
    ' EXAMPLE:
    ' DirectoryFromFile("C:\Program Files\My Program\MyData.txt")
    ' Returns "C:\Program Files\My Program"
    ' Se l'argomento NoBarraFinale non è specificato la barra \ non viene inserita alla fine della stringa

    Dim sAns As String
    
    sAns = Trim(fullpath)
    
    If Len(sAns) = 0 Then Exit Function
    
    If InStr(sAns, "\") = 0 Then Exit Function

    If Right(sAns, 1) = "\" Then
        DirectoryFromFile = sAns
        Exit Function
    End If

    Do Until Right(sAns, 1) = "\"
        sAns = Left(sAns, Len(sAns) - 1)
    Loop
    
    If NoBarraFinale = False Then
        If Right(sAns, 1) <> "\" Then sAns = sAns & "\"
    Else
        If Right(sAns, 1) = "\" Then sAns = Left(sAns, Len(sAns) - 1)
    End If
    
    DirectoryFromFile = sAns

End Function

Public Function CreateFolder(ByVal destDir As String, Optional ByVal CancellaCartella As Boolean = False) As Boolean
    ' To use this function, simply pass the FULL PATH of the folder you wish to create.
    ' Make sure that destDir always end with "\" (without quotes).
    ' The function returns TRUE if operation was successful. Otherwise, it returns FALSE.
    ' Example:
    ' If CreateFolder("C:\Folder1\Folder2\Folder3\", true) Then
    '   MsgBox "Folder Creation successful!"
    ' Else
    '   MsgBox "Folder Creation failed!"
    ' End If

    Dim i As Long
    Dim prevDir As String
    
    If CancellaCartella = True Then Call KillFolder(destDir)
    
    On Error Resume Next
     
    For i = Len(destDir) To 1 Step -1
        If Mid(destDir, i, 1) = "\" Then
            prevDir = Left(destDir, i - 1)
            Exit For
        End If
    Next i
     
    If prevDir = "" Then CreateFolder = False: Exit Function
    If Not Len(Dir(prevDir & "\", vbDirectory)) > 0 Then
        If Not CreateFolder(prevDir) Then CreateFolder = False: Exit Function
    End If
     
    On Error GoTo errDirMake
    If FileExists(destDir) = False Then MkDir destDir
    CreateFolder = True
    Exit Function
     
errDirMake:
    CreateFolder = False
    
End Function

Public Function KillFolder(ByVal fullpath As String) As Boolean
    ' DELETES A FOLDER, INCLUDING ALL SUB-DIRECTORIES, FILES, REGARDLESS OF THEIR ATTRIBUTES
    ' PARAMETER: FullPath = FullPath of Folder to Delete
    ' RETURNS:   True is successful, false otherwise
    ' REQUIRES:  Reference to Microsoft Scripting Runtime
    '            Caution in use for obvious reasons
    ' EXAMPLE:   'KillFolder("D:\MyOldFiles")

    On Error Resume Next
    Dim oFso As New Scripting.FileSystemObject

    ' Deletefolder method does not like the "\" at end of fullpath
    If Right(fullpath, 1) = "\" Then fullpath = Left(fullpath, Len(fullpath) - 1)

    If oFso.FolderExists(fullpath) Then
        'Setting the 2nd parameter to true
        'forces deletion of read-only files
        oFso.DeleteFolder fullpath, True
        KillFolder = Err.Number = 0 And oFso.FolderExists(fullpath) = False
    End If

End Function

Public Function ShellDelete(ParamArray vntFileName() As Variant)
    ' Cancella un file Avvisando l'utente tramite il messaggio di windows
    Dim i As Integer
    Dim sFileNames As String
    Dim SHFileOp As SHFILEOPSTRUCT
    
    For i = LBound(vntFileName) To UBound(vntFileName)
            sFileNames = sFileNames & vntFileName(i) & vbNullChar
    Next
    sFileNames = sFileNames & vbNullChar
    
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileNames
        .fFlags = FOF_ALLOWUNDO
    End With
    
    ShellDelete = SHFileOperation(SHFileOp)

End Function


Public Sub CopiaFile(ByVal SourceFile As String, ByVal DestinationFile As String, Optional ByVal CancellaPrimaEsistente As Boolean = False)
    On Error GoTo Eror
    
    ' Se.......... cancello il file
    If CancellaPrimaEsistente = True And FileExists(DestinationFile) = True Then
        Kill (DestinationFile)
        DoEvents
    End If
    
    FileCopy SourceFile, DestinationFile
    
    Exit Sub
    
Eror:
    If Err.Number = 75 Then
        MsgBox "Please select a file to copy", vbCritical, App.ProductName
    End If
    
    If Err.Number = 70 Then
        MsgBox "To copy just one file you must uncheck the CheckBox", vbCritical, App.ProductName
    End If
    
    If Err.Number = 61 Then
        If MsgBox("Il disco è pieno. Se vuoi continuare libera spazio e premi OK, altrimenti premi Cancella", vbOKCancel + vbCritical, App.ProductName) = vbOK Then
            Exit Sub
        Else
            Exit Sub
        End If
    End If

End Sub

Public Function CreaFile(ByVal PatchFileName As String, Optional ByVal Testo As String = "") As Boolean
    Dim File1 As Integer

    On Error GoTo Errore

    File1 = FreeFile
    
    Open PatchFileName$ For Output As #File1
        Print #File1, Testo
    Close #File1

    CreaFile = True
    Exit Function
    
Errore:
    CreaFile = False

End Function

Public Function LeggiFile(ByVal PatchFileName As String, Optional ByVal MaxCaratteri As Integer = 0) As String
    Dim NumFile As Integer
    
    If PatchFileName = "" Or FileExists(PatchFileName) = False Then Exit Function
    
    NumFile = FreeFile()
    Open PatchFileName For Input As #NumFile
    
    If MaxCaratteri <= 0 Then
        LeggiFile = Input(LOF(NumFile), NumFile)
    Else
        'Legge i primi caratteri del file
        LeggiFile = Input(MaxCaratteri, #NumFile)
    End If
    
    Close #NumFile
    
End Function

Public Function ContaLineeFile(ByVal file_name As String) As Long
    ' Legge un file linea per linea e conta il numero di linee
    Dim fNum As Integer
    Dim lines As Long
    Dim one_line As String

    If Dir$(file_name) = "" Then
        'MsgBox "Il file non esiste. L'operazione sarà annullata"
        ContaLineeFile = 0
        Exit Function
    End If

    fNum = FreeFile
    Open file_name For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, one_line
        lines = lines + 1
    Loop
    Close fNum

    ContaLineeFile = lines
    
End Function

Public Function FileExistsOld(ByVal TheFileName As String) As Boolean
    ' Vecchia funzione non più usata

    ' Guarda se il file indicato nella stringa passata esiste oppure no
    ' e restituisce True o False
    On Error GoTo FileExists_Err
    
    If Len(TheFileName$) = 0 Then
        FileExistsOld = False
        Exit Function
    End If
    
    If Len(Dir$(TheFileName$)) Then
        FileExistsOld = True
    Else
        FileExistsOld = False
    End If
    
FileExists_Err:

End Function

Public Function CartellaDesktop(ByVal NomeCartella As String, Optional ByVal CancellaPrima As Boolean = True) As String
    Dim DesktopPath As String
    Dim obj As Object

    Set obj = CreateObject("WScript.Shell")
    'DesktopPath = obj.SpecialFolders("AllUsersDesktop")
    DesktopPath = obj.SpecialFolders("Desktop")

    If CancellaPrima = True Then
        KillFolder DesktopPath & "\" & NomeCartella
        DoEvents
    End If

    ' Creo la cartella nel Desktop
    If CreateFolder(DesktopPath & "\" & NomeCartella & "\", True) = False Then
        MsgBox "Errore nella creazione della cartella nel Desktop", vbCritical
        CartellaDesktop = ""
        Exit Function
    End If

    CartellaDesktop = DesktopPath & "\" & NomeCartella
    
End Function

