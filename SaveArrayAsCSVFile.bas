Attribute VB_Name = "SaveArrayAsCSVFile"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

Public Function SaveAsCSV(myArray() As String, ByVal sFilename As String, Optional ByVal PrimaRiga As String = "", Optional ByVal sDelimiter As String = ";", Optional ByVal Messaggio As Boolean = False, Optional ByVal Estensione As String = ".csv", Optional ByVal NelMessaggioRigheInMeno As Integer = 0) As Long
    ' SaveAsCSV saves an array as csv file. Choosing a delimiter different as a comma, is optional.
    Dim cntRecord As Long
    Dim n As Long
    Dim M As Long
    Dim sCSV As String 'csv string to print
    
    On Error GoTo Errore
    
    ' Check extension and correct if needed
    If InStr(sFilename, Estensione) = 0 Then
        sFilename = sFilename & Estensione
    Else
        While (Len(sFilename) - InStr(sFilename, Estensione)) > 3
            sFilename = Left(sFilename, Len(sFilename) - 1)
        Wend
    End If
    
    If MultiDimensional(myArray) = False Then '1 dimension
        'save the file
        Open sFilename For Output As #7
        If PrimaRiga <> "" Then
            Print #7, PrimaRiga
            cntRecord = cntRecord + 1
        End If
        For n = 0 To UBound(myArray, 1)
            Print #7, myArray(n, 0)
            cntRecord = cntRecord + 1
        Next n
        
    Else 'more dimensional
        'save the file
        Open sFilename For Output As #7
        If PrimaRiga <> "" Then
            Print #7, PrimaRiga
            cntRecord = cntRecord + 1
        End If
        For n = 0 To UBound(myArray, 1)
            sCSV = ""
            
            For M = 0 To UBound(myArray, 2)
                ' sostituisco il carattere sDelimiter con .
                sCSV = sCSV & Replace(myArray(n, M), sDelimiter, ".", , , vbTextCompare) & sDelimiter
            Next M
            
            sCSV = Left(sCSV, Len(sCSV) - 1) 'Rimuovo l'ultimo delimitatore
            Print #7, sCSV
            cntRecord = cntRecord + 1
        Next n
    End If
    
    Close #7
    
    If PrimaRiga <> "" Then cntRecord = cntRecord - 1
    cntRecord = cntRecord - NelMessaggioRigheInMeno
    If Messaggio <> False Then MsgBox " " & cntRecord & " record scritti "
    SaveAsCSV = cntRecord
    Exit Function
    
Errore:
      Close #7
      SaveAsCSV = -1
      MsgBox "Errore nell'esportazione del file!"

End Function

Public Function ImportCSVinArray(ByVal sFilename As String, Optional ByVal sDelimiter As String = ";", Optional ByVal Messaggio As Boolean = False, Optional ByVal SaltaRighe As Integer = 0, Optional ByVal AggiungiRigaIntestazione As Boolean = False) As String()
    ' Function ImportCSVinArray imports a csv file into an array. Choosing a delimiter different as a comma, is optional.
    Dim myArray() As String
    Dim sSplit() As String
    Dim sLine As String
    Dim lRows As Long
    Dim lColumns As Long
    Dim lCounter As Long
    Dim lTmp As Long
    Dim cnt As Long
    
    On Error GoTo ErrHandler_ImportCSVinArray
    
    If Dir(sFilename) <> "" Then
      ' Calcolo quante righe e colonne servono -----------------------------------------------
      lRows = 0
      lColumns = 0
      Open sFilename For Input As #7
      
      While Not (EOF(7))
        If SaltaRighe <> 0 And SaltaRighe > cnt Then
            Line Input #7, sLine
            cnt = cnt + 1
        Else
            Line Input #7, sLine
            If Len(sLine) > 0 Then
              sSplit() = Split(sLine, sDelimiter)
              If UBound(sSplit) > lColumns Then
                lTmp = UBound(sSplit)
                If lTmp > lColumns Then lColumns = UBound(sSplit)
              End If
              lRows = lRows + 1
            End If
        End If
      Wend
      If AggiungiRigaIntestazione = True Then lRows = lRows + 1
      
      Close #7
      '----------------------------------------------------------------------------------------
      
      cnt = 0
      
      'fill array -----------------------------------------------------------------------------
      
      If lColumns = 1 Then 'no csv file!
        ReDim myArray(lRows - 1)
        
        Open sFilename For Input As #7
        lRows = 0
        While Not (EOF(7))
            If SaltaRighe <> 0 And SaltaRighe > cnt Then
                Line Input #7, sLine
                cnt = cnt + 1
            Else
                Line Input #7, sLine
                If Len(sLine) > 0 Then
                    myArray(lRows) = sLine
                    lRows = lRows + 1
                End If
            End If
        Wend
        Close #7
        
        ElseIf lColumns > 1 Then 'multidimensional csv file
            ReDim myArray(lRows - 1, lColumns)
            lRows = 0
            
            If AggiungiRigaIntestazione = True Then
                ' Salto la riga per l'intestazione
                For lCounter = 0 To lColumns
                    myArray(0, lCounter) = "Col. " & lCounter + 1
                Next
                cnt = 0
                lRows = lRows + 1
            End If
            
            Open sFilename For Input As #7

           While Not (EOF(7))
                If SaltaRighe <> 0 And SaltaRighe > cnt Then
                    Line Input #7, sLine
                    cnt = cnt + 1
                Else
                    Line Input #7, sLine
                    If Len(sLine) > 0 Then
                      sSplit() = Split(sLine, sDelimiter)
                      
                      For lCounter = 0 To UBound(sSplit)
                        myArray(lRows, lCounter) = sSplit(lCounter)
                      Next lCounter
    
                      lRows = lRows + 1
                    End If
                End If
            Wend
            Close #7
            
        Else ' Nessun dato importabile
            arrListViewVuoto = True
            MsgBox "Nel file non sono stati trovati campi importabili. Prova a cambiare il carattere di separazione dei campi." & vbNewLine & "Il caratere che stai utilizzando ora é " & sDelimiter, vbInformation, App.ProductName
            GoTo Esci
            
      End If
      
      'return function
      ImportCSVinArray = myArray
      arrListViewVuoto = False
    Else
        arrListViewVuoto = True
    End If

Esci:
    Exit Function
    
ErrHandler_ImportCSVinArray:
    arrListViewVuoto = True
    Close #7

End Function

Public Function MultiDimensional(CheckArray() As String) As Boolean

    On Error GoTo ErrHandler_MultiDimensional
    
    If UBound(CheckArray, 2) > 0 Then
      MultiDimensional = True 'more than 1 dimension
    End If
    
    Exit Function
    
ErrHandler_MultiDimensional:
      MultiDimensional = False '1 dimension
  
End Function

