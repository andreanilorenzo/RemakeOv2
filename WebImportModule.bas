Attribute VB_Name = "WebImportModule"
Option Explicit

Public Function LatLonFromString(ByVal URL As String, ByRef Latitudine, ByRef Longitudine, ByVal CostrFind As String) As Boolean
    ' Prende le coordinate direttamente dalla URL restituita
    '
    ' URL = L'indirizzo web della pagina
    ' CostrFind = La stringa nella quale sono contenute le lon e lat
    '
    Dim scratch As String
    Dim pos1, pos2 As Integer
    Dim Comodo1, Comodo2 As String
    Dim arrTmp
    
    Dim strLatA As String
    Dim strLatB As String
    Dim strLonA As String
    Dim strLonB As String
    
    If URL = "" Then GoTo Esci
    
    'http://www.multimap.com/maps/?&t=l&map=44.01485,10.12846|17|4&loc=IT:44.01485:10.12846:17
    '"&map=Latitudine,Longitudine|"
    
    If InStr(1, CostrFind, "Latitudine") < InStr(1, CostrFind, "Longitudine") Then
        Comodo1 = Replace(CostrFind, "Latitudine", "°°_°_°°")
        Comodo1 = Replace(Comodo1, "Longitudine", "°°_°_°°")
        arrTmp = Split(Comodo1, "°°_°_°°")
        
        strLatA = arrTmp(0)
        strLatB = arrTmp(1)
        
        strLonA = arrTmp(UBound(arrTmp) - 1)
        strLonB = arrTmp(UBound(arrTmp))
    Else
        ' Codice ancora da scrivere.......
        '...............
        GoTo Esci
    End If

    If Var(DebugMode).Valore = 1 Then WriteLog "LatLonFromString: " & URL, "Debug"

    scratch = URL
    
    pos1 = InStr(1, scratch, strLatA)
    
    If pos1 > 0 Then
        pos2 = InStr(pos1, scratch, strLatB)
        If pos2 > pos1 Then
            Latitudine = Mid$(scratch, pos1 + Len(strLatA), pos2 - (pos1 + Len(strLatA)))
        Else
            Latitudine = "0.0"
        End If
    Else
        Latitudine = "0.0"
    End If
    
    If IsNumeric(Latitudine) = False Then
        Latitudine = "0.0"
    End If
    
    
    pos1 = InStr(1, scratch, strLonA)
    If pos1 > 0 Then
        pos2 = InStr(pos1, scratch, strLonB)
        If pos2 > pos1 Then
            Longitudine = Mid$(scratch, pos1 + Len(strLonA), pos2 - (pos1 + Len(strLonA)))
        ElseIf pos2 = pos1 Then
            Longitudine = Right$(scratch, Len(scratch) - pos1)
        Else
            Longitudine = "0.0"
        End If
    Else
        Longitudine = "0.0"
    End If
    
    If IsNumeric(Longitudine) = False Then
        Longitudine = "0.0"
    End If
    
    LatLonFromString = True
    
    Exit Function

Esci:
    LatLonFromString = False

End Function

Public Function EnumLink(WebBrowser1 As WebBrowser, Optional ByVal Filtra As String = "", Optional ByVal Contiene As String = "") As String()
    ' Note that this program sets a reference to the Microsoft HTML Object Library.
    '
    ' Filtra:contiene il testo che il link deve contenere
    ' I primi dieci caratteri contengono l'istruzione mentre il resto contiene la stringa
    '
    ' Contiene: carica solo i link che contengono questo
    '
    Dim doc As HTMLDocument
    Dim a_link As HTMLAnchorElement
    Dim txt As String
    Dim pIstruzione As String
    Dim arrTmp() As String
    Dim cnt As Integer

    'On Error Resume Next
    
    If Filtra <> "" Then
        pIstruzione = Trim(Left$(Filtra, 10))
        Filtra = Right$(Filtra, Len(Filtra) - 10)
    End If
    
    ReDim arrTmp(1)
    
    Set doc = WebBrowser1.Document
    
    ' Scorro tutti i link
    For Each a_link In doc.links
        If Filtra = "" Then
            If Contiene = "" Or InStr(1, a_link.href, Contiene, vbTextCompare) > 0 Then
                cnt = cnt + 1
                ReDim Preserve arrTmp(cnt)
                arrTmp(cnt - 1) = a_link.href
            End If
            
        ElseIf pIstruzione = "LEFT" Then
            If Left$(a_link.href, Len(Filtra)) = Filtra Then
                If Contiene = "" Or InStr(1, a_link.href, Contiene, vbTextCompare) > 0 Then
                    cnt = cnt + 1
                    ReDim Preserve arrTmp(cnt)
                    arrTmp(cnt - 1) = a_link.href
                End If
            End If
            
        End If
        
    Next a_link
    
    EnumLink = arrTmp
    
    Set doc = Nothing

End Function

Public Function GetValoreFromURL(ByVal URL As String, Chiave As String, Optional ByVal SplitURL As String = "&", Optional ByVal SplitValore As String = "=") As String
    ' Restituisce il valore della chiave
    Dim i As Integer
    Dim varTmp
    Dim varRet
    
    On Error GoTo Esci
    
    If URL = "" Then GoTo Esci

    varTmp = Split(URL, SplitURL)
    
    For i = 0 To UBound(varTmp)
        If varTmp(i) <> "" And Left$(varTmp(i), Len(Chiave)) = Chiave Then
            varRet = Split(varTmp(i), SplitValore)
            GetValoreFromURL = varRet(1)
            Exit For
        End If
    Next
    
    Exit Function

Esci:
    GetValoreFromURL = ""

End Function
