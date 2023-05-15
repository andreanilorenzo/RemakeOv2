Attribute VB_Name = "SalvaLeggiXML"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module
    
    ' Si utilizza così:
    '
    ' Call ScriviChiaveXML("Impostazioni", "Generali", "tmpFile", "Valore", PosizioneFileXML, False)
    '
    ' Valore = LeggiChiaveXML("Impostazioni", "Generali", "tmpFile", PosizioneFileXML)
    '
    ' Call CancellaChiaveXML("Impostazioni", "Generali", "tmpFile", PosizioneFileXML)
    ' Call CancellaSezioneXML("Impostazioni", PosizioneFileXML)
    '
Dim X As New DOMDocument

Public Function LeggiXML() As Boolean
    'Dim strTmp As String
    'If Var(GestioneErrori).Valore = 0 Then On Error Resume Next
    
    'VersioneProgramma = LeggiChiaveXML("Impostazioni", "Directory", "TomTomHd", XmlFileConfig, "TomTom GO", 0)
    
    'LeggiXML = True
    
    'Exit Function
    
Errore:
    'GestErr Err, "Errore nella funzione LeggiXML."
    'LeggiXML = False
    
End Function

Private Sub CaricaXML(sPathFile As String)
    Dim strTmp As String
    
    X.async = False
    X.setProperty "SelectionLanguage", "XPath"
    X.Load (sPathFile)

    If X.xml = "" Then
        X.loadXML ("<xml></xml>")
    End If
    
End Sub

Public Function LeggiChiaveXML(ByVal sSezione As String, ByVal sSubSezione As String, ByVal sChiave As String, ByVal sPathFile As String, Optional ByVal Predefinito As String = "", Optional ByVal Indice As Integer = -1) As String
    On Error Resume Next
    
    CaricaXML sPathFile

    LeggiChiaveXML = X.selectNodes("/xml/" & sSezione & "/" & sSubSezione & "[ " & " Chiave = '" & sChiave & "' ]").Item(0).selectSingleNode("Valore").Text
    
    If LeggiChiaveXML = "" And Predefinito <> "" Then
        LeggiChiaveXML = Predefinito
        Call ScriviChiaveXML(sSezione, sSubSezione, sChiave, Predefinito, sPathFile, False, Indice)
    End If
    
End Function

Public Sub ScriviChiaveXML(ByVal sSezione As String, ByVal sSubSezione As String, ByVal sChiave As String, ByVal sValore As Variant, ByVal sPathFile As String, Optional ByVal LeggiDopo As Boolean = False, Optional ByVal Indice As Integer = -1)
    If Var(GestioneErrori).Valore = 0 Then On Error Resume Next
    Dim xl As IXMLDOMNodeList
    Dim xs As IXMLDOMNode
    Dim xn As IXMLDOMNode
    Dim xe As IXMLDOMNode
    
    CaricaXML sPathFile
    CreaSezioneXML sSezione
    
    'If Indice = -1 Then
        Set xl = X.selectNodes("/xml/" & sSezione & "/" & sSubSezione & "[ " & " Chiave = '" & sChiave & "' ]")
    'Else
    '    Set xl = X.selectNodes("/xml/" & sSezione & "/" & sSubSezione & "[ " & " Chiave = '" & sChiave & "' ]")
    'End If
    
    If xl.length = 0 Then
        Set xn = X.createNode(1, sSubSezione, "")
        Set xe = X.createTextNode(vbCrLf)
        xn.appendChild xe
        
        Set xs = X.createElement("Chiave")
        xs.Text = sChiave
        xn.appendChild xs
        Set xe = X.createTextNode(vbCrLf)
        xn.appendChild xe
       
        Set xs = X.createElement("Valore")
        xs.Text = sValore
        xn.appendChild xs
        Set xe = X.createTextNode(vbCrLf)
        xn.appendChild xe
       
        X.selectSingleNode("/xml/" & sSezione).appendChild xn
        Set xe = X.createTextNode(vbCrLf)
        X.selectSingleNode("/xml/" & sSezione).appendChild xe
           
    Else
        xl.Item(0).selectSingleNode("Valore").Text = sValore
    End If
    
    X.Save (sPathFile)
    
    If LeggiDopo = True Then Call LeggiXML
    
End Sub

Private Sub CreaSezioneXML(sSezione As String)
    Dim xl As IXMLDOMNodeList
    Dim xn As IXMLDOMNode
    Dim xe As IXMLDOMNode
    
    Set xl = X.selectNodes("/xml/*[ name() = '" & sSezione & "' ]")
        If xl.length = 0 Then
        Set xe = X.createTextNode(vbCrLf)
        X.selectSingleNode("/xml").appendChild xe
        Set xn = X.createNode(1, sSezione, "")
        Set xe = X.createTextNode(vbCrLf)
        xn.appendChild xe
        X.selectSingleNode("/xml").appendChild xn
    End If
    
End Sub

Public Function CancellaChiaveXML(sSezione As String, sSubSezione As String, sChiave As String, sPathFile As String)
    Dim xn As IXMLDOMNode

    CaricaXML sPathFile

    Set xn = X.selectNodes("/xml/" & sSezione & "/" & sSubSezione & "[ " & " Chiave = '" & sChiave & "' ]").Item(0)
    
    If IsObject(xn) Then
        xn.parentNode.removeChild xn
    End If
    
End Function

Public Sub CancellaSezioneXML(sSezione As String, sPathFile As String)
    Dim xn As IXMLDOMNode

    CaricaXML sPathFile

    Set xn = X.selectNodes("/xml/*[ name() = '" & sSezione & "' ]").Item(0)
        If IsObject(xn) Then
        xn.parentNode.removeChild xn
    End If
    
End Sub

