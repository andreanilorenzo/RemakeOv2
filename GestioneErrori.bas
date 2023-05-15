Attribute VB_Name = "GestioneErroriModule"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

Public Function GestErr(ByVal Errore As ErrObject, Optional ByVal Messaggio As String = "", Optional ByVal SoloTesto As Boolean = False, Optional ByVal SuUnaRiga As Boolean = False) As String
    
    Screen.MousePointer = vbDefault

    Messaggio = Trim$(Messaggio)
    If Messaggio <> "" Then Messaggio = Messaggio & "  " & vbNewLine
    Messaggio = Messaggio & "Err: " & CStr(Errore.Number) & " - Desc: " & Errore.Description & " - " & Errore.Source
    If SuUnaRiga = True Then Messaggio = Replace(Messaggio, vbNewLine, "")
    
    Screen.MousePointer = vbDefault
    
    If SoloTesto = False Then
        MsgBox Messaggio, vbCritical, App.ProductName
    End If
    
    GestErr = Messaggio

End Function

