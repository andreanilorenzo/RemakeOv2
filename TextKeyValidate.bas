Attribute VB_Name = "TextKeyValidate"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

Public Enum TypeData
    Numeric = 1
    AlphaNumeric = 2
    Alpha = 3
    TimeDate = 4
End Enum

Public Sub KeyInvioTab(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
        Exit Sub
    End If
   
End Sub

Public Function KeyValidate(KeyAscii As Integer, Optional enmDataType As TypeData = Numeric, Optional TextBox As TextBox, Optional NumeroCaratteriMax As Long = 0, Optional ConvertToCapital As Boolean = False, Optional CancellaDuplicati As String = "", Optional KeyTab As Boolean = False) As Integer
   Dim PosCursore As Long
   Dim strTmp As String
   
    If KeyTab = True Then
        KeyValidate = 0
        Call KeyInvioTab(KeyAscii)
    End If
   
    If ConvertToCapital = True And (enmDataType = Alpha Or enmDataType = AlphaNumeric) Then
       If KeyAscii >= 97 And KeyAscii <= 122 Then
          KeyValidate = KeyAscii - 32
       End If
    End If
   
    Select Case enmDataType
        Case Is = Numeric
           If KeyAscii < 48 Or KeyAscii > 57 Then
              If Not KeyAscii = 8 Then
                 KeyValidate = 0
              Else
                 ' 8 = tasto cancella (freccia sinistra)
                 KeyValidate = KeyAscii
              End If
              ' Se è stato premuto il segno meno
              If KeyAscii = 45 Then
                KeyValidate = KeyAscii
              End If
           Else
              KeyValidate = KeyAscii
           End If
        
        Case Is = Alpha
            Select Case KeyAscii
               Case 8
                  KeyValidate = KeyAscii
               Case 97 To 122
                  KeyValidate = KeyAscii
               Case 65 To 90
                  KeyValidate = KeyAscii
               Case Else
                  KeyValidate = 0
            End Select
      
        Case Is = AlphaNumeric Or enmDataType = TimeDate
            KeyValidate = KeyAscii
    End Select

    If KeyValidate = 0 Or KeyValidate = 8 Or Len(TextBox.Text) = 0 Then Exit Function
    
    If Not TextBox Is Nothing And CancellaDuplicati <> "" Then
        TextBox.Text = TrimDUP(TextBox.Text, CancellaDuplicati)
    End If
    
     If Not TextBox Is Nothing And NumeroCaratteriMax <> 0 Then
        If Len(TextBox.Text) < NumeroCaratteriMax Then Exit Function
        PosCursore = TextBox.SelStart
        If PosCursore >= NumeroCaratteriMax Then
            KeyValidate = 0
        Else
            ' Cerco il carattere dopo il cursore del mouse
            strTmp = Mid(TextBox.Text, PosCursore + 1, 1)
            ' Cancello il carattere dopo il cursore del mouse
            strTmp = Replace(TextBox.Text, strTmp, "", PosCursore + 1, 1, vbTextCompare)
            ' Scrivo nel TextBox il nuovo testo senza il carattere dopo il cursore del mouse
            TextBox.Text = Left(TextBox.Text, PosCursore) & strTmp
            ' Imposto la posizione del cursore del mouse
            TextBox.SelStart = PosCursore
            ' Assegno il carattere
            KeyValidate = KeyAscii
        End If
    End If
    
End Function

Private Function TrimDUP(TextIN, Optional TrimChar = " ") As String
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


