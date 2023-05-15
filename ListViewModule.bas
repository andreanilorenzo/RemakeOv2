Attribute VB_Name = "ListViewModule"
' Obbliga a dichiarare le costanti
Option Explicit
' Rende le variabili publiche solo in questa applicazione
Option Private Module

'--------------------------------------------------------------------------------------------------------
' Per la funzione di modifica della cella con la Textbox
Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Type LVHITTESTINFO
  pt As POINTAPI
  Flags As Long
  iItem As Long
  iSubItem As Long
End Type

Public Const LVI_NOITEM = -1
Public Const LVIR_LABEL = 2

Public Const LVM_FIRST = &H1000
Public Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Public Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)

Public Const LVM_SETTEXTBKCOLOR As Long = (LVM_FIRST + 38)

Private Const WM_DESTROY = &H2
Private Const WM_KILLFOCUS = &H8

Private Const OLDWNDPROC = "OldWndProc"

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)

Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As KeyCodeConstants) As Integer
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'--------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------------
' Per le funzione NascondiScrollBar
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SB_HWND = 2
Private Const SB_BOTH = 3
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
'--------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------------
' Per le funzioni AutoSize
'Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'--------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------------
' Per la selezione di tutta la riga della ListView
Public Const LVS_EX_FULLROWSELECT = &H20
'Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'--------------------------------------------------------------------------------------------------------

' Color Constants
Public Const vbViolet = &HFF8080
Public Const vbVioletBright = &HFFC0C0
Public Const vbForestGreen = &H228B22
Public Const vbGray = &HE0E0E0
Public Const vbLightBlue = &HFFD3A4
Public Const vbLightGreen = &HABFCBD
Public Const vbGreenLemon = &HB3FFBE
Public Const vbYellowBright = &HC0FFFF
Public Const vbOrange = &H2CCDFC

' L'array per i dati della ListView
Public arrListView() As String
Public arrListViewVuoto As Boolean

Dim strIntestColonna() As String
Dim strIntestColListView() As String

Public Sub NascondiScrollBar(ListView1 As ListView, Optional Nascondi As Boolean = True)
    
    ShowScrollBar ListView1.hwnd, SB_HORZ, Not Nascondi
    ShowScrollBar ListView1.hwnd, SB_VERT, Not Nascondi
    
End Sub

Public Sub ListViewKillDupes(ListView As ListView)
    ' This code snippet is to remove or kill duplicates in vb6 listview.
    '
    ' Call it like this:
    '
    ' Private Sub Command1_Click()
    '   ListViewKillDupes (ListView)
    ' End Sub

    Dim Search1 As Long
    Dim Search2 As Long
    Dim KillDupe As Long
    
    KillDupe = 0
    
    For Search1& = 1 To ListView.ListItems.count - 1
        For Search2& = Search1& + 1 To ListView.ListItems.count - 1
            KillDupe = KillDupe + 1
            If ListView.ListItems.Item(Search1&) = ListView.ListItems.Item(Search2&) Then
                ListView.ListItems.remove (Search2&)
                Search2& = Search2& - 1
            End If
        Next Search2&
    Next Search1&
    
End Sub


Public Sub OrdinaColonnaByTag(ListView1 As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader, Optional Ordinamento As Integer = 0)
    On Error Resume Next
    
    With ListView1
        ' Display the hourglass cursor whilst sorting
        Dim lngCursor As Long
        lngCursor = .MousePointer
        .MousePointer = vbHourglass
        
        ' Prevent the ListView control from updating on screen -
        ' this is to hide the changes being made to the listitems and also to speed up the sort
        LockWindowUpdate .hwnd
        
        ' Check the data type of the column being sorted, and act accordingly
        Dim l As Long
        Dim strFormat As String
        Dim strData() As String
        
        Dim lngIndex As Long
        lngIndex = ColumnHeader.index - 1
    
        Select Case UCase$(ColumnHeader.Tag)
        
        Case "DATE"
            ' Sort by date.
            strFormat = "YYYYMMDDHhNnSs"
        
            ' Loop through the values in this column. Re-format
            ' the dates so as they can be sorted alphabetically,
            ' having already stored their visible values in the
            ' tag, along with the tag's original value
        
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), strFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .count
                        With .Item(l)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), strFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            ' Sort the list alphabetically by this column
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.index - 1
            .Sorted = True
            
            ' Restore the previous values to the 'cells' in this
            ' column of the list from the tags, and also restore
            ' the tags to their original values
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
            
        Case "NUMBER"
            ' Sort Numerically
            strFormat = String(30, "0") & "." & String(30, "0")
        
            ' Loop through the values in this column. Re-format the values so as they
            ' can be sorted alphabetically, having already stored their visible
            ' values in the tag, along with the tag's original value
        
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), strFormat)
                                Else
                                    .Text = "&" & InvNumber(Format(0 - CDbl(.Text), strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .count
                        With .Item(l)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), strFormat)
                                Else
                                    .Text = "&" & InvNumber(Format(0 - CDbl(.Text), strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            ' Sort the list alphabetically by this column
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.index - 1
            .Sorted = True
            
            ' Restore the previous values to the 'cells' in this
            ' column of the list from the tags, and also restore
            ' the tags to their original values
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
        
        Case Else   ' Assume sort by string
            
            ' Sort alphabetically. This is the only sort provided
            ' by the MS ListView control (at this time), and as
            ' such we don't really need to do much here
        
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.index - 1
            .Sorted = True
            
        End Select
    
        ' Unlock the list window so that the OCX can update it
        LockWindowUpdate 0&
        
        ' Restore the previous cursor
        .MousePointer = lngCursor
    
    End With
    
End Sub

Private Function InvNumber(ByVal Number As String) As String
    ' Function used to enable negative numbers to be sorted
    ' alphabetically by switching the characters
    
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
    
End Function

Public Sub OrdinaBySelezionati(ListView1 As ListView)
    
End Sub

Public Sub OrdinaColonna(ListView1 As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader, Optional Ordinamento As Integer = 0)
    '  Ordinamento:
    '  0  Automatico
    '  1  Crescente
    ' -1  Decrescente
    
    ' Si usa così:
    ' call OrdinaColonna(ListView1, ListView1.ColumnHeaders.Item(Colonna), 1)
    '
    Static iLast As Integer
    Dim iCur As Integer
    
    Select Case Ordinamento
    Case Is = 1
        With ListView1
            .Sorted = True
            iCur = ColumnHeader.index - 1
            .SortOrder = 0
            .SortKey = iCur
            iLast = iCur
        End With
        
    Case Is = -1
        With ListView1
            .Sorted = True
            iCur = ColumnHeader.index - 1
            .SortOrder = 1
            .SortKey = iCur
            iLast = iCur
        End With
        
    Case Else ' Predefinito
        With ListView1
            .Sorted = True
            iCur = ColumnHeader.index - 1
            If iCur = iLast Then .SortOrder = IIf(.SortOrder = 1, 0, 1)
            .SortKey = iCur
            iLast = iCur
        End With

    End Select
    
End Sub

Public Sub OrdinaColonnaByNumeroCol(ListView1 As ListView, Colonna As Long, Optional Ordinamento As Integer = 0)
    '  Ordinamento:
    '  0  Automatico
    '  1  Crescente
    ' -1  Decrescente
    
    Call OrdinaColonna(ListView1, ListView1.ColumnHeaders.Item(Colonna), Ordinamento)
        
End Sub

Public Sub NumeraListView(ListView As ListView)
    ' Inserisce la numerazione nella prima colonna della ListView
    Dim cnt As Integer

    For cnt = 1 To ListView.ListItems.count
        ListView.ListItems(cnt).Text = ElaboraNumeroRiga(cnt)
    Next
    
End Sub

Public Function ElaboraNumeroRiga(Valore As Variant, Optional nCifre As Integer = 5) As String
    ElaboraNumeroRiga = Format$(Valore, "00000")
End Function

Public Sub ChangeSelezMultipla(ListView As ListView)

    lVar(SelezMultipla) = Not Var(SelezMultipla).Valore
    ListView.MultiSelect = Var(SelezMultipla).Valore
    Call SelezionaRigaListView(ListView, ListView.SelectedItem.index)
    
End Sub

Public Function CancellaCeckedListView(ListView As ListView) As Integer
    Dim cnt As Long
    
    CancellaCeckedListView = 0

    ' Se non ci sono righe esco
    If ListView.ListItems.count = 0 Then Exit Function

    For cnt = ListView.ListItems.count To 1 Step -1
        If ListView.ListItems.Item(cnt).Checked = True Then
            ListView.ListItems.remove (cnt)
            CancellaCeckedListView = CancellaCeckedListView + 1
        End If
    Next

End Function

Public Sub SetCeckedListView(ListView1 As ListView, Optional AncheNeretto As Boolean = False, Optional ByVal bCecked As Boolean = False)
    Dim cnt As Long
        
    ' Se non ci sono righe esco
    If ListView1.ListItems.count = 0 Then Exit Sub

    LockWindowUpdate ListView1.hwnd

    For cnt = 1 To ListView1.ListItems.count
        ListView1.ListItems.Item(cnt).Checked = bCecked
    Next

    If AncheNeretto = True Then Call ControllaCheck(ListView1)
    
    LockWindowUpdate 0&
    
End Sub

Public Function ControllaRigaChecked(ListView1 As ListView, ByVal Item As MSComctlLib.ListItem, Optional Icona As Boolean = True) As Integer
    '
    Dim cnt As Long
    
    On Error Resume Next

    Set ListView1.SelectedItem = Item
    
    If Item.Checked Then
        ' user checked an item
        If Icona = True Then ListView1.ListItems.Item(ListView1.SelectedItem.index).SmallIcon = 6
        ListView1.ListItems.Item(ListView1.SelectedItem.index).Bold = True
    Else
        ' user unchecked an item
        If Icona = True Then ListView1.ListItems.Item(ListView1.SelectedItem.index).SmallIcon = 7
        ListView1.ListItems.Item(ListView1.SelectedItem.index).Bold = False
    End If
    
    For cnt = 1 To ListView1.ColumnHeaders.count - 1
        If Item.Checked Then
            ' user checked an item
            ListView1.ListItems.Item(ListView1.SelectedItem.index).ListSubItems(cnt).Bold = True
        Else
            ' user unchecked an item
            ListView1.ListItems.Item(ListView1.SelectedItem.index).ListSubItems(cnt).Bold = False
        End If
    Next
    
    Call AutoSizeColonne(ListView1, 0)
    
End Function

Public Sub ControllaCheck(ListView1 As ListView, Optional Icona As Boolean = True)
    ' Aggiunge il neretto e l'icona alla riga Checked listview
    Dim cnt As Long
    Dim cntCol As Long
    
    On Error Resume Next

    LockWindowUpdate ListView1.hwnd

    ' Scorro tutte le righe
    For cnt = 1 To ListView1.ListItems.count
        If Icona = True Then
            If ListView1.ListItems.Item(cnt).Checked Then
                ' user checked an item
                ListView1.ListItems.Item(cnt).SmallIcon = 6
            Else
                ' user unchecked an item
                ListView1.ListItems.Item(cnt).SmallIcon = 7
            End If
        End If
        
        ' Scorro tutte le colonne
        For cntCol = 1 To ListView1.ColumnHeaders.count - 1
            If ListView1.ListItems.Item(cnt).Checked Then
                'ListView1.ListItems.Item(ListView1.SelectedItem.Index).ListSubItems(cntCol).Bold = True
                ListView1.ListItems.Item(cnt).ListSubItems(cntCol).Bold = True
            Else
                ' user unchecked an item
                ListView1.ListItems.Item(cnt).ListSubItems(cntCol).Bold = False
            End If
        Next
    Next

    Call AutoSizeColonne(ListView1, 0)

    LockWindowUpdate 0&

End Sub

Public Function GetNumeroRigheChecked(ListView1 As ListView)
    Dim RigheChecked As Long
    Dim cnt As Long

    ' Scorro tutte le righe
    For cnt = 1 To ListView1.ListItems.count
        If ListView1.ListItems(cnt).Checked = True Then
            RigheChecked = RigheChecked + 1
        End If
    Next
    
    GetNumeroRigheChecked = RigheChecked
    
End Function

Public Sub CercaChecked(ListView1 As ListView, Optional Avanti As Boolean = True)
    '
    Dim cnt As Long
    Dim IniziaDa As Long
    
    ' Se non ci sono righe esco
    If ListView1.ListItems.count = 0 Then Exit Sub
    
    IniziaDa = ListView1.SelectedItem.index
    If Avanti = True Then
        IniziaDa = IniziaDa + 1
    Else
        IniziaDa = IniziaDa - 1
    End If
    If IniziaDa < -1 Then IniziaDa = 1
    
    If Avanti = True Then
        For cnt = IniziaDa To ListView1.ListItems.count
            If ListView1.ListItems.Item(cnt).Checked = True Then
                Call SelezionaRigaListView(ListView1, cnt)
                Exit For
            End If
        Next
    Else
        For cnt = IniziaDa To 1 Step -1
            If ListView1.ListItems.Item(cnt).Checked = True Then
                Call SelezionaRigaListView(ListView1, cnt)
                Exit For
            End If
        Next
    End If

End Sub

Public Sub ColoraSfondoRiga(ListView As ListView, Optional Colore As ColorConstants = vbRed, Optional Riga As Long)
    
    ' Se non ci sono righe esco
    If ListView.ListItems.count = 0 Then Exit Sub
    ' Se Riga è negativa esco
    If Riga < -1 Then Exit Sub
    ' Se Riga è maggiore del numero di righe esco
    If Riga > ListView.ListItems.count Then Exit Sub
    
    Call SendMessage(ListView.hwnd, LVM_SETTEXTBKCOLOR, 0&, Colore)

End Sub

Public Sub ColorListViewRow(lv As ListView, RowNbr As Long, RowColor As OLE_COLOR)
    ' Colora il carattere della riga di una ListView
    ' Inputs : lv - The ListView
    '         RowNbr - The index of the row to be colored
    '         RowColor - The color to color it (Esempio vbRed)
    Dim itmX As ListItem
    Dim lvSI As ListSubItem
    Dim intIndex As Integer
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    Set itmX = lv.ListItems(RowNbr)
    itmX.ForeColor = RowColor
    For intIndex = 1 To lv.ColumnHeaders.count - 1
        Set lvSI = itmX.ListSubItems(intIndex)
        lvSI.ForeColor = RowColor
        DoEvents
    Next
    If lv.ListItems.count > 1 Then lv.ListItems(2).Selected = True
    lv.ListItems(1).Selected = True

    Set itmX = Nothing
    Set lvSI = Nothing
    
    Exit Sub

Errore:
    GestErr Err, "Errore nella funzione ColorListViewRow."

End Sub

Public Sub FullRowSelection(pListview As Object)
    'Questa funzione definisce il metodo di selezione delle righe.
    ' Normalmente la ListView seleziona solo il main item di ogni riga.
    ' E' possibile farle selezionare tutta la riga.
    
    Dim lStyle As Long
    
    'Ottiene lo stile esteso corrente
    lStyle = SendMessageLong(pListview.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    
    'Modifica lo stile esteso invertendo il flag LVS_EX_FULLROWSELECT
    'In questo modo ogni volta che si fa click, se è attivo lo disattiva mentre se è disattivato lo attiva.
    'Se si desidera attivarlo e basta, utilizzare l'OR invece dello XOR
    lStyle = lStyle Or LVS_EX_FULLROWSELECT
    
    'Imposta il nuovo stile esteso
    Call SendMessageLong(pListview.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle)

End Sub

Public Function ListViewChecked(ListView As ListView, Riga As Long) As Boolean
    ' Per verificare se nella riga è stato impostato il ceck
    
    If Riga = 0 Or Riga > ListView.ListItems.count Then
        ListViewChecked = False
        Exit Function
    End If
    
    ListView.ListItems.Item(Riga).Selected = True
    
    If ListView.ListItems(Riga).Checked Then
        ListViewChecked = True
    Else
        ListViewChecked = False
    End If
    
End Function

Public Sub CaricaArrayDaListView(ListView As ListView, Optional SaveCheck As Boolean = False)
    Dim cnt As Long
    Dim cntCol As Long
    Dim nColonne As Long
    
    ' Se non ci sono righe esco
    If ListView.ListItems.count = 0 Then Exit Sub
    
    ' Conto le colonne
    If SaveCheck = True Then
        nColonne = ListView.ColumnHeaders.count
    Else
        nColonne = ListView.ColumnHeaders.count - 1
    End If

    ' Cancello l'array
    ReDim arrListView(ListView.ListItems.count, nColonne)
    
    ' Scorro tutte le righe.....
    For cnt = 0 To ListView.ListItems.count
        ' Scorro tutte le colonne.....
        For cntCol = 0 To nColonne
            If cnt = 0 Then
                ' Inserisco le intestazioni di colonna
                If (SaveCheck = False) Or (SaveCheck = True And cntCol < nColonne) Then
                    arrListView(cnt, cntCol) = Trim(ListView.ColumnHeaders.Item(cntCol + 1))
                Else
                    arrListView(cnt, cntCol) = "%Check%"
                End If
            Else
                ' Inserisco le altre righe
                If cntCol = 0 Then
                    ' per la prima colonna
                    arrListView(cnt, cntCol) = Trim(ListView.ListItems(cnt))
                ElseIf cntCol < nColonne Then
                    ' per le altre colonne
                    arrListView(cnt, cntCol) = Trim(ListView.ListItems(cnt).SubItems(cntCol))
                ElseIf cntCol = nColonne Then
                    ' per l'ultima colonna
                    If SaveCheck = False Then
                        arrListView(cnt, cntCol) = Trim(ListView.ListItems(cnt).SubItems(cntCol))
                    Else
                        ' salvo il check
                        arrListView(cnt, cntCol) = ListView.ListItems(cnt).Checked
                    End If
                End If
            End If
        Next
    Next
    
    arrListViewVuoto = False
    
End Sub

Public Sub CaricaListViewDaArray(ListView As ListView, Optional CancellaListView As Boolean = True, Optional SaveCheck As Boolean = False, Optional AdattaLarghezzaUltimaColonna As Boolean = True)
    Dim cntRecord As Long
    Dim itmX As ListItem
    Dim cnt As Long
    Dim cnt1 As Long
    Dim cntRow As Long
    Dim cntCol As Long
    Dim NumeriColonna As Boolean
    Dim start As Long
    Dim strDiffColonne() As String
    Dim bTrovato As Boolean
    Dim strTmp As String
    Dim n As Long

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    ' Se l'array è vuoto esco
    If arrListViewVuoto = True Then Exit Sub
    
    ' Cancello la ListView
    If CancellaListView = True Then
        ListView.ListItems.Clear
        DoEvents
        cntRow = 0
    Else
        ' Prendo il numero più alto alto della lista
        If ListView.ListItems.count > 0 Then
            For cnt = 1 To ListView.ListItems.count
                If ListView.ListItems(cnt) > cntRow Then
                    cntRow = ListView.ListItems(cnt)
                End If
            Next
        End If
    End If
    
    If SaveCheck = True Then
        n = 1
    Else
        n = 0
    End If
    
    ' Controllo se la seconda riga della prima colonna contiene il numero di intestazione della riga
    If IsNumeric(arrListView(1, 0)) Then
        NumeriColonna = True
    Else
        NumeriColonna = True
        ' Aggiungo una collonna all'array
        ReDim Preserve arrListView(UBound(arrListView, 1), UBound(arrListView, 2) + 1)
        ' Sposto i valori contenuti nelle colonne per lasciare la colonna vuota in prima posizione
        For cnt = 0 To UBound(arrListView, 1)
            For cntCol = UBound(arrListView, 2) To 0 Step -1
                If cntCol = 0 Then
                    If cnt = 0 Then
                        arrListView(cnt, 0) = "    "
                    Else
                        ' Scrivo le intestazioni delle righe
                        cntRow = cntRow + 1
                        arrListView(cnt, 0) = cntRow
                    End If
                Else
                    arrListView(cnt, cntCol) = arrListView(cnt, cntCol - 1)
                End If
            Next
        Next
    End If
    
    ' Carico le intestazioni di colonna dell'array
    ReDim strIntestColonna(UBound(arrListView, 2))
    For cntCol = 0 To UBound(arrListView, 2)
        strTmp = arrListView(0, cntCol)
        ' se il titolo della colonna non è vuoto....
        If strTmp <> "" Then
            strIntestColonna(cntCol) = arrListView(0, cntCol)
            
        Else
            ' ....creo un titolo per la colonna
            strIntestColonna(cntCol) = cntCol + 1 & " Colonna: " & cntCol + 1
        End If
    Next
    
    ' Carico le intestazioni di colonna della ListView
    ReDim strIntestColListView(ListView.ColumnHeaders.count - 1)
    For cnt = 0 To ListView.ColumnHeaders.count - 1
        strIntestColListView(cnt) = Trim(ListView.ColumnHeaders(cnt + 1).Text)
    Next
    
    ' Se mancano delle colonne le aggiungo -------------------------------------------
    If UBound(strIntestColListView) < UBound(strIntestColonna) Then

        ' Cerco le colonne che sono nell'Array ma non sono nella ListView
        cnt1 = 0
        For cnt = 0 To UBound(strIntestColonna)
            bTrovato = False
            For cntCol = 0 To UBound(strIntestColListView)
                If strIntestColonna(cnt) = strIntestColListView(cntCol) Then
                    bTrovato = True
                    Exit For
                End If
            Next
            If bTrovato = False Then
                ReDim Preserve strDiffColonne(cnt1)
                ' ...e le inserisco nell'array strDiffColonne
                strDiffColonne(cnt1) = strIntestColonna(cnt)
                cnt1 = cnt1 + 1
            End If
        Next

        For cntCol = 0 To UBound(strIntestColonna)
            If (cntCol > UBound(strIntestColListView)) And strIntestColonna(cntCol) <> "%Check%" Then
              ' Aggiungo la colonna alla fine.... ma non quella per il check ("%Check%")
              ListView.ColumnHeaders.Add , , cntCol & " *%*%*"
            End If
        Next

        ' Cambio il nome della colonna aggiunta
        cnt1 = 0
        For cntCol = 0 To ListView.ColumnHeaders.count - 1
            ' Se trovo una colonna inserita.....
            If Right$(Trim(ListView.ColumnHeaders(cntCol + 1).Text), 5) = "*%*%*" Then
                ListView.ColumnHeaders(cntCol + 1).Text = strDiffColonne(cnt1)
                cnt1 = cnt1 + 1
            End If
        Next
        
        DoEvents
        
        ' Carico di nuovo le intestazioni di colonna della ListView
        For cnt = 0 To ListView.ColumnHeaders.count - 1
            ReDim Preserve strIntestColListView(cnt)
            strIntestColListView(cnt) = Trim(ListView.ColumnHeaders(cnt + 1))
        Next
    End If '-------------------------------------------
        
    If MultiDimensional(arrListView) = False Then ' Array ad una dimensione---------
        For cnt = 0 To UBound(arrListView, 1)
          Set itmX = ListView.ListItems.Add(, , Format(cntRecord + 1, "00000"))
          itmX.SubItems(cntCol) = arrListView(cnt, 1)
          cntRecord = cntRecord + 1
        Next cnt
    
    Else ' Array a più dimensioni----------------------------------------------------
        For cnt = 0 To UBound(arrListView, 1)
            ' Se siamo nella prima riga...
            If cnt = 0 Then
                'Niente
            Else
              cntCol = 0
              ' Aggiungo la riga con la prima colonna
              Set itmX = ListView.ListItems.Add(, , Format(CLng(arrListView(cnt, 0)), "00000"))
              
              ' Scrivo i dati nelle colonne
              For cnt1 = 0 To UBound(arrListView, 2)
                  If cntCol <> 0 Then
                      ' Scrivo i dati nelle altre colonne
                      If (SaveCheck = False) Or (SaveCheck = True And cntCol < UBound(arrListView, 2)) Then
                        If arrListView(cnt, cntCol) = "" Then
                            itmX.SubItems(CercaColonna(strIntestColonna(cntCol))) = ""
                        Else
                            itmX.SubItems(CercaColonna(strIntestColonna(cntCol))) = arrListView(cnt, cntCol)
                        End If
                      Else
                         ' Imposto il check
                         ListView.ListItems(cnt).Checked = arrListView(cnt, cntCol)
                      End If
                  End If
                  cntCol = cntCol + 1
              Next
              
            cntRecord = cntRecord + 1
            End If
        Next
    End If
    
    If AdattaLarghezzaUltimaColonna = True Then Call AutoSizeUltimaColonna(ListView)
    
    SelezionaRigaListView ListView, ListView.ListItems.count
    
Esci:
    Set itmX = Nothing
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione CaricaListViewDaArray."
    GoTo Esci
    
End Sub

Private Function CercaColonna(strDaCercare As String) As Long
    Dim cnt As Long
    
    For cnt = 0 To UBound(strIntestColListView)
        If strIntestColListView(cnt) = strDaCercare Then
            CercaColonna = cnt
            Exit Function
        End If
    Next
    
    ' Se non trovo la colonna assegno l'ultima
    CercaColonna = UBound(strIntestColListView)
    
End Function

Public Function GetNumeroRighe(ListView As ListView) As Long
    GetNumeroRighe = ListView.ListItems.count
End Function

Public Function GetValoreRiga(ListView As ListView, Riga As Long, Optional Separatore As String = vbTab) As String
    ' Restituisce una stringa contenente il valore dell'intera riga
    Dim cnt As Long
    Dim strTmp As String
    Dim valTmp As String
    
    For cnt = 1 To ListView.ColumnHeaders.count - 1
        If cnt > 1 Then strTmp = strTmp & Separatore
        valTmp = ListView.ListItems.Item(Riga).SubItems(cnt)
        strTmp = strTmp & ListView.ListItems.Item(Riga).SubItems(cnt)
    Next
    
    GetValoreRiga = strTmp
    
End Function

Public Function GetRigaSelezionata(ListView As ListView, X As Single, Y As Single) As Long
    ' Restituisce il numero di riga data la posizione del mouse
    Dim Item As ListItem

    If ListView.ListItems.count = 0 Then
        GetRigaSelezionata = 0
        GoTo Esci
    End If

    ' Seleziono la riga nella ListView
    Set Item = ListView.HitTest(X, Y)
    
    If Item Is Nothing Then
        Set Item = ListView.SelectedItem
        If Not Item Is Nothing Then
            Item.Selected = False
            GetRigaSelezionata = 0
        End If
    Else
        Set ListView.SelectedItem = Item
        GetRigaSelezionata = Item.index
    End If
    
Esci:
    Set Item = Nothing
    DoEvents

End Function

Public Function GetValoreCella(ListView As ListView, ByVal Riga As Long, ByVal Colonna As Long) As String
    ' Restituisce il valore contenuto in una cella
    
    ' Se non ci sono righe esco
    If ListView.ListItems.count = 0 Or ControllaCella(ListView, Riga, Colonna) = False Then
        GetValoreCella = ""
        Exit Function
    End If

    If Colonna = 0 Then
        ' per la prima colonna
        GetValoreCella = ListView.ListItems(Riga)
    Else
        ' per le altre colonne
        GetValoreCella = ListView.ListItems(Riga).SubItems(Colonna)
    End If

End Function

Public Function GetValoreCellaByNome(ListView As ListView, ByVal Riga As Long, ByVal sNomeColonna As String) As Variant
    ' Restituisce il valore contenuto in una cella
    Dim Colonna As Long
    
    ' Se non ci sono righe esco
    If ListView.ListItems.count = 0 Then
        Exit Function
    End If
    
    Colonna = GetNumColDaIntestazione(ListView, sNomeColonna)
    If Colonna > 0 Then
        GetValoreCellaByNome = GetValoreCella(ListView, Riga, Colonna - 1)
    End If
    
End Function

Public Function GetValoreCellaByMouseMove(ListView As ListView, ByVal NomeColonna As String, ByVal X As Single, ByVal Y As Single) As String
    Dim HitItem As ListItem
    Dim nCol As Integer

    With ListView
        nCol = GetNumColDaIntestazione(ListView, NomeColonna)
        Set HitItem = .HitTest(X, Y)
        If Not HitItem Is Nothing Then
            GetValoreCellaByMouseMove = HitItem.ListSubItems(nCol - 1).Text
        End If
    End With

End Function

Public Function GetNumColDaIntestazione(ListView As ListView, ByVal Intestazione As String) As Long
    ' Restituisce:
    ' x: numero della colonna se la colonna è stata trovata
    ' 0: se la colonna non è stata trovata
    ' -x: il numero delle colonne che sono state trovate (se ne ha trovate più di una)
    Dim cntCol As Long
    Dim ColTrovate As Long
    Dim PosTrovata As Variant
    
    ColTrovate = 0
    
    ' Scorro tutte le colonne (esclusa la prima)
    For cntCol = 2 To ListView.ColumnHeaders.count
        PosTrovata = InStr(1, LCase(ListView.ColumnHeaders.Item(cntCol)), LCase(Trim(Intestazione)), vbTextCompare)
        ' Se trovo l'intestazione di colonna simile alla stringa da cercare......
        If PosTrovata <> 0 Then
            GetNumColDaIntestazione = cntCol
            ColTrovate = ColTrovate + 1
        End If
    Next
    
    Select Case ColTrovate
        Case Is = 0
            ' Nessuna colonna trovata
            GetNumColDaIntestazione = 0
        Case Is = 1
            ' Il valore è già stato impostato nel ciclo For
        Case Is > 1
            ' Il numero delle colonne trovate in negativo
            GetNumColDaIntestazione = -ColTrovate
        Case Else
            ' Non si sa mai...
            GetNumColDaIntestazione = 0
    End Select

End Function

Public Sub ScriviCella(ListView As ListView, Riga As Long, Colonna As Long, Valore As String, Optional CreaSeNonEsiste As Boolean = False)
    ' Scrive in una cella di una colonna e di una riga già esistenti
    ' Non scrive nella prima colonna e nell'intestazione
    
    'On Error Resume Next
    
    ' Se Riga è 0 oppure e maggiore del numero di righe totali esco
    If Riga = 0 Or Riga > ListView.ListItems.count Then Exit Sub
    ' Se Colonna è 0 oppure e maggiore del numero di colonne totali esco
    If Colonna = 0 Or Colonna > ListView.ColumnHeaders.count Then Exit Sub
    
    ListView.ListItems(Riga).SubItems(Colonna - 1) = Valore

End Sub

Public Function TotaleRigheListView(ListView1 As ListView) As Long
    
    TotaleRigheListView = ListView1.ListItems.count
    
    If TotaleRigheListView = 0 Then
        ListView1.ColumnHeaders.Item(1).Text = ""
    Else
        ListView1.ColumnHeaders.Item(1).Text = " Tot: " & TotaleRigheListView
    End If
    
End Function

Public Function CercaEvidenzia(ListView As ListView, ByVal Cerca As String, Optional ByVal Colonna As Integer = 1, Optional ByVal OrdinaDopo As Boolean = True) As Integer
    ' Cerca il valore Cerca nella lista ed evidenzia le righe che lo contengono
    '
    ' Restituisce:
    '  x numero di righe che sono state evidenziate
    ' -1 in caso di errore
    Dim cntRiga As Long
    Dim cntColonna As Long
    Dim cntTrovati As Integer: cntTrovati = 0
    
    Cerca = LCase(Cerca)
    ListView.MultiSelect = True
    
    ' Se Cerca è vuoto oppure non ci sono righe esco dalla funzione
    If Cerca = "" Or ListView.ListItems.count = 0 Then
        CercaEvidenzia = -1
        Exit Function
    End If

    ' Scorro tutte le righe.....
    For cntRiga = 1 To ListView.ListItems.count
        If InStr(1, LCase(ListView.ListItems(cntRiga).SubItems(Colonna)), Cerca, vbTextCompare) <> 0 Then
            ' Seleziono la riga
            ListView.ListItems(cntRiga).Selected = True
            cntTrovati = cntTrovati + 1
        Else
            ListView.ListItems(cntRiga).Selected = False
        End If
    Next
    
    If OrdinaDopo = True Then Call OrdinaBySelezionati(ListView)
    CercaEvidenzia = cntTrovati
    ListView.SetFocus
    
End Function

Public Function CercaSostituisci(ListView As ListView, ByVal Cerca As String, ByVal Sostituisci As String, Optional Riga As Long = -1, Optional Colonna As Long = -1, Optional AncheIntestazione As Boolean = False, Optional AnchePrimaColonna As Boolean = False, Optional PartiDa As Long = 0) As Integer
    ' Restituisce:
    '  1 se tutto ok
    ' -1 in caso di errore
    '
    ' PartiDa:
    '  0 sostituisce tutte le stringhe trovate
    '  1 sostituisce solo la prima stringa a partire da sinistra
    ' -1 sostituisce solo la prima stringa a partire da destra
    Dim cntRiga As Long
    Dim cntColonna As Long
    Dim strTmp As String
    Dim PartiDaTmp As Long
    
    If PartiDa = 0 Then
        PartiDaTmp = -1 ' Verranno eseguite tutte le sostituzioni possibili nella funzione Replace
    Else
        PartiDaTmp = PartiDa
    End If
    
    If PartiDa < 0 Then
        Cerca = StrReverse(Cerca)
        Sostituisci = StrReverse(Sostituisci)
    End If
    
    ' Se Cerca è vuoto oppure non ci sono righe esco dalla funzione
    If Cerca = "" Or ListView.ListItems.count = 0 Then
        CercaSostituisci = -1
        Exit Function
    End If
    
    ' Cerco nella riga---------------------------------------------------------------------------------
    If Riga <> -1 And Riga > 0 Then
        ' Ancora da fare............................
        CercaSostituisci = -1
        Exit Function
    End If

    ' Cerco nella colonna---------------------------------------------------------------------------------
    If (Colonna <> -1 And Colonna >= 0) And (AnchePrimaColonna = False And Colonna <> 0) Then
        ' Scorro tutte le righe.....
        For cntRiga = 0 To ListView.ListItems.count
            If cntRiga = 0 And AncheIntestazione = True Then
                ' Le intestazioni di colonna---------------------------------
                If PartiDa >= 0 Then
                    strTmp = Replace(ListView.ColumnHeaders.Item(Colonna).Text, Cerca, Sostituisci, , PartiDaTmp, vbTextCompare)
                Else
                    ' Prendo il valore assoluto
                    PartiDaTmp = Abs(PartiDaTmp)
                    strTmp = Replace(StrReverse(ListView.ColumnHeaders.Item(Colonna).Text), Cerca, Sostituisci, , PartiDaTmp, vbTextCompare)
                    strTmp = StrReverse(strTmp)
                End If
                ListView.ColumnHeaders.Item(Colonna).Text = strTmp
            End If
            If cntRiga > 0 Then
                ' Le altre righe...............
                If Colonna = 0 Then
                    ' per la prima colonna-------------
                    If PartiDa >= 0 Then
                        strTmp = Replace(ListView.ListItems(cntRiga).Text, Cerca, Sostituisci, , PartiDaTmp, vbTextCompare)
                    Else
                        ' Prendo il valore assoluto
                        PartiDaTmp = Abs(PartiDaTmp)
                        strTmp = Replace(StrReverse(ListView.ListItems(cntRiga).Text), Cerca, Sostituisci, , PartiDaTmp, vbTextCompare)
                        strTmp = StrReverse(strTmp)
                    End If
                    ListView.ListItems(cntRiga).Text = strTmp
                Else
                    ' per le altre colonne-------------
                    If PartiDa >= 0 Then
                        strTmp = Replace(ListView.ListItems(cntRiga).SubItems(Colonna - 1), Cerca, Sostituisci, , PartiDaTmp, vbTextCompare)
                    Else
                        ' Prendo il valore assoluto
                        PartiDaTmp = Abs(PartiDaTmp)
                        strTmp = Replace(StrReverse(ListView.ListItems(cntRiga).SubItems(Colonna - 1)), Cerca, Sostituisci, , PartiDaTmp, vbTextCompare)
                        strTmp = StrReverse(strTmp)
                    End If
                    ListView.ListItems(cntRiga).SubItems(Colonna - 1) = strTmp
                End If
            End If
        Next
        
        CercaSostituisci = 1
    Else
        CercaSostituisci = -1
    End If

    ' Tutte le righe e tutte le colonnne---------------------------------------------------------------------------------
    If Riga = -1 And Colonna = -1 Then
        ' Scorro tutte le righe.....
        For cntRiga = 0 To ListView.ListItems.count
            ' Scorro tutte le colonne.....
            For cntColonna = 0 To ListView.ColumnHeaders.count - 1

                If cntRiga = 0 And AncheIntestazione = True Then
                    ' Le intestazioni di colonna
                    If PartiDa >= 0 Then
                        strTmp = Replace(ListView.ColumnHeaders.Item(cntColonna + 1).Text, Cerca, Sostituisci, , , vbTextCompare)
                    Else
                        ' Prendo il valore assoluto
                        PartiDaTmp = Abs(PartiDaTmp)
                        strTmp = Replace(StrReverse(ListView.ColumnHeaders.Item(cntColonna + 1).Text), Cerca, Sostituisci, , , vbTextCompare)
                        strTmp = StrReverse(strTmp)
                    End If
                    ListView.ColumnHeaders.Item(cntColonna + 1).Text = strTmp
                End If
                If cntRiga > 0 Then
                    ' Le altre righe
                    If cntColonna = 0 And AnchePrimaColonna = True Then
                        ' per la prima colonna
                        If PartiDa >= 0 Then
                            strTmp = Replace(ListView.ListItems(cntRiga).Text, Cerca, Sostituisci, , , vbTextCompare)
                        Else
                            ' Prendo il valore assoluto
                            PartiDaTmp = Abs(PartiDaTmp)
                            strTmp = Replace(StrReverse(ListView.ListItems(cntRiga).Text), Cerca, Sostituisci, , , vbTextCompare)
                            strTmp = StrReverse(strTmp)
                        End If
                        ListView.ListItems(cntRiga).Text = strTmp
                    End If
                    If cntColonna > 0 Then
                        ' per le altre colonne
                        If PartiDa >= 0 Then
                            strTmp = Replace(ListView.ListItems(cntRiga).SubItems(cntColonna), Cerca, Sostituisci, , , vbTextCompare)
                        Else
                            ' Prendo il valore assoluto
                            PartiDaTmp = Abs(PartiDaTmp)
                            strTmp = Replace(StrReverse(ListView.ListItems(cntRiga).SubItems(cntColonna)), Cerca, Sostituisci, , , vbTextCompare)
                            strTmp = StrReverse(strTmp)
                        End If
                        ListView.ListItems(cntRiga).SubItems(cntColonna) = strTmp
                    End If
                End If
            Next
        Next
        CercaSostituisci = 1
    Else
        CercaSostituisci = -1
    End If


End Function

Public Sub SelezionaRigaListView(ListView As ListView, Riga As Long, Optional AssicuratiCheSiaVisibile As Boolean = True)
    Dim Item As ListItem
    
    ' Se non ci sono righe esco
    If ListView.ListItems.count = 0 Then Exit Sub
    
    If Riga <= 0 Then Exit Sub
    
    If Riga <= ListView.ListItems.count Then
        Set Item = ListView.ListItems(Riga)
    End If
    
    If Not Item Is Nothing Then
        Set ListView.SelectedItem = Item
        If AssicuratiCheSiaVisibile = True Then ListView.SelectedItem.EnsureVisible
    End If

    Set Item = Nothing

End Sub

Public Sub CancellaRiga(ListView As ListView, Optional Riga As Long = -1, Optional MultiDelete As Boolean = False)
    Dim index As Integer
    
    ' Se non ci sono righe esco
    If ListView.ListItems.count = 0 Then Exit Sub
    ' Se Riga è negativa esco
    If Riga < -1 Then Exit Sub
    ' Se Riga è maggiore del numero di righe esco
    If Riga > ListView.ListItems.count Then Exit Sub
    
    If MultiDelete = False Then
        If Riga = -1 Then
            ListView.ListItems.remove (ListView.SelectedItem.index)
        Else
            ListView.ListItems.remove (Riga)
        End If
    Else
        ' Delete multiple row
        For index = ListView.ListItems.count To 1 Step -1
          If ListView.ListItems(index).Selected Then
              ListView.ListItems.remove (index)
          End If
        Next
    End If

    If ListView.ListItems.count = 0 Then Exit Sub
    ' Seleziono la riga
    Call SelezionaRigaListView(ListView, ListView.SelectedItem.index)
    ListView.SetFocus

End Sub

Public Function ControllaCella(ListView As ListView, Riga As Long, Colonna As Long, Optional SoloNonEsiste As Boolean = False) As Boolean
    ' Un modo un po rudimentale per verificare se una cella è vuota oppure non esiste
    
    On Error GoTo Error
    Dim Val As String
    
    Val = (ListView.ListItems.Item(Riga).ListSubItems.Item(Colonna).Text)
    
    If SoloNonEsiste = False Then
        If Val <> "" Then
            ControllaCella = True
        Else
            ControllaCella = False
        End If
    Else
        ControllaCella = True
    End If

    Exit Function

Error:
    ControllaCella = False
    
End Function

Public Function ListView_GetSubItemRect(hwnd As Long, iItem As Long, iSubItem As Long, code As Long, prc As RECT) As Boolean
    ' Per la funzione di modifica della cella con la Textbox
    prc.Top = iSubItem
    prc.Left = code
    ListView_GetSubItemRect = SendMessage(hwnd, LVM_GETSUBITEMRECT, ByVal iItem, prc)
    
End Function

Public Function ListView_SubItemHitTest(hwnd As Long, plvhti As LVHITTESTINFO) As Long
    ' Per la funzione di modifica della cella con la Textbox
    ListView_SubItemHitTest = SendMessage(hwnd, LVM_SUBITEMHITTEST, 0, plvhti)
  
End Function

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' Per la funzione di modifica della cella con la Textbox
  
      Select Case uMsg
    
        ' Hide the TextBox when it loses focus (its LostFocus event it not fired
        ' when losing focus to a window outside the app).
        Case WM_KILLFOCUS
          ' OLDWNDPROC will be gone after UnSubClass is called, HideTextBox
          ' calls UnSubClass.
          Call CallWindowProc(GetProp(hwnd, OLDWNDPROC), hwnd, uMsg, wParam, lParam)
          Call frmWeb.HideTextBox(True)
          Exit Function
        
        ' Unsubclass the window when it's destroyed in case someone forgot...
        Case WM_DESTROY
          ' OLDWNDPROC will be gone after UnSubClass is called!
          Call CallWindowProc(GetProp(hwnd, OLDWNDPROC), hwnd, uMsg, wParam, lParam)
          Call UnSubClass(hwnd)
          Exit Function
          
      End Select
      
      WndProc = CallWindowProc(GetProp(hwnd, OLDWNDPROC), hwnd, uMsg, wParam, lParam)
      
End Function

Public Function SubClass(hwnd As Long, lpfnNew As Long) As Boolean
  ' Per la funzione di modifica della cella con la Textbox

  Dim lpfnOld As Long
  Dim fSuccess As Boolean
  
  If (GetProp(hwnd, OLDWNDPROC) = 0) Then
    lpfnOld = SetWindowLong(hwnd, GWL_WNDPROC, lpfnNew)
    If lpfnOld Then
      fSuccess = SetProp(hwnd, OLDWNDPROC, lpfnOld)
    End If
  End If
  
  If fSuccess Then
    SubClass = True
  Else
    If lpfnOld Then Call UnSubClass(hwnd)
    MsgBox "Unable to successfully subclass &H" & Hex(hwnd), vbCritical
  End If
  
End Function

Public Function UnSubClass(hwnd As Long) As Boolean
  ' Per la funzione di modifica della cella con la Textbox

  Dim lpfnOld As Long
  
  lpfnOld = GetProp(hwnd, OLDWNDPROC)
  If lpfnOld Then
    If RemoveProp(hwnd, OLDWNDPROC) Then
      UnSubClass = SetWindowLong(hwnd, GWL_WNDPROC, lpfnOld)
    End If
  End If

End Function

Public Sub AutoSizeColonne(ListView As ListView, Optional ByVal Colonna As Long = -1, Optional ByVal EscludiPrimaColonna As Boolean = True, Optional ByVal MisuraMinima As Integer = -1)
    Dim col2adjust As Long
    Dim LstBold As Boolean
    Dim i As Integer
   
    'On Error Resume Next
   
    'Size each column based on the width
    'of the widest list item in the column.
    'If the items are shorter than the column
    'header text, the header text is truncated.
    
    If EscludiPrimaColonna = True Then
        i = 1
    Else
        i = 0
    End If
    
    If Colonna = -1 Then
        For col2adjust = i To ListView.ColumnHeaders.count - 1
            ListView.Font.Bold = True
            Call SendMessage(ListView.hwnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE)
            ListView.Font.Bold = False
            If MisuraMinima > 0 And ListView.ColumnHeaders.Item(col2adjust).Width < MisuraMinima Then
                ListView.ColumnHeaders.Item(col2adjust).Width = MisuraMinima
            End If
        Next
        
    Else
        ListView.Font.Bold = True
        Call SendMessage(ListView.hwnd, LVM_SETCOLUMNWIDTH, Colonna - 1, ByVal LVSCW_AUTOSIZE)
        ListView.Font.Bold = False
        If Colonna > 0 And MisuraMinima > 0 And ListView.ColumnHeaders.Item(Colonna - 1).Width < MisuraMinima Then
            ListView.ColumnHeaders.Item(Colonna).Width = MisuraMinima
        End If
    
    End If
    
End Sub

Public Sub AutoSizeIntestazione(ListView As ListView)
   Dim col2adjust As Long
  'Size each column based on the maximum of
  'EITHER the column header text width, or,
  'if the items below it are wider, the
  'widest list item in the column.
  '
  'The last column is always resized to occupy
  'the remaining width in the control.

   For col2adjust = 0 To ListView.ColumnHeaders.count - 1
      Call SendMessage(ListView.hwnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next
   
End Sub

Public Sub AutoSizeUltimaColonna(ListView As ListView)
   Dim col2adjust As Long
  'Because applying LVSCW_AUTOSIZE_USEHEADER
  'to the last column in the control always
  'sets its width to the maximum remaining control
  'space, calling SendMessage passing just the
  'last column index will resize only the last column,
  'resulting in a listview utilizing the full
  'control width space. To see this resizing in practice,
  'create a wide listview and press the "Size to Contents"
  'button followed by the "Maximize Width" button.
  '
  'By explanation:  if a listview had a total width of 2000
  'and the first three columns each had individual widths of 250,
  'calling this will cause the last column to widen
  'to cover the remaining 1250.

  'Calling this will force the data to remain within the
  'listview. If a column other than the last column is
  'widened or narrowed, the last column will become
  'sized to ensure all data remains within the control.
  'This could truncate text depending on the overall
  'widths of the other columns; the minimum width is still
  'based on the length of that column's header text.
   
   col2adjust = ListView.ColumnHeaders.count - 1
   
   Call SendMessage(ListView.hwnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)

End Sub

Public Sub SetListViewColor(pCtrlListView As ListView, pCtrlPictureBox As PictureBox, Optional AltezzaRiga As Long = 2.75, Optional Color1 As Long = vbWhite, Optional Color2 As Long = vbLightGreen)

On Error GoTo SetListViewColor_Error

    Dim iLineHeight As Long
    Dim iBarHeight  As Long
    Dim lBarWidth   As Long
    Dim lColor1     As Long
    Dim lColor2     As Long
 
    lColor1 = Color1
    lColor2 = Color2
    
    pCtrlPictureBox.Height = 180
    pCtrlPictureBox.Width = 165
    
    If pCtrlListView.View = lvwReport Then
        pCtrlListView.Picture = LoadPicture("")
        pCtrlListView.Refresh
        pCtrlPictureBox.Cls
        
        pCtrlPictureBox.AutoRedraw = True
        pCtrlPictureBox.BorderStyle = vbBSNone
        pCtrlPictureBox.ScaleMode = vbTwips
        pCtrlPictureBox.Visible = False
        
        pCtrlListView.PictureAlignment = lvwTile
        pCtrlPictureBox.Font = pCtrlListView.Font
        pCtrlPictureBox.Top = pCtrlListView.Top
        pCtrlPictureBox.Font = pCtrlListView.Font
        With pCtrlPictureBox.Font
            .Size = pCtrlListView.Font.Size + AltezzaRiga
            .Bold = pCtrlListView.Font.Bold
            .Charset = pCtrlListView.Font.Charset
            .Italic = pCtrlListView.Font.Italic
            .Name = pCtrlListView.Font.Name
            .Strikethrough = pCtrlListView.Font.Strikethrough
            .Underline = pCtrlListView.Font.Underline
            .Weight = pCtrlListView.Font.Weight
        End With
        pCtrlPictureBox.Refresh
        iLineHeight = pCtrlPictureBox.TextHeight("W") + Screen.TwipsPerPixelY
    
        iBarHeight = (iLineHeight * 1)
        lBarWidth = pCtrlListView.Width
    
        pCtrlPictureBox.Height = iBarHeight * 2
        pCtrlPictureBox.Width = lBarWidth
    
        'paint the two bars of color
        pCtrlPictureBox.Line (0, 0)-(lBarWidth, iBarHeight), lColor1, BF
        pCtrlPictureBox.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), lColor2, BF
        
        pCtrlPictureBox.AutoSize = True
        'set the pCtrlListView picture to the
        'pCtrlPictureBox image
        pCtrlListView.Picture = pCtrlPictureBox.image
    Else
        pCtrlListView.Picture = LoadPicture("")
    End If
    
    pCtrlListView.Refresh
    Exit Sub
    
SetListViewColor_Error:
    'clear pCtrlListView's picture and then exit
    pCtrlListView.Picture = LoadPicture("")
    pCtrlListView.Refresh
    
End Sub

Public Function InvertiColonne(ListView1 As ListView, ByVal Colonna1 As Integer, ByVal Colonna2 As Integer) As Boolean
    Dim cnt As Long
    Dim sCol1 As String
    Dim sCol2 As String
    
    If (ListView1.ListItems.count = 0) Or (Colonna1 = Colonna2) Then
        InvertiColonne = False
        Exit Function
    End If
    
    LockWindowUpdate ListView1.hwnd

    For cnt = 1 To ListView1.ListItems.count
        sCol1 = ListView1.ListItems(cnt).SubItems(Colonna1)
        sCol2 = ListView1.ListItems(cnt).SubItems(Colonna2)
        ListView1.ListItems(cnt).SubItems(Colonna1) = sCol2
        ListView1.ListItems(cnt).SubItems(Colonna2) = sCol1
    Next

    LockWindowUpdate 0&
    
    InvertiColonne = True

End Function

Public Function SommaColonne(ListView1 As ListView, ByVal Colonna1 As Integer, ByVal Colonna2 As Integer, Optional CancellaCol2 As Boolean = False) As Boolean
    ' Aggiunge i dati della colonna2 nella colonna1
    Dim cnt As Long
    Dim sCol1 As String
    Dim sCol2 As String
    
    If (ListView1.ListItems.count = 0) Or (Colonna1 = Colonna2) Then
        SommaColonne = False
        Exit Function
    End If
    
    LockWindowUpdate ListView1.hwnd

    For cnt = 1 To ListView1.ListItems.count
        sCol1 = Trim$(ListView1.ListItems(cnt).SubItems(Colonna1))
        sCol2 = Trim$(ListView1.ListItems(cnt).SubItems(Colonna2))
        ListView1.ListItems(cnt).SubItems(Colonna1) = Trim$(sCol1 & " " & sCol2)
        If CancellaCol2 = True Then ListView1.ListItems(cnt).SubItems(Colonna2) = ""
    Next

    LockWindowUpdate 0&
    
    SommaColonne = True

End Function

Public Function SpostaColonna(ListView1 As ListView, ByVal Colonna1 As Integer, ByVal Colonna2 As Integer) As Boolean
    ' Sposta la Colonna1 nella posizione Colonna2
    Dim cnt As Long
    
    If (ListView1.ListItems.count = 0) Or (Colonna1 = Colonna2) Then
        SpostaColonna = False
        Exit Function
    End If
    
    For cnt = Colonna1 To Colonna2 - 1
        InvertiColonne ListView1, cnt, cnt + 1
    Next
    
    SpostaColonna = True
    
End Function

