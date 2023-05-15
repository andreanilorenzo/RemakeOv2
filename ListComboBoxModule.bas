Attribute VB_Name = "ListComboBoxModule"
Option Explicit

'-------------------------------------------------------------------------------------
' Dichiaro la funzione BloccaComboBox
Private Const EM_SETREADONLY = &HCF
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'-------------------------------------------------------------------------------------

Public Sub SortListBox(oLB As ListBox)
    'Forces a resort of a (Sort=True) ListBox
    Dim aItem() As String
    Dim aData() As Long
    Dim I As Integer, iUB As Integer

    With oLB
        iUB = .ListCount - 1
        If iUB >= 0 Then 'not empty
            ReDim aItem(0 To iUB)
            ReDim aData(0 To iUB)
            For I = 0 To iUB
                aItem(I) = .List(I)
                aData(I) = .ItemData(I)
            Next
            .Clear
            For I = 0 To iUB
                .AddItem aItem(I)
                .ItemData(.NewIndex) = aData(I)
            Next
        End If
    End With
End Sub

Public Sub BloccaComboBox(Combo1 As ComboBox, Optional Blocca As Boolean = True, Optional Ricorsiva As Boolean = False, Optional Testo As String = "")
    ' Imposta il ComboBox come non editabile
    '
    ' Blocca:       Imposta la modalità Bloccato/Sbloccato
    ' Testo:        Se presente scrive il testo nella ComboBox
    ' Ricorsiva:    Ad ogni chiamata cambia lo stato Bloccato/Sbloccato
    '               Il parametro Ricorsiva ha la precedenza sul parametro Blocca
    '
    Dim hwndEdit As Long
    Static State As Boolean

    ' Get the handle to the edit portion of the combo control
    hwndEdit = FindWindowEx(Combo1.hwnd, 0&, vbNullString, vbNullString)
       
    If Ricorsiva = True Then
        State = Not State
    Else
        State = Blocca
    End If
    
    If hwndEdit <> 0 Then
        ' Scrivo il testo.... funziona solo se: (style 0)
        If Testo <> "" Then Call SetWindowText(hwndEdit, Testo)
        ' Disable the edit control's editing ability
        Call SendMessage(hwndEdit, EM_SETREADONLY, CLng(Abs(State)), ByVal 0&)
    End If
   
End Sub

Public Sub LoadComboBox(Directory As String, TheList As ComboBox, Optional CancellaPrima As Boolean = True)
    ' Legge i dati da file di testo e li scrive in una ComboBox
    Dim MyString As String
    Dim File1 As Integer
    On Error Resume Next
    
    If FileExists(Directory) = False Then Exit Sub
    
    If CancellaPrima = True Then TheList.Clear
    
    File1 = FreeFile
    
    Open Directory For Input Access Read As #File1
    Do While EOF(File1) = False
        Line Input #File1, MyString
        'DoEvents
        TheList.AddItem MyString
    Loop
    Close #File1
    
End Sub

Public Sub SaveComboBox(ByVal Directory As String, ByVal TheList As ComboBox)
    ' Legge i dati da una ComboBox e li scrive su di un file di testo
    Dim savelist As Long
    Dim File2 As Integer
    On Error Resume Next
        
    If TheList.ListCount = 0 Then Exit Sub

    File2 = FreeFile

    Open Directory$ For Output As #File2
    
    For savelist& = 0 To TheList.ListCount - 1
        Print #File2, TheList.List(savelist&)
    Next
    
    Close #File2
    
End Sub

Public Function CercaInComboBox(ByVal Combo1 As ComboBox, ByRef Testo As String) As Integer
    ' Cerca del testo nella ComboBox,
    ' seleziona l'ultima riga che contiene il testo cercato
    ' e restituisce quante volte è stato trovato il testo
    Dim cnt As Integer
    Dim cnt1 As Integer
    
    For cnt = 0 To Combo1.ListCount - 1
        If InStr(1, Combo1.List(cnt), Testo, vbTextCompare) <> 0 Then
            Combo1.ListIndex = cnt
            cnt1 = cnt1 + 1
        End If
    Next
    
    CercaInComboBox = cnt1
    
End Function

