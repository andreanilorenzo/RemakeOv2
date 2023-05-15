VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ITEM_MARGIN = 4
Private Const CHECKBOX_WIDTH = 14

' --- required API declarations ---
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Necessaria per la funzione SelectText
Private Declare Function SendMessageByStri Lib "user32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
' Per la funzione AddItemTab
Private Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Const LB_SETTABSTOPS = &H192  'set the tab-stop positions
Private Const LB_ADDSTRING = &H180    'add a string and re-sort a sorted listbox

Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const WM_VSCROLL = &H115
Private Const SB_BOTTOM = 7

' Solo per SelectText
Private Const LB_SELECTSTRING = &H18C

Private Declare Function GetDC Lib "user32" (ByVal hWND As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWND As Long, ByVal hdc As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_SINGLELINE = &H20
Private Const DT_CALCRECT = &H400

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWND As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_VSCROLL = &H200000

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXVSCROLL = 2

'--------------------------------------------------------------------------------
' Per la funzione cerca
Private Declare Function SendMessagebyString Lib "user32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const LB_FINDSTRINGEXACT = &H1A2
'--------------------------------------------------------------------------------

' Private class variables
Private mvarListBox As ListBox
Private m_lMaxItemWidth As Long
Private m_hItemFont As Long
Private m_ListBoxHwnd As Long
Private pDragDropMode As Boolean

Public Function CercaStringaCompleta(ByVal SearchText As String) As Long
    'Input:
    'SearchText: String to Search For
    '
    'Returns: ListIndex of Item if found or -1 if not found
       
    Dim lHwnd As Long
    Dim lMsg As Long
    
    On Error Resume Next
    lHwnd = mvarListBox.hWND
    
    lMsg = LB_FINDSTRINGEXACT
    
    CercaStringaCompleta = SendMessagebyString(lHwnd, lMsg, -1, SearchText)
    
End Function

Public Function CaricaFileTxtArray(FullPatchFileName As String) As String()
    ' Carica il file di testo nella ListBox e restituisce un array contenente le linee contenute nella ListBox
    Dim Lines() As String
    Dim i As Integer
    
    If FileEsisteCls(FullPatchFileName) = True Then
        ' Apro il file, leggo le linee e le inserisco nell'array Lines
        Open FullPatchFileName For Input As #1
            Lines = Split(Input(LOF(1), 1), vbCrLf)
        Close #1
        
        For i = 0 To UBound(Lines)
            ' Questa funzione permette di dividere la linea in più parti ed inserirla nell'array strTemp
            ' strTemp = Split(Lines(i), "*")
            'Aggiungo la linea nel ListBox evitando di inserire le linee vuote
            If Lines(i) <> "" Then mvarListBox.AddItem Lines(i)
        Next
        
        ' Assegno l'array Lines all'array della funzione
        CaricaFileTxtArray = Lines
    Else
            mvarListBox.AddItem " File di log non trovato"
    End If
  
End Function

Public Sub Attach(ByRef pListBox As ListBox)
' Call this sub before you will use other methods of this class
    
    Dim FontInt As IFont

    Set mvarListBox = pListBox
    m_ListBoxHwnd = mvarListBox.hWND
   
    ' Determining the handle of the font used in the specified listbox.
    ' Using the IFont interface we can retreive the handle of the font
    ' used in the specified listbox.
    ' We'll use this handle further when we'll calculate the width of
    ' listbox items
    Set FontInt = pListBox.Font
    m_hItemFont = FontInt.hFont
   
End Sub

Public Sub AddItem(ByRef psItemText As String, Optional ByVal pbScrollBottom As Boolean = True, Optional CancellaDuplicati As Boolean = False, Optional TuttoMinuscolo As Boolean = False)
' The following routine adds a string to a specified list box
' and displays the horizontal scroll bar in it if required
   
   Dim lItemWidth As Long
   
   mvarListBox.AddItem psItemText
   
   ' Calculating the width of the currently added item
   lItemWidth = flGetItemWidth(psItemText)
   
   ' Determining whether we need to display the horizontal scroll bar
   If lItemWidth > m_lMaxItemWidth Then
      m_lMaxItemWidth = lItemWidth
      pSetHorExtent
   End If
   
   ' Scrolling the listbox to be sure that the user see the last item
   If pbScrollBottom Then
      SendMessageByLong m_ListBoxHwnd, WM_VSCROLL, SB_BOTTOM, 0
   End If

   ' Cancello gli elemmenti duplicati
   If CancellaDuplicati = True Then
      KillDuplicati TuttoMinuscolo
   End If
  
End Sub

Public Sub AddItemTab(ByRef psItemText As String, Optional ByVal pbScrollBottom As Boolean = True, Optional CancellaDuplicati As Boolean = False, Optional TuttoMinuscolo As Boolean = False)
    Dim iItems     As Integer
    Dim sItem      As String
    Dim i          As Integer
    Dim l          As Long
    ReDim aTabs(3) As Long
    
    ' Set up a listbox with TAB delimited columns.
    ' Add the desired tabstops to an array.
    
    ' NOTE: tabstops are defined in terms of "dialog units". While the
    '       GetDialogBaseUnits function combined with a simple calculation
    '       can be used to convert between dialog units and pixels, the
    '       easiest way to set tabstops where you want is by trial and error.

    aTabs(0) = 10
    aTabs(1) = 30
    aTabs(2) = 40
    
    'Clear any existing tabs.
    Call SendMessageArray(mvarListBox.hWND, LB_SETTABSTOPS, 0&, 0&)
    'Set the tabs.
    Call SendMessageArray(mvarListBox.hWND, LB_SETTABSTOPS, 3&, aTabs(0))
    
    Call SendMessage(mvarListBox.hWND, LB_ADDSTRING, 0&, psItemText)
    
    ' Scrolling the listbox to be sure that the user see the last item
    If pbScrollBottom Then
       SendMessageByLong m_ListBoxHwnd, WM_VSCROLL, SB_BOTTOM, 0
    End If
    
    ' Cancello gli elemmenti duplicati
    If CancellaDuplicati = True Then
       KillDuplicati TuttoMinuscolo
    End If

End Sub

Public Sub TextSelect(psItemText As String)
' Ottima da inserire nella routine dell'evento Change di un TextBox
' Esempio:
' Private Sub Text1_Change()
'    ListBox.TextSelect (Text1.Text)
' End Sub

    If psItemText <> "" Then
        SendMessageByStri mvarListBox.hWND, LB_SELECTSTRING, -1, psItemText
    End If

End Sub
    
Public Sub KillDuplicati(Optional TuttoMinuscolo As Boolean = False)
' Cancella i valori duplicati nella ListBox

        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
        
    KillDupe = 0
    
    For Search1& = 0 To mvarListBox.ListCount - 1
        For Search2& = Search1& + 1 To mvarListBox.ListCount - 1
            KillDupe = KillDupe + 1
            If TuttoMinuscolo = False Then
                If mvarListBox.List(Search1&) = mvarListBox.List(Search2&) Then
                    mvarListBox.RemoveItem Search2&
                    Search2& = Search2& - 1
                End If
            Else
                If LCase$(mvarListBox.List(Search1&)) = LCase$(mvarListBox.List(Search2&)) Then
                    mvarListBox.RemoveItem Search2&
                    Search2& = Search2& - 1
                End If
            End If
        Next Search2&
    Next Search1&
    
End Sub

Public Property Get List(piIndex As Integer) As String
' Restituisce il valore contenuto in ListBox all'indice piIndex
    
    If piIndex = -1 Then Exit Sub
    List = mvarListBox.List(piIndex)
    
End Property

Public Property Let List(ByRef piIndex As Integer, ByVal piValue As String)
' Scrivo il valore di piValue nell'indice piIndex di ListBox
' Esempio:
' ListBox.List(piIndex) = piValue

    If piIndex <= mvarListBox.ListCount - 1 Then
        mvarListBox.List(piIndex) = piValue
    Else
    End If
    
End Property

Public Property Get ListIndex() As Integer
    ListIndex = mvarListBox.ListIndex
End Property
Public Property Let ListIndex(Index As Integer)
    If mvarListBox.ListIndex >= 0 Then mvarListBox.ListIndex = Index
End Property

Public Property Get FontBold() As Boolean
    FontBold = mvarListBox.FontBold
End Property
Public Property Let FontBold(Valore As Boolean)
    mvarListBox.FontBold = Valore
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = mvarListBox.FontItalic
End Property
Public Property Let FontItalic(Valore As Boolean)
    mvarListBox.FontItalic = Valore
End Property
        
Public Property Get FontName() As String
    FontName = mvarListBox.FontName
End Property
Public Property Let FontName(Valore As String)
    mvarListBox.FontName = Valore
End Property
         
Public Property Get FontSize() As Long
    FontSize = mvarListBox.FontSize
End Property
Public Property Let FontSize(Valore As Long)
    mvarListBox.FontSize = Valore
End Property

Public Property Get FontStrikethru() As Boolean
    FontStrikethru = mvarListBox.FontStrikethru
End Property
Public Property Let FontStrikethru(Valore As Boolean)
    mvarListBox.FontStrikethru = Valore
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = mvarListBox.FontUnderline
End Property
Public Property Let FontUnderline(Valore As Boolean)
    mvarListBox.FontUnderline = Valore
End Property

Public Property Get FontWeight() As Long
    FontWeight = mvarListBox.Font.Weight
End Property
Public Property Let FontWeight(Valore As Long)
    mvarListBox.Font.Weight = Valore
End Property

Public Property Get Width() As Long
    Width = mvarListBox.Width
End Property
Public Property Let Width(Valore As Long)
    mvarListBox.Width = Valore
End Property

Public Property Get Left() As Long
    Left = mvarListBox.Left
End Property
Public Property Let Left(Valore As Long)
    mvarListBox.Left = Valore
End Property

Public Property Get Top() As Long
    Top = mvarListBox.Top
End Property
Public Property Let Top(Valore As Long)
    mvarListBox.Top = Valore
End Property


Public Property Get TopIndex() As Long
    TopIndex = mvarListBox.TopIndex
End Property
Public Property Let TopIndex(Valore As Long)
    mvarListBox.TopIndex = Valore
End Property


Public Property Get DragDropMode() As Boolean
    DragDropMode = pDragDropMode
End Property

Public Property Let DragDropMode(Attiva As Boolean)
    ' Attivo o disattivo la modalità DragDrop
    Select Case Attiva
        Case Is = True
            mvarListBox.OLEDropMode = vbOLEDropManual
            mvarListBox.OLEDragMode = vbOLEDragAutomatic
        Case Is = False
            mvarListBox.OLEDropMode = vbOLEDropNone
            mvarListBox.OLEDragMode = vbOLEDropNone
    End Select
    pDragDropMode = Attiva
End Property

Public Property Get ListCount() As Integer
' Conta le righe della ListBox e ritorna il valore
   
    ListCount = mvarListBox.ListCount
    
End Property

Public Sub Clear()

   mvarListBox.Clear
   m_lMaxItemWidth = 0
   SendMessageByLong m_ListBoxHwnd, LB_SETHORIZONTALEXTENT, 0, 0
   
End Sub

Public Property Let TextChange(ByVal piIndex As Integer, psItemText As String)
' Use this write-only property to change the caption of a list box item
   
   Dim lPrevItemWidth As Long
   Dim lNewItemWidth As Long
   
   lPrevItemWidth = flGetItemWidth(mvarListBox.List(piIndex))
   mvarListBox.List(piIndex) = psItemText
   lNewItemWidth = flGetItemWidth(psItemText)
   
   If lPrevItemWidth = m_lMaxItemWidth Then
      If lNewItemWidth > lPrevItemWidth Then
         m_lMaxItemWidth = lNewItemWidth
         pSetHorExtent
      Else
         RefreshHScroll
      End If
   Else
      If lNewItemWidth > m_lMaxItemWidth Then
         m_lMaxItemWidth = lNewItemWidth
         pSetHorExtent
      End If
   End If
   
End Property

Public Sub Selected(Valore As Long)
    On Error GoTo Selected_Error
    If Valore = -1 Then Exit Sub
    mvarListBox.Selected(Valore) = True
    Exit Sub
Selected_Error:
    Exit Sub
End Sub

Public Sub Move(sinistra As Integer, Optional superiore As Integer, Optional larghezza As Integer, Optional altezza As Integer)
    On Error GoTo Move_Error
    mvarListBox.Move sinistra, superiore, larghezza, altezza
    Exit Sub
Move_Error:
    Exit Sub
End Sub

Public Sub pSetHorExtent()
   SendMessageByLong m_ListBoxHwnd, LB_SETHORIZONTALEXTENT, m_lMaxItemWidth + ITEM_MARGIN + IIf(mvarListBox.Style = vbListBoxCheckbox, CHECKBOX_WIDTH, 0), 0
End Sub

Public Sub RemoveItem(ByVal piIndex As Integer, Optional SelezionaAltro As Boolean = False)
    Dim lItemWidth As Long
   
    If piIndex < 0 Then Exit Sub
   
   lItemWidth = flGetItemWidth(mvarListBox.List(piIndex))
   
   mvarListBox.RemoveItem piIndex
   
   If mvarListBox.ListCount = 0 Then
      m_lMaxItemWidth = 0
      SendMessageByLong m_ListBoxHwnd, LB_SETHORIZONTALEXTENT, 0, 0
      Exit Sub
   End If
   
   If lItemWidth = m_lMaxItemWidth Then
      ' we remove the item with the maximum width
      ' and must recalculate the maximum width and the
      ' horizontal extent of our list box
      RefreshHScroll
   End If
   
   If SelezionaAltro = False Then Exit Sub
   
    Select Case piIndex
        Case Is < mvarListBox.ListCount
            mvarListBox.ListIndex = piIndex
        Case Is = mvarListBox.ListCount
            mvarListBox.ListIndex = piIndex - 1
        Case Is > mvarListBox.ListCount
    End Select
   
End Sub

Public Sub RefreshHScroll()
' Recalculates the maximum width for all items and
' displays the horizontal scroll bar if required
   
   Dim hdc As Long
   Dim rc As RECT
   Dim hOldFont As Long
   Dim i As Integer
   Dim lItemWidth As Long
   
   ' We use the same technique like in the flGetItemWidth function,
   ' but sligtly optimize it: we don't need to select/deselect the
   ' required font for each item - we can do it at once outside the cycle
   hdc = GetDC(m_ListBoxHwnd) ' retrieving HDC for the listbox
   hOldFont = SelectObject(hdc, m_hItemFont) ' selecting the required font
   
   m_lMaxItemWidth = 0
   For i = 0 To mvarListBox.ListCount - 1
      DrawText hdc, mvarListBox.List(i), -1, rc, DT_SINGLELINE + DT_CALCRECT
      lItemWidth = rc.Right - rc.Left
      If lItemWidth > m_lMaxItemWidth Then
         m_lMaxItemWidth = lItemWidth
      End If
   Next
   
   ' restoring the previous font
   Call SelectObject(hdc, hOldFont)
   ReleaseDC m_ListBoxHwnd, hdc
   
   pSetHorExtent
End Sub

Private Function flGetItemWidth(psItemText As String)
   Dim hdc As Long
   Dim rc As RECT
   Dim hOldFont As Long
   
   hdc = GetDC(m_ListBoxHwnd) ' retrieving HDC for the listbox
   hOldFont = SelectObject(hdc, m_hItemFont) ' selecting the required font
   ' if you specify the DT_CALCRECT flag,
   ' DrawText only Determines the width and height of the rectangle
   ' required to display the text:
   DrawText hdc, psItemText, -1, rc, DT_SINGLELINE + DT_CALCRECT
   flGetItemWidth = rc.Right - rc.Left
   ' restoring the previous state
   Call SelectObject(hdc, hOldFont)
   ReleaseDC m_ListBoxHwnd, hdc
   
End Function

Public Property Get MinWidthNoHScroll() As Long
' Returns the minimum width of the list box when the horizontal scroll bar disappears
' NOTE: the return value in pixels
   
   Dim bHasVScrBar As Boolean
   
   bHasVScrBar = GetWindowLong(m_ListBoxHwnd, GWL_STYLE) And WS_VSCROLL
   
   MinWidthNoHScroll = m_lMaxItemWidth + IIf(bHasVScrBar, GetSystemMetrics(SM_CXVSCROLL), 0) + _
      IIf(mvarListBox.Style = vbListBoxCheckbox, CHECKBOX_WIDTH, 0) + ITEM_MARGIN + 4

End Property

Private Function FileEsisteCls(TheFileName As String) As Boolean
    ' Guarda se il file indicato nella stringa passata esiste oppure no
    ' e restituisce True o False
    On Error GoTo FileEsisteCls_Err
    
    If Len(TheFileName$) = 0 Then
        FileEsisteCls = False
        Exit Function
    End If
    
    If Len(Dir$(TheFileName$)) Then
        FileEsisteCls = True
    Else
        FileEsisteCls = False
    End If
    
FileEsisteCls_Err:

End Function

Public Function SommaListBox(sForm As String, sCtrl As String, iColumn As Integer) As Variant
    'Si richiama cosi: x = SumListbox([Name],[ListBox].[Name],1)
    Dim frm As Form
    Dim ctrl As Control
    Dim i As Integer
    Dim vSum As Variant

    On Error Resume Next
    Set frm = Forms(sForm)

    If Err <> 0 Then
        SommaListBox = "ERR!FormNotFound"
        Exit Function
    End If
    
    On Error Resume Next
    Set ctrl = frm(sCtrl)
    If Err <> 0 Then
        SommaListBox = "ERR!ListboxNotFound"
        Exit Function
    End If
    
    If iColumn > ctrl.ColumnCount Then
        SommaListBox = "ERR!ColumnNotFound"
        Exit Function
    End If

    vSum = 0
    
    For i = 0 To ctrl.ListCount - 1
        vSum = vSum + ctrl.Column(iColumn - 1, i)
    Next i
    
    SommaListBox = vSum

End Function

