VERSION 5.00
Begin VB.Form frmLogFile 
   Caption         =   "Log File"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   Icon            =   "frmLogFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   7395
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optVisualizza 
      Caption         =   "Note"
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton optVisualizza 
      Caption         =   "Creati"
      Height          =   195
      Index           =   6
      Left            =   6240
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton optVisualizza 
      Caption         =   "Intermedi"
      Height          =   195
      Index           =   5
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optVisualizza 
      Caption         =   "Cancellati"
      Height          =   195
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton optVisualizza 
      Caption         =   "Copiati"
      Height          =   195
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton optVisualizza 
      Caption         =   "Errori"
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton optVisualizza 
      Caption         =   "Tutto"
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5310
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7335
   End
   Begin VB.Label lblTotInElenco 
      Alignment       =   2  'Center
      Caption         =   "lblTotInElenco"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "Visualizza:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LBHS As clsListBox
Dim CarattereFiltro As String
Dim ArrayElencoListBox() As String

Private Sub Form_Load()
    
    ' Inizializzo la classe per la ListBox
    Set LBHS = New clsListBox
    LBHS.Attach List1
    
    MousePointer = 11
       
    If ContoErrori = 0 Then
        optVisualizza(6) = True
        CarattereFiltro = "="
    Else
        optVisualizza(1) = True
        CarattereFiltro = "*"
    End If

    LBHS.Clear
    ArrayElencoListBox = LBHS.CaricaFileTxtArray(App.path & "\Remakeov2.log")
        
    lblTotInElenco.Caption = ".............................................in elenco ci sono " & LBHS.ListCount & " files............................................."
    Call FiltraListBox
    
    frmRemakeov2.Enabled = False

    MousePointer = 0

End Sub

Private Sub Form_Resize()
    Dim Sin As Integer 'Il borbo da lasciare a sinistra
    Dim Inf As Integer 'Il bordo da lasciare sotto
    
    On Error Resume Next
    
    Sin = 0
    Inf = 0
    LBHS.Move LBHS.Left, LBHS.Top, Me.ScaleWidth - LBHS.Left - Sin, Me.ScaleHeight - LBHS.Top - Inf

End Sub

Private Sub FiltraListBox(Optional CancellaListBox As Boolean = True)
    Dim cnt As Integer
    Dim cntCarr As Integer
    Dim arrayFiltro() As String
    Dim lngTmp As Long
    
    'If ArrayElencoListBox <= 1 Then Exit Sub
    
    lblTotInElenco.Caption = "........................................................................................................................................"
    DoEvents

    ' Cancello la ListBox
    If CancellaListBox = True Then LBHS.Clear
    
    If CarattereFiltro <> "" Then
        arrayFiltro = Split(CarattereFiltro, ";")
        ' Carico i dati dell'array nella ListBox
        For cnt = 0 To UBound(ArrayElencoListBox)
            For cntCarr = 0 To UBound(arrayFiltro)
                ' Cerco il carattere nella strina contenuta nell'array (se l'esito è negativo la funzione resituisce 0)
                lngTmp = InStr(1, ArrayElencoListBox(cnt), arrayFiltro(cntCarr), vbTextCompare)
                'Inserisco solo le linee che contengono il carattere di filtro
                If lngTmp <> 0 Then LBHS.AddItem (ArrayElencoListBox(cnt))
            Next
        Next
    Else
        For cnt = 0 To UBound(ArrayElencoListBox)
            LBHS.AddItem (ArrayElencoListBox(cnt))
        Next
    End If
    
    ' Aggiorno lo scroll orrizzontale
    LBHS.RefreshHScroll
    lblTotInElenco.Caption = ".............................................in elenco ci sono " & LBHS.ListCount & " files............................................."
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmRemakeov2.Enabled = True

End Sub

Private Sub optVisualizza_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MousePointer = 11
    
    If optVisualizza(0) = True Then CarattereFiltro = ""
    If optVisualizza(1) = True Then CarattereFiltro = "*"
    If optVisualizza(2) = True Then CarattereFiltro = "°"
    If optVisualizza(3) = True Then CarattereFiltro = ">" & ";<"
    If optVisualizza(4) = True Then CarattereFiltro = "-"
    If optVisualizza(5) = True Then CarattereFiltro = "+"
    If optVisualizza(6) = True Then CarattereFiltro = "="
    
    Call FiltraListBox
    
    MousePointer = 0
    
End Sub
