VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmPrendiRegPr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aggiorna elenco"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   Icon            =   "frmPrendiRegPr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox cekSoloSigla 
      BackColor       =   &H80000016&
      Caption         =   "Solo sigla PR"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.ComboBox cmbRegione 
      Height          =   315
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   6
      Text            =   "cmbRegione"
      ToolTipText     =   "ATTENZIONE: funzione non ancora completata"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ComboBox cmbProvincia 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "cmbProvincia"
      ToolTipText     =   "ATTENZIONE: funzione non ancora completata"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "&Annulla  [Esc]"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdProvince 
      Caption         =   "Scarica Province"
      Height          =   1215
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdRegioni 
      Caption         =   "Scarica Regioni"
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "&Salva ed esci"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3540
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   5805
      ExtentX         =   10239
      ExtentY         =   6244
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmPrendiRegPr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Regioni As String
Private Province As String
Private Cosa As String
Private DocW As Object

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Regioni = "http://www.comuni-italiani.it/regioni.html"
    Province = "http://www.comuni-italiani.it/province.html"

    cmbRegione.Text = "Regione"
    cmbProvincia.Text = "Provincia"
    Call BloccaComboBox(cmbRegione)
    Call BloccaComboBox(cmbProvincia)
    
    cmdSalva.Enabled = False
    
End Sub

Private Sub cmdProvince_Click()
    Cosa = "PR"
    cmdSalva.Enabled = False
    WebBrowser1.Navigate2 Province
End Sub

Private Sub cmdRegioni_Click()
    Cosa = "RE"
    cmdSalva.Enabled = False
    WebBrowser1.Navigate2 Regioni
End Sub

Private Sub WebBrowser1_DownloadComplete()
    
    Set DocW = WebBrowser1.Document
    
    Call ImportaDatiPaesi
    
End Sub

Private Sub ImportaDatiPaesi()
    Dim TheSource As String
    Dim arrScratch() As String
    Dim arrDati() As String
    Dim strTmp As String
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim I As Integer
    Dim cnt As Integer
    
    ReDim arrScratch(0)
    ReDim arrDati(0)
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    cmdSalva.Enabled = False
    
    If WebBrowser1.LocationURL <> "" And WebBrowser1.LocationURL <> "http:///" And WebBrowser1.LocationURL <> "about:blank" Then
        TheSource = WebBrowser1.Document.body.outerText
        
        If TheSource <> "" Then
            arrScratch = Split(TheSource, vbNewLine)
            
            cnt = 0
            For I = 0 To UBound(arrScratch)
                ' Prendo solo le righe che hanno come primo valore un numero
                If IsNumeric(Left$(Trim$(arrScratch(I)), 1)) Then
                    arrScratch(cnt) = Mid$(arrScratch(I), 2, Len(arrScratch(I)) - 1)
                    ' Tolgo il numero iniziale
                    If IsNumeric(Left$(Trim$(arrScratch(cnt)), 1)) Then
                        arrScratch(cnt) = Mid$(arrScratch(cnt), 2, Len(arrScratch(cnt)) - 1)
                        ' Tolgo il numero iniziale
                        If IsNumeric(Left$(Trim$(arrScratch(cnt)), 1)) Then
                            arrScratch(cnt) = Mid$(arrScratch(cnt), 2, Len(arrScratch(cnt)) - 1)
                        End If
                    End If
                   cnt = cnt + 1
                End If
            Next
            
            ' Accorcio l'array
            ReDim Preserve arrScratch(cnt - 1)
            ' Preparo il nuovo array
            ReDim arrDati(cnt - 1, 1)
            
            ' Scrivo i dati nel nuovo array
            For I = 0 To UBound(arrDati)
                arrDati(I, 0) = Left$(arrScratch(I), CercaPrimaPosNumero(arrScratch(I)))
                If Cosa = "PR" Then arrDati(I, 1) = Right$(arrScratch(I), 2)
            Next
            
            If Cosa = "RE" Then
                cmbRegione.Clear
                For I = 0 To UBound(arrDati)
                    'Debug.Print arrDati(i, 0) & "   " & arrDati(i, 1)
                    If arrDati(I, 0) <> "" Then cmbRegione.AddItem arrDati(I, 0)
                Next
            End If
            
            If Cosa = "PR" Then
                cmbProvincia.Clear
                For I = 0 To UBound(arrDati)
                   'Debug.Print arrDati(i, 0) & "   " & arrDati(i, 1)
                    If arrDati(I, 1) <> "" Then
                        If cekSoloSigla.Value = 0 Then
                            cmbProvincia.AddItem arrDati(I, 1) & " " & arrDati(I, 0)
                        Else
                            cmbProvincia.AddItem arrDati(I, 1)
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    If cmbRegione.ListCount > 0 And Cosa = "RE" Then
        cmbRegione.Text = "+ Regione"
        lblInfo.Caption = "Elenco delle Regioni aggiornato con successo!"
        cmdSalva.Enabled = True
    
    ElseIf cmbProvincia.ListCount > 0 And Cosa = "PR" Then
        cmbProvincia.Text = "+ Provincia"
        lblInfo.Caption = "Elenco delle Province aggiornato con successo!"
        cmdSalva.Enabled = True
    End If
    
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione ImportaDati."
    
End Sub

Private Function CercaPrimaPosNumero(ByVal Stringa As String) As Integer
    Dim cnt As Integer
    
    For cnt = 1 To Len(Stringa)
        If IsNumeric(Mid$(Stringa, cnt, 1)) Then
            CercaPrimaPosNumero = cnt - 1
            Exit Function
        End If
    Next
    
    CercaPrimaPosNumero = 0
    
End Function

Private Sub cmdSalva_Click()
    
    Call SaveComboBox(Var(RegioniCsv).Valore, cmbRegione)
    Call SaveComboBox(Var(ProvinceCsv).Valore, cmbProvincia)

    frmWeb.cmbRegione.LoadFile Var(RegioniCsv).Valore
    frmWeb.cmbProvincia.LoadFile Var(ProvinceCsv).Valore
    
    frmWeb.ckRegione.Value = 0
    frmWeb.ckProvincia.Value = 0
    frmWeb.cmbRegione.Text = "+ Regione"
    frmWeb.cmbProvincia.Text = "+ Provincia"

    Unload Me
    
End Sub

