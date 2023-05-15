VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmUpdateFile 
   Caption         =   "Aggiornamento File"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   Icon            =   "frmUpdateFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   10380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picComandi 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1080
      ScaleHeight     =   735
      ScaleWidth      =   8355
      TabIndex        =   3
      Top             =   4920
      Width           =   8355
      Begin VB.CommandButton cmdCercaNuovi 
         Caption         =   "Cerca &nuovi file"
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         ToolTipText     =   "Seleziona i file che sono presenti sul sito ma non sono presenti nella cartella del computer"
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cmbTipoPDI 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Text            =   "cmbTipoPDI"
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmdDownload 
         Caption         =   "Download file"
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox cekRemake 
         Caption         =   "cekRemake"
         Height          =   195
         Left            =   5280
         TabIndex        =   7
         ToolTipText     =   "Dopo aver scaricato i file apri automaticamente la finestra per effettuare il trattamento"
         Top             =   360
         Value           =   1  'Checked
         Width           =   135
      End
      Begin VB.CommandButton cmdEsci 
         Cancel          =   -1  'True
         Caption         =   "&Esci  [Esc]"
         Height          =   375
         Left            =   6480
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdControlla 
         Caption         =   "&Controlla file"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdAggiorna 
         Caption         =   "&Aggiorna file"
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComctlLib.ImageList imglRunSearch 
      Left            =   1800
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   61
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":0EFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":36B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":53BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":56D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":582E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":5988
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":5AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":5DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":6C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":7AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":894A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":979C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":BF4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":C268
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":D0BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":D3D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":EA2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":ED48
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":F062
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":F37C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":F696
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":FF70
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":1084A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":1169C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":124EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":13340
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":13C1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":14A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":16776
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":16A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":178E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":1A094
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":1C846
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":1E550
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":20D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":22A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":232E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":24138
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":25792
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2606C
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":26946
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":27798
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":28DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":29C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":29F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2A838
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2AB52
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2AE6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2B186
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2B2E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2B5FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2B914
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2BC2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2BF48
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2C262
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2C57C
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2C6D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2C830
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":2C98A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   1080
      ScaleHeight     =   990
      ScaleWidth      =   2295
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   2325
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   5636
      View            =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblInfo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblHelp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmUpdateFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tyFile
    Nome As String
    Data As String
    trovato As String
    DataNuovo As String
    
    Gruppo As String
    Formato As String
    Mappa As String
    Immagine As String
    Note As String
    Modalità As String
End Type

Private siEsce As Boolean
Private ListaFile() As tyFile
Private FileTrovati As Integer

Private Sub cmdAggiorna_Click()
    Dim cnt As Long
    Dim Indice As Variant
    
    Call SetCeckedListView(frmDownload.ListView1)
    
    If frmDownload.ListView1.ListItems.Count = 0 Then Exit Sub
    
    For cnt = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(cnt).Checked = True Then
            Indice = Trim(GetValoreCella(ListView1, cnt, 6))
            If Indice > 0 And Indice <> "" Then
                frmDownload.ListView1.ListItems.Item(CLng(Indice)).Checked = True
            End If
        End If
    Next

    Call ControllaCheck(frmDownload.ListView1)
    
    frmDownload.Visible = True
    frmDownload.elencoDa = xml
    frmDownload.cekRemake.Value = cekRemake.Value
    frmDownload.cmdScaricaFile_Click (0)
    Unload Me

End Sub

Private Sub cmdCercaNuovi_Click()
    Dim cnt As Long
    Dim cnt1 As Long
    Dim bNuovo As Boolean
    Dim NuoviFile As Integer
    
    lblInfo.Caption = "Ricerca dei nuovi file in corso......."
    lblInfo.Refresh
    cmdDownload.Visible = False
    frmDownload.cekFiltraPerMappa.Value = 1
    frmDownload.cmdScaricaElenco_Click (0)
    Call SetCeckedListView(ListView1)
    NuoviFile = 0
    frmDownload.cmbTipoPDI.ListIndex = cmbTipoPDI.ListIndex
    
    For cnt = 1 To frmDownload.ListView1.ListItems.Count
        bNuovo = True
        
        For cnt1 = 1 To ListView1.ListItems.Count
            ' Cerco i nuovi file
            If LCase$(Replace$(frmDownload.ListView1.ListItems.Item(cnt).ListSubItems.Item(1).Text, " ", "_")) = LCase$(Replace$(ListView1.ListItems.Item(cnt1).ListSubItems.Item(1).Text, " ", "_")) Then
                ' Ho trovato il file
                bNuovo = False
                Exit For
            End If
        Next
        
        If bNuovo = True Then
            frmDownload.ListView1.ListItems.Item(cnt).Checked = True
        Else
            frmDownload.ListView1.ListItems.Item(cnt).Checked = False
        End If
        
    Next

    Call ControllaCheck(frmDownload.ListView1)
    Unload Me
    frmDownload.SetFocus
    
End Sub

Private Sub cmdControlla_Click()
    Dim cnt As Long
    Dim cnt1 As Long
    Dim bTrovato As Boolean
    Dim DataMio As Date
    Dim DataWeb As Date
    Dim NuoviFile As Integer
    
    lblInfo.Caption = "Ricerca dei file aggiornati in corso......."
    lblInfo.Refresh
    cmdDownload.Visible = False
    frmDownload.cekFiltraPerMappa.Value = 0
    frmDownload.cmdScaricaElenco_Click (0)
    Call SetCeckedListView(ListView1)
    NuoviFile = 0
    
    For cnt = 1 To ListView1.ListItems.Count
        bTrovato = False
        
        For cnt1 = 1 To frmDownload.ListView1.ListItems.Count
            ' Cerco i file aggiornati
            If LCase$(Replace$(frmDownload.ListView1.ListItems.Item(cnt1).ListSubItems.Item(1).Text, " ", "_")) = LCase$(Replace$(ListView1.ListItems.Item(cnt).ListSubItems.Item(1).Text, " ", "_")) Then
                
                DataMio = CDate(ListView1.ListItems.Item(cnt).ListSubItems.Item(2).Text)
                DataWeb = CDate(frmDownload.ListView1.ListItems.Item(cnt1).ListSubItems.Item(2).Text)
                
                If DataMio < DataWeb Then
                    ListView1.ListItems.Item(cnt).Checked = True
                    Call ScriviCella(ListView1, cnt, GetNumColDaIntestazione(ListView1, "Nuovo"), Replace$(CStr(DataWeb), "/", "-"))
                    NuoviFile = NuoviFile + 1
                Else
                    ListView1.ListItems.Item(cnt).Checked = False
                    Call ScriviCella(ListView1, cnt, GetNumColDaIntestazione(ListView1, "Nuovo"), " - - - ")
                End If
                
                bTrovato = True
                Exit For
            
            End If
        Next
        
        If bTrovato = True Then
            Call ScriviCella(ListView1, cnt, GetNumColDaIntestazione(ListView1, "Indice"), frmDownload.ListView1.ListItems.Item(cnt1).index)
            Call ScriviCella(ListView1, cnt, GetNumColDaIntestazione(ListView1, "Trovato"), "Si")
            Call ColorListViewRow(ListView1, cnt, vbBlack)
        Else
            Call ScriviCella(ListView1, cnt, GetNumColDaIntestazione(ListView1, "Trovato"), "No")
            Call ColorListViewRow(ListView1, cnt, vbRed)
        End If
        
    Next
    
    If NuoviFile > 0 Then
        Call ControllaCheck(ListView1)
        lblInfo.Caption = "Sono stati trovati " & NuoviFile & " file aggiornati." & vbNewLine & "Per attivare la procedura download automatico premi il tasto ""Aggiorna File""."
        cmdAggiorna.Enabled = True
        cekRemake.Enabled = True
        cmdDownload.Move cmdEsci.Left, cmdEsci.Top, cmdEsci.Width, cmdEsci.Height
        cmdDownload.Visible = False
        cmdAggiorna.SetFocus
    Else
        lblInfo.Caption = "Non sono stati trovati dei file aggiornati." & vbNewLine & "Premi il tasto ""Download File"" per scaricare altri file, opprue ""Esci"" per tornare al menu principale."
        cmdAggiorna.Enabled = False
        cekRemake.Enabled = False
        cmdDownload.ZOrder
        cmdDownload.Move cmdAggiorna.Left, cmdAggiorna.Top, cmdAggiorna.Width, cmdAggiorna.Height
        cmdDownload.Visible = True
    End If
    
End Sub

Private Sub cmdDownload_Click()
    frmDownload.Visible = True
    Unload Me
End Sub

Private Sub cmdEsci_Click()
    siEsce = True
    Unload Me
End Sub

Private Sub Form_Activate()
    frmMain.Visible = False
End Sub

Private Sub Form_Initialize()
    frmDownload.Visible = False
End Sub

Private Sub Form_Load()
    Dim LarghezzaColonne As Variant
    Dim ColonneListView1 As Variant
    Dim TagColonne As Variant
    Dim strTmp As String
    Dim cnt As Long
    
    ' Creo gli array con i dati delle colonne
    LarghezzaColonne = Array(1200, 3000, 1300, 3000, 1000, 600, 1000)
    ColonneListView1 = Array("", "Nome", "Data", "Nuovo", "Trovato", "Note", "Indice")
    ' Serve per la funzione di ordinamento dei dati nella colonna
    TagColonne = Array("NUMBER", "STRING", "DATE", "DATE", "STRING", "STRING", "NUMBER")
    
    With ListView1
        .HideSelection = False
        .FullRowSelect = True
        .MultiSelect = False
        .View = lvwReport
        .LabelEdit = lvwManual ' Evita che si possa editare la prima colonna
        .Icons = imglRunSearch
        .SmallIcons = imglRunSearch
        .ColumnHeaderIcons = imglRunSearch
        For cnt = 0 To UBound(ColonneListView1)
            .ColumnHeaders.Add , "x" & cnt + 1, " " & ColonneListView1(cnt)
            .ColumnHeaders.Item(cnt + 1).Width = LarghezzaColonne(cnt)
            .ColumnHeaders.Item(cnt + 1).Tag = TagColonne(cnt)
        Next
    End With
    Call SetListViewColor(ListView1, Picture1, 1, vbWhite, vbGreenLemon)
    Call AutoSizeUltimaColonna(ListView1)
    
    strTmp = "Da questa finestra si possono aggiornare i file dei POI presenti nella cartella: "
    strTmp = strTmp & vbNewLine & Var(PoiScaricati).Valore
    lblHelp.Caption = strTmp
    
    FileTrovati = 0
    
    Call ElencoFile
    If FileTrovati > 0 Then
        lblInfo.Caption = "Premi il tasto ""Controlla File"" per controllare se ci sono dei file più aggiornati nell'elenco XML di poigps. Ricordati che devi essere registrato e loggato correttamente sul sito www.poigps.com. I file vengono controllati e scaricati in modalità XML, cioè utilizzando l'elenco dei file presente all'indirizzo http://www.poigps.com/poi.xml"
        Call ElencoFileInListView
        Call ControllaCheck(ListView1)
    Else
        cmdControlla.Enabled = False
        lblInfo.Caption = "Non sono stati trovati file nella cartella: " & vbNewLine & Var(PoiScaricati).Valore & vbNewLine & "Premi il tasto ""Download File"" per scaricare i file."
        cmdDownload.ZOrder
        cmdDownload.Move cmdEsci.Left, cmdEsci.Top, cmdEsci.Width, cmdEsci.Height
        cmdDownload.Visible = True
    End If
    
    If frmDownload.cmbTipoPDI.ListCount > 0 Then
        For cnt = 0 To frmDownload.cmbTipoPDI.ListCount - 1
            cmbTipoPDI.AddItem frmDownload.cmbTipoPDI.List(cnt)
        Next
        cmbTipoPDI.ListIndex = 0
        BloccaComboBox cmbTipoPDI
    End If
    
    siEsce = False
    cmdAggiorna.Enabled = False
    cekRemake.Enabled = False
    
End Sub

Private Sub ElencoFile()
    Dim cnt As Long
    Dim StmpFile As String

    cnt = 0
    StmpFile = Dir(Var(PoiScaricati).Valore & "\*.ov2")
    
    While StmpFile <> ""
        ReDim Preserve ListaFile(cnt)
        ListaFile(cnt).Nome = Left$(StmpFile, Len(StmpFile) - 4)
        ListaFile(cnt).Data = GetDataFile(Var(PoiScaricati).Valore & "\" & StmpFile, "M")
        
        StmpFile = Dir()
        cnt = cnt + 1
        FileTrovati = cnt
    Wend
    
End Sub

Private Sub ElencoFileInListView()
    Dim cntRecord As Long
    Dim cnt As Long
    Dim itmX As Variant

    ListView1.ListItems.Clear
        
    cntRecord = 0
    For cnt = 0 To UBound(ListaFile) ' Scorro tutte le righe dell'array
        Set itmX = ListView1.ListItems.Add(, , Format(cntRecord + 1, "00000"))
        itmX.SubItems(GetNumColDaIntestazione(ListView1, "Nome") - 1) = ListaFile(cnt).Nome
        itmX.SubItems(GetNumColDaIntestazione(ListView1, "Data") - 1) = ListaFile(cnt).Data
        itmX.SubItems(GetNumColDaIntestazione(ListView1, "Nuovo") - 1) = " "
        itmX.SubItems(GetNumColDaIntestazione(ListView1, "Trovato") - 1) = " "
        itmX.SubItems(GetNumColDaIntestazione(ListView1, "Note") - 1) = " "
        itmX.SubItems(GetNumColDaIntestazione(ListView1, "Indice") - 1) = " "
        cntRecord = cntRecord + 1
    Next
    
    TotaleRigheListView ListView1
    
    ListView1.Refresh

End Sub

Private Sub cmbTipoPDI_Click()
    frmDownload.cmbTipoPDI.ListIndex = cmbTipoPDI.ListIndex
End Sub

Private Sub Form_Resize()
   ' Width = Larghezza  Height = Altezza
    On Error Resume Next

    If Me.Width < 10500 Then Me.Width = 10500
    If Me.Height < 6200 Then Me.Height = 6200

    With lblHelp
        .Move 0, .Top, Me.ScaleWidth, .Height
    End With

    With picComandi
        .Move (Me.ScaleWidth - .Width) / 2, Me.ScaleHeight - .Height, .Width, .Height
    End With
    
    With lblInfo
        .Move 0, picComandi.Top - .Height, Me.ScaleWidth, .Height
    End With

    With ListView1
        .Move 0, lblHelp.Top + lblHelp.Height, Me.ScaleWidth, lblInfo.Top - (lblHelp.Top + lblHelp.Height)
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If siEsce = True Then
        Unload frmDownload
        DoEvents
    Else
        frmDownload.Visible = True
    End If
    
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ControllaRigaChecked(ListView1, Item)
End Sub
