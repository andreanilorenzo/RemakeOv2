VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{191731CB-C038-4B30-B988-C682B9E20321}#1.0#0"; "gptabxp.ocx"
Begin VB.Form frmImpostazioni 
   Caption         =   "Impostazioni"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   Icon            =   "frmImpostazioni.frx":0000
   LinkTopic       =   "frmImpostazioni"
   ScaleHeight     =   6975
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "&Esci  [Esc]"
      Height          =   320
      Left            =   8880
      TabIndex        =   1
      ToolTipText     =   " Esce dal programma "
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picImpo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Index           =   1
      Left            =   1560
      ScaleHeight     =   6105
      ScaleWidth      =   10065
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2520
      Width           =   10095
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
         Left            =   1200
         ScaleHeight     =   990
         ScaleWidth      =   2295
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   2325
      End
      Begin MSComctlLib.ListView ListViewMtr 
         Height          =   1995
         Left            =   480
         TabIndex        =   9
         Top             =   600
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   3519
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
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
   End
   Begin VB.PictureBox picImpo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Index           =   2
      Left            =   240
      ScaleHeight     =   6105
      ScaleWidth      =   10065
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   10095
      Begin VB.PictureBox Picture2 
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
         Left            =   1440
         ScaleHeight     =   990
         ScaleWidth      =   2295
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   2325
      End
   End
   Begin VB.PictureBox picImpo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Index           =   3
      Left            =   3600
      ScaleHeight     =   6105
      ScaleWidth      =   10065
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3240
      Width           =   10095
   End
   Begin MSComctlLib.ImageList imglRunSearch 
      Left            =   7320
      Top             =   120
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
            Picture         =   "frmImpostazioni.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":30F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":4DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":5114
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":526E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":53C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":5522
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":583C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":668E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":74E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":838A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":91DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":B98E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":BCA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":CAFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":CE14
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":E46E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":E788
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":EAA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":EDBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":F0D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":F9B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":1028A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":110DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":11F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":12D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":1365A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":144AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":161B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":164D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":17322
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":19AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":1C286
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":1DF90
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":20742
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2244C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":22D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":23B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":251D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":25AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":26386
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":271D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":28832
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":29684
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2999E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2A278
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2A592
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2A8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2ABC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2AD20
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2B03A
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2B354
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2B66E
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2B988
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2BCA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2BFBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2C116
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2C270
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpostazioni.frx":2C3CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   245
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "lblInfo"
      Top             =   480
      Width           =   10280
   End
   Begin GpTabXP.GpTabStrip GpTabStrip1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11880
      BackColor       =   16777215
      ForeColor       =   0
      Style           =   1
      TabColor        =   -2147483643
      TabStyle        =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picImpo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   0
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   2535
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picImpo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   4
      Left            =   3120
      ScaleHeight     =   3105
      ScaleWidth      =   6705
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3120
      Width           =   6735
      Begin SHDocVwCtl.WebBrowser WebXml 
         Height          =   2175
         Left            =   1560
         TabIndex        =   11
         Top             =   600
         Width           =   4575
         ExtentX         =   8070
         ExtentY         =   3836
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
   End
End
Attribute VB_Name = "frmImpostazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrRow As Long

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If FormIsLoad("frmMain") = True Then frmMain.Visible = False
    ListViewMtr.SetFocus
End Sub

Private Sub Form_Load()
    Dim cnt As Long
    
    For cnt = 1 To picImpo.UBound
        picImpo(cnt).BackColor = GpTabStrip1.BackColor
    Next
    picImpo(1).ZOrder

    'ListViewMtr.ToolTipText = "Doppio click sulla riga per modificare il valore (non tutte le righe sono modificabili)"
    
    With GpTabStrip1
         .Tabs.Add , , " Generali       "
         .Tabs.Add , , "                "
         .Tabs.Add , , "                "
         .Tabs.Add , , " File Xml       "
    End With

    With lblInfo
        .Enabled = True
        .Text = "ATENZIONE: la modifica di questi valori comporterà la modifica delle impostazioni del programma."
        .BackColor = GpTabStrip1.BackColor
        .BackColor = vbYellow
        .ForeColor = vbBlack
        .ZOrder
    End With

    PreparaListView ListViewMtr
    
    cmdEsci.ZOrder
    
    ' Carico il file xml nella ListViewMtr
    Matrice2ListView ListViewMtr
    TotaleRigheListView ListViewMtr

    ' Carico il file xml nel webbrowser
    WebXml.Navigate (XmlFileConfig)
        
End Sub

Private Sub Matrice2ListView(ListView1 As ListView)
    Dim cnt As Integer
    Dim itmX As Variant

    For cnt = 1 To UBound(mtrVarXml)
        Set itmX = ListView1.ListItems.Add(, , Format(cnt, "000"))
        itmX.SubItems(1) = mtrVarXml(cnt).Sezione
        itmX.SubItems(2) = mtrVarXml(cnt).SubSezione
        itmX.SubItems(3) = mtrVarXml(cnt).Nome
        itmX.SubItems(4) = CStr(mtrVarXml(cnt).Valore)
        itmX.SubItems(5) = mtrVarXml(cnt).Opzioni
        itmX.SubItems(6) = mtrVarXml(cnt).Predefinito
        itmX.SubItems(7) = mtrVarXml(cnt).Descrizione
        itmX.SubItems(8) = cnt
        'ListView1.ListItems(cnt).ToolTipText = mtrVarXml(cnt).Descrizione
    Next
        
End Sub

Private Sub PreparaListView(ListView1 As ListView)
    Dim LarghezzaColonne As Variant
    Dim ColonneListView1 As Variant
    Dim TagColonne As Variant
    Dim cnt As Long
    
    ' Creo gli array con i dati delle colonne
    LarghezzaColonne = Array(1200, 1500, 1800, 1900, 4000, 2, 2, 10, 500)
    ColonneListView1 = Array("", "Sezione", "SubSezione", "Chiave", "Valore", "Opzioni", "Predefinito", "Descrizione", "Indice")
        ' Serve per la funzione di ordinamento dei dati nella colonna
          TagColonne = Array("NUMBER", "STRING", "STRING", "STRING", "STRING", "STRING", "STRING", "STRING", "NUMBER")
    
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

End Sub

Private Sub Form_Resize()
    Dim t As Integer
    Dim B As Integer
    Dim cnt As Integer
    
    On Error Resume Next
    
    t = 600
    B = 50

    ' Make sure the form is not minimized
    If Me.WindowState <> vbMinimized Then
        ' Maintain a minimum height and width in order to not set a negative width or height
        If Me.Height < 3000 Or Me.Width < 3000 Then
            If Me.Height < 3000 Then Me.Height = 3000
            If Me.Width < 3000 Then Me.Width = 3000
        Else
            GpTabStrip1.Move Me.ScaleLeft + B, Me.ScaleTop + B, Me.ScaleWidth - B - B, Me.ScaleHeight - B - B
            cmdEsci.Move GpTabStrip1.Left + GpTabStrip1.Width - cmdEsci.Width, GpTabStrip1.Top
            
            For cnt = 0 To picImpo.UBound
                picImpo(cnt).Move GpTabStrip1.Left + B, GpTabStrip1.Top + t, GpTabStrip1.Width - (B * 2), GpTabStrip1.Height - t - B
            Next
            
            lblInfo.Move picImpo(0).Left, picImpo(0).Top - lblInfo.Height, picImpo(0).Width
            
            ListViewMtr.Move 0, 0, picImpo(1).Width, picImpo(2).Height
            WebXml.Move 0, 0, picImpo(4).Width, picImpo(4).Height
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
    frmMain.WindowState = vbNormal
    frmMain.ZOrder
    frmMain.SetFocus
End Sub

Private Sub GpTabStrip1_Click()
    Dim cnt As Integer
    
    'For cnt = 1 To picImpo.UBound
    '    picImpo(cnt).Visible = False
    '    picImpo(cnt).TabStop = False
    'Next
    
    picImpo(GpTabStrip1.SelectTabItem.index).Visible = True
    picImpo(GpTabStrip1.SelectTabItem.index).TabStop = True
    picImpo(GpTabStrip1.SelectTabItem.index).ZOrder
    picImpo(GpTabStrip1.SelectTabItem.index).SetFocus
    
    lblInfo.ZOrder
    cmdEsci.ZOrder
    
End Sub

Private Sub ModificaVariabili(ListView1 As ListView)
    Dim ValoreCellaOpzioni As String
    Dim strTmp As String
    Dim arrTmp
    Dim tmpValore As String
    Dim SepEst As String: SepEst = Chr(0)

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    ValoreCellaOpzioni = GetValoreCella(ListView1, CurrRow, GetNumColDaIntestazione(ListView1, "Opzioni") - 1)
    
    Select Case Left$(ValoreCellaOpzioni, 8)
        Case Is = "°file°°°"
            Dim clsCmd As New clsCommonDialog
            ' The filter entries must be seperated by nulls and terminated with two nulls
            clsCmd.filter = "File " & SplitOne(ValoreCellaOpzioni, "||", 1) & SepEst & SplitOne(ValoreCellaOpzioni, "||", 1) & SepEst & SepEst
            clsCmd.FilterIndex = 1 ' Imposta la posizione predefinita della combo con il file da aprire
            'clsCmd.Filename = NomeFileDefault
            clsCmd.DefaultExtension = SplitOne(ValoreCellaOpzioni, "||", 1)
            clsCmd.DialogTitle = "Seleziona il file"
            clsCmd.hwnd = Me.hwnd       ' the dialog will not be modal unless a hWnd is specified.
            clsCmd.InitDir = DirectoryFromFile(GetValoreCella(ListView1, CurrRow, GetNumColDaIntestazione(ListView1, "Predefinito") - 1))
            clsCmd.Flags = StandardFlag.OpenFile
            clsCmd.ShowOpen
            If clsCmd.CancelPressed = False Then
                ' Scrivo il valore della chiave
                SalvaValoreChiave ListView1, clsCmd.filename
            End If
            Set clsCmd = Nothing
        
        Case Is <> "", Is = "°cancel°"
            If ValoreCellaOpzioni <> "°cancel°" Then
                arrTmp = Split(ValoreCellaOpzioni, "||")
            Else
                arrTmp = Split("", "")
            End If
            tmpValore = GetValoreCella(ListView1, CurrRow, GetNumColDaIntestazione(ListView1, "Valore") - 1)
            
            strTmp = frmMessageCombo.ApriForm(arrTmp, tmpValore, "Chiave: " & GetValoreCella(ListView1, CurrRow, GetNumColDaIntestazione(ListView1, "Chiave") - 1), "Valore attuale:  " & tmpValore & vbNewLine & "Valore predefinito: " & GetValoreCella(ListView1, CurrRow, GetNumColDaIntestazione(ListView1, "Predefinito") - 1))
            If strTmp <> "" Or ValoreCellaOpzioni = "°cancel°" Then
                ' Scrivo il valore della chiave
                SalvaValoreChiave ListView1, strTmp
            End If
            
        Case Else
            ' Nessuna operazione
            
    End Select
    
    Exit Sub
Errore:
    GestErr Err, "Errore nella funzione ModificaVariabili."
  
End Sub

Private Sub SalvaValoreChiave(ListView1 As ListView, Valore As String)

    ' Modifico il valore nella cella della ListView
    Call ScriviCella(ListView1, CurrRow, GetNumColDaIntestazione(ListView1, "Valore"), Valore)
    ' Scrivo il valore nel file Xml
    lVar(GetValoreCella(ListView1, CurrRow, GetNumColDaIntestazione(ListView1, "Indice") - 1)) = Valore

End Sub

Private Sub ListViewMtr_DblClick()
    Call ModificaVariabili(ListViewMtr)
End Sub

Private Sub ListViewMtr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Or Button = vbLeftButton Then
        CurrRow = GetRigaSelezionata(ListViewMtr, X, Y)
    End If
End Sub

Private Sub ListViewMtr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTmp As String
    
    strTmp = Replace$(GetValoreCellaByMouseMove(ListViewMtr, "Descrizione", X, Y), vbCr, "   ")
    strTmp = Replace$(strTmp, vbLf, "   ")
    If Trim(strTmp) = "" Then strTmp = " <--> "
    ListViewMtr.ToolTipText = strTmp
    
End Sub

Private Sub ListViewMtr_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call OrdinaColonnaByTag(ListViewMtr, ColumnHeader)
End Sub
