VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWeb 
   Caption         =   "Crea e Modifica"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   13905
   Icon            =   "frmWeb.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picShadow 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.PictureBox picSlider 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   135
      ScaleWidth      =   12855
      TabIndex        =   1
      Top             =   3120
      Width           =   12855
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser4 
      Height          =   4500
      Left            =   0
      TabIndex        =   34
      Top             =   3285
      Width           =   9765
      ExtentX         =   17224
      ExtentY         =   7937
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
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList imglRunSearch 
      Left            =   480
      Top             =   1320
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
            Picture         =   "frmWeb.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":30F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":4DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":5114
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":526E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":53C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":5522
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":583C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":668E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":74E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":838A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":91DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":B98E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":BCA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":CAFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":CE14
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":E46E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":E788
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":EAA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":EDBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":F0D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":F9B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":1028A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":110DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":11F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":12D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":1365A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":144AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":161B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":164D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":17322
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":19AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":1C286
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":1DF90
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":20742
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2244C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":22D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":23B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":251D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":25AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":26386
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":271D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":28832
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":29684
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2999E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2A278
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2A592
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2A8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2ABC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2AD20
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2B03A
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2B354
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2B66E
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2B988
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2BCA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2BFBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2C116
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2C270
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeb.frx":2C3CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picIntestazione 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1200
      ScaleHeight     =   375
      ScaleWidth      =   11175
      TabIndex        =   13
      Top             =   120
      Width           =   11175
      Begin VB.CommandButton cmdStop 
         Caption         =   "&S T O P"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8880
         TabIndex        =   40
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdMappe 
         Caption         =   "Mappe"
         Height          =   375
         Left            =   7560
         TabIndex        =   31
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox picBmp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   18
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdCosaFai 
         Caption         =   "cmdCosaFai"
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   0
         Width           =   5295
      End
      Begin VB.CommandButton cmdEsci 
         Cancel          =   -1  'True
         Caption         =   "&Esci  [Esc]"
         Height          =   375
         Left            =   9960
         TabIndex        =   16
         Top             =   0
         Width           =   1215
      End
      Begin VB.CheckBox ckAggiungi 
         Appearance      =   0  'Flat
         Caption         =   "Aggiungi nella lista"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   480
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1800
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1440
      Width           =   1575
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
      Left            =   1320
      ScaleHeight     =   990
      ScaleWidth      =   2295
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   2325
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1755
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   3096
      View            =   2
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
   Begin SHDocVwCtl.WebBrowser WebBrowser3 
      Height          =   4500
      Left            =   0
      TabIndex        =   4
      Top             =   3285
      Width           =   10845
      ExtentX         =   19129
      ExtentY         =   7937
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
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   4500
      Left            =   0
      TabIndex        =   2
      Top             =   3285
      Width           =   11925
      ExtentX         =   21034
      ExtentY         =   7937
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
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4500
      Left            =   0
      TabIndex        =   3
      Top             =   3285
      Width           =   12885
      ExtentX         =   22728
      ExtentY         =   7937
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
      Location        =   "http:///"
   End
   Begin VB.PictureBox pictPos 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   12735
      TabIndex        =   5
      Top             =   2400
      Width           =   12735
      Begin VB.CommandButton cmdImpQuestaPag 
         Caption         =   "&Imp. pagina"
         Height          =   375
         Left            =   8160
         TabIndex        =   42
         ToolTipText     =   "Importa le coordinate dalla pagina caricata"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbSitoCoordinate 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Text            =   "cmbSitoCoordinate"
         Top             =   120
         Width           =   2175
      End
      Begin VB.CheckBox ckSoloVuote 
         Caption         =   "Solo Vuote"
         Height          =   195
         Left            =   10560
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
      Begin VB.CheckBox ck1 
         Caption         =   "Solo Corrente Riga"
         Height          =   195
         Left            =   11040
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox ckDaCorrRiga 
         Caption         =   "Da Corrente Riga"
         Height          =   195
         Left            =   9480
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox ckSoloNulle 
         Caption         =   "Solo 0,0"
         Height          =   195
         Left            =   9480
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdImportaPosizione 
         Caption         =   "Importa &posizione"
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   120
         Width           =   4335
      End
   End
   Begin VB.PictureBox picNav 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   12735
      TabIndex        =   35
      Top             =   2400
      Width           =   12735
      Begin VB.CommandButton cmdEnd 
         Caption         =   "&x"
         Height          =   300
         Left            =   12360
         TabIndex        =   41
         ToolTipText     =   "Vai avanti una pagina"
         Top             =   160
         Width           =   300
      End
      Begin VB.CommandButton cmdNavAva 
         Caption         =   "-&>"
         Height          =   375
         Left            =   10800
         TabIndex        =   39
         ToolTipText     =   "Vai avanti una pagina"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNavInd 
         Caption         =   "&<-"
         Height          =   375
         Left            =   0
         TabIndex        =   38
         ToolTipText     =   "Torna indietro una pagina"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txtIndNaviga 
         Height          =   375
         Left            =   500
         TabIndex        =   37
         Text            =   "txtIndNaviga"
         Top             =   120
         Width           =   10305
      End
      Begin VB.CommandButton cmdNaviga 
         Caption         =   "> &Vai >"
         Height          =   375
         Left            =   11400
         TabIndex        =   36
         ToolTipText     =   "Naviga alla pagina impostata nel campo indirizzi"
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox picNomi 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      ScaleHeight     =   615
      ScaleWidth      =   12735
      TabIndex        =   20
      Top             =   2400
      Width           =   12735
      Begin Remakeov2.RMComboView cmbRegione 
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HotButtonBackColor=   0
      End
      Begin Remakeov2.RMComboView cmbProvincia 
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HotButtonBackColor=   0
      End
      Begin VB.CommandButton cmdRicercaSequenziale 
         Caption         =   "Ricerca Sequenziale"
         Height          =   375
         Left            =   5880
         TabIndex        =   29
         ToolTipText     =   $"frmWeb.frx":2C524
         Top             =   120
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.OptionButton optSito 
         Caption         =   "Pagine Bianche"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   28
         Top             =   300
         Width           =   1455
      End
      Begin VB.OptionButton optSito 
         Caption         =   "Pagine Gialle"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   60
         Width           =   1335
      End
      Begin VB.CheckBox cekPagine 
         Caption         =   "Tutte le pagine"
         Height          =   375
         Left            =   10920
         TabIndex        =   26
         Top             =   120
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox ckRegione 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   60
         Width           =   255
      End
      Begin VB.CheckBox ckProvincia 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   300
         Width           =   255
      End
      Begin VB.TextBox txtCognome 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3720
         TabIndex        =   23
         Text            =   "Cognome o Nome Azienda"
         Top             =   0
         Width           =   2040
      End
      Begin VB.TextBox txtNome 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3720
         TabIndex        =   22
         Text            =   "Nome"
         Top             =   300
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.TextBox txtDove 
         Height          =   315
         Left            =   3960
         TabIndex        =   21
         Text            =   "Dove"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdImporta 
         Caption         =   "<<<<-------  Scegli da quale sito importare i dati"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   30
         Top             =   120
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_hwndLV As Long   ' ListView1.hWnd
Private m_hwndTB As Long   ' TextBox1.hWnd
Private m_iItem As Long    ' ListItem.Index whose SubItem is being edited
Private m_iSubItem As Long ' Zero based index of ListView1.ListItems(m_iItem).SubItem being edited

Dim doc As Object
Dim sLinkNext As String
Dim CurrRow As Long
Dim sURL As String
Dim SitoImportazione As String
Dim bEditet As Boolean ' Indica se il file è stato modificato
Dim strClipBoard As String ' Il contenuto della ClipBoard

Dim TextKeyCodeZero As Boolean
Dim CosaFai As String
Dim FileDaAprire As String
Dim bStop As Boolean

Dim CicliDocComplete As Integer ' Serve nel caso si debba attendere più di un ciclo del WebBrowser
Dim LastWebNav As String ' La stringa per verificare l'ultima navigazione inviata al webbrowser

Private cReg As New cRegistry

Dim Dragging As Boolean ' Flag that tells us if the slider is moving

' Classe per il menu CommonDialog
Dim clsCmd As New clsCommonDialog
' Classe per il menu PopUp
Dim mnu As clsMenu

Private Sub ckProvincia_Click()
    Static Indice As Integer
    
    If cmbProvincia.ListIndex <> -1 Then
        Indice = cmbProvincia.ListIndex
    End If

    If ckProvincia.value = 1 Then
        cmbProvincia.Enabled = True
        ' Seleziono il primo elemento
        cmbProvincia.ListIndex = Indice
        txtCognome.Enabled = True
        txtNome.Enabled = True
        cmdRicercaSequenziale.Visible = True
        Call BloccaWebBrowser(True)
    Else
        cmbProvincia.Enabled = False
        cmbProvincia.Text = "Provincia"
        txtCognome.Enabled = False
        txtNome.Enabled = False
        cmdRicercaSequenziale.Visible = False
        Call BloccaWebBrowser(False)
    End If
    
    ckRegione.value = 0
    
End Sub

Private Sub ckRegione_Click()
    Static Indice As Integer
    
    If cmbRegione.ListIndex <> -1 Then
        Indice = cmbRegione.ListIndex
    End If
    
    If ckRegione.value = 1 Then
        cmbRegione.Enabled = True
        ' Seleziono il primo elemento
        cmbRegione.ListIndex = Indice
        txtCognome.Enabled = True
        txtNome.Enabled = True
        cmdRicercaSequenziale.Visible = True
        Call BloccaWebBrowser(True)
    Else
        cmbRegione.Enabled = False
        cmbRegione.Text = "Regione"
        txtCognome.Enabled = False
        txtNome.Enabled = False
        cmdRicercaSequenziale.Visible = False
        Call BloccaWebBrowser(False)
    End If
    
    ckProvincia.value = 0
    
End Sub

Private Sub CkSoloNulle_Click()
    Call ControllaCkSolo(0)
End Sub

Private Sub CkSoloVuote_Click()
    Call ControllaCkSolo(1)
End Sub

Private Sub ControllaCkSolo(Valore)
    
    Select Case Valore
        Case Is = 0
            ckSoloVuote.value = 0
        Case Is = 1
            ckSoloNulle.value = 0
    End Select
    
End Sub

Private Sub cmbProvincia_Click()
    txtDove.Text = cmbProvincia.List(cmbProvincia.ListIndex)
End Sub

Private Sub cmbProvincia_ClickPopUp(ValoreCliccato As Long)
    Call MenuCombo(cmbProvincia.Name, ValoreCliccato)
End Sub

Private Sub cmbRegione_Click()
    txtDove.Text = cmbRegione.List(cmbRegione.ListIndex)
End Sub

Private Sub cmbRegione_ClickPopUp(ValoreCliccato As Long)
    Call MenuCombo(cmbRegione.Name, ValoreCliccato)
End Sub

Private Sub MenuCombo(NomeCombo As String, ByVal Indice As Long)
    
    Select Case Indice
        Case Is = 10
            frmPrendiRegPr.Show vbModal, Me
        
        Case Is = 20
            If NomeCombo = "cmbRegione" Then cmbRegione.SaveFile Var(RegioniCsv).Valore, True
            If NomeCombo = "cmbProvincia" Then cmbProvincia.SaveFile Var(ProvinceCsv).Valore, True
            
    End Select
End Sub

Private Sub cmdCosaFai_Click()
    Call ControllaModalità
End Sub

Public Sub ControllaModalità(Optional AttivaMod As String = "")
    
    If AttivaMod <> "" Then
        ' Imposta la condizione scelta (praticamente imposta quella precedente a quella scelta)
        Select Case AttivaMod
            Case Is = "Edit"
                CosaFai = "Navigazione Web"
                
            Case Is = "ImpNomi"
                CosaFai = "Edit"
                
            Case Is = "Importa Posizione"
                CosaFai = "ImpNomi"
                
            Case Is = "Navigazione Web"
                CosaFai = "Importa Posizione"
        End Select
    End If
    
    Select Case CosaFai
        Case "", "Navigazione Web"
            CosaFai = "Edit"
            cmdCosaFai.Caption = "Modalità: Edita lista   (clicca per cambiare)"
            WebBrowser1.Visible = False
            WebBrowser2.Visible = False
            WebBrowser3.Visible = False
            WebBrowser4.Visible = False
            picNomi.Visible = False
            pictPos.Visible = False
            picNav.Visible = False
            cmbRegione.Visible = False
            cmbProvincia.Visible = False
        Case "Edit"
            CosaFai = "ImpNomi"
            cmdCosaFai.Caption = "Modalità: Importa Nomi   (clicca per cambiare)"
            WebBrowser1.Visible = False
            WebBrowser2.Visible = True
            WebBrowser3.Visible = False
            WebBrowser4.Visible = False
            picNomi.Visible = True
            pictPos.Visible = False
            picNav.Visible = False
            cmbRegione.Visible = True
            cmbProvincia.Visible = True
            cmbRegione.ZOrder
            cmbProvincia.ZOrder
        Case "ImpNomi"
            CosaFai = "Importa Posizione"
            cmdCosaFai.Caption = "Modalità: Importa Posizione   (clicca per cambiare)"
            WebBrowser1.Visible = True
            WebBrowser2.Visible = False
            WebBrowser3.Visible = False
            WebBrowser4.Visible = False
            picNomi.Visible = False
            pictPos.Visible = True
            picNav.Visible = False
            cmbRegione.Visible = False
            cmbProvincia.Visible = False
        Case "Importa Posizione"
            CosaFai = "Navigazione Web"
            cmdCosaFai.Caption = "Modalità: Navigazione Web   (clicca per cambiare)"
            WebBrowser1.Visible = False
            WebBrowser2.Visible = False
            WebBrowser3.Visible = False
            WebBrowser4.Visible = True
            picNomi.Visible = False
            pictPos.Visible = False
            picNav.Visible = True
            cmbRegione.Visible = False
            cmbProvincia.Visible = False
        Case Else
            CosaFai = ""
    End Select

    If CosaFai = "Edit" Then
        picSlider.Move 0, Me.Height
        picShadow.Move 0, Me.Height
        Call SliderMove
    Else
        picSlider.Move 0, 3100
        picShadow.Move 0, 3100
        Call SliderMove
    End If
    
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdMappe_Click()
    frmImmagine.Show
    DoEvents
End Sub

Private Sub cmdNavInd_Click()
    On Error Resume Next
    WebBrowser4.GoBack
End Sub

Private Sub cmdNavAva_Click()
    On Error Resume Next
    WebBrowser4.GoForward
End Sub

Private Sub cmdNaviga_Click()
    WebBrowser4.Navigate2 txtIndNaviga.Text
End Sub

Private Sub cmdEnd_Click()
    WebBrowser4.Navigate2 "about:blank"
End Sub

Private Sub cmdStop_Click()
    bStop = True
End Sub

Private Sub WebBrowser4_DocumentComplete(ByVal pDisp As Object, url As Variant)
    
    If LCase(WebBrowser4.LocationURL) <> "about:blank" Then
        txtIndNaviga.Text = WebBrowser4.LocationURL
    End If
    
End Sub

Private Sub Form_Activate()
    ListView1.SetFocus
End Sub

Private Sub Form_Load()
    Dim LarghezzaColonne As Variant
    Dim i As Long

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    Me.picShadow.Height = 25

    Text1.Visible = False
    m_hwndTB = Text1.hwnd

    LarghezzaColonne = Array(0, 1150, 2800, 2200, 980, 1600, 650, 1200, 1200, 980, 980, 3000, 1200)
    
    With ListView1
        .HideSelection = False
        .MultiSelect = Var(SelezMultipla).Valore
        .FullRowSelect = True
        .View = lvwReport
        m_hwndLV = .hwnd
        .LabelEdit = lvwManual ' Evita che si possa editare la prima colonna
        .Icons = imglRunSearch
        .SmallIcons = imglRunSearch
        .ColumnHeaderIcons = imglRunSearch
        .ColumnHeaders.Add , "x1", "    "
        .ColumnHeaders.Item(1).Width = LarghezzaColonne(1)
        .ColumnHeaders.Add , "x2", " 1 Descrizione"
        .ColumnHeaders.Item(2).Width = LarghezzaColonne(2)
        .ColumnHeaders.Add , "x3", " 2 Indirizzo"
        .ColumnHeaders.Item(3).Width = LarghezzaColonne(3)
        .ColumnHeaders.Add , "x4", " 3 Cap"
        .ColumnHeaders.Item(4).Width = LarghezzaColonne(4)
        .ColumnHeaders.Add , "x5", " 4 Città"
        .ColumnHeaders.Item(5).Width = LarghezzaColonne(5)
        .ColumnHeaders.Add , "x6", " 5 Pr"
        .ColumnHeaders.Item(6).Width = LarghezzaColonne(6)
        .ColumnHeaders.Add , "x7", " 6 Telefono"
        .ColumnHeaders.Item(7).Width = LarghezzaColonne(7)
        .ColumnHeaders.Add , "x8", " 7 Categoria"
        .ColumnHeaders.Item(8).Width = LarghezzaColonne(8)
        .ColumnHeaders.Add , "x9", " 8 Latitudine"
        .ColumnHeaders.Item(9).Width = LarghezzaColonne(9)
        .ColumnHeaders.Add , "x10", " 9 Longitudine"
        .ColumnHeaders.Item(10).Width = LarghezzaColonne(10)
        .ColumnHeaders.Add , "x11", " 10 Sito Web"
        .ColumnHeaders.Item(11).Width = LarghezzaColonne(11)
        .ColumnHeaders.Add , "x12", " 11 Desc.Ov2"
        .ColumnHeaders.Item(11).Width = LarghezzaColonne(12)
    End With

    Call SetListViewColor(ListView1, Picture1, 1, vbWhite, vbGreenLemon)
    Call AutoSizeUltimaColonna(ListView1)

    CurrRow = 0
    
    With cmbRegione
        .Style = Checkboxes
        .LoadFile Var(RegioniCsv).Valore, True
        .Enabled = False
        .Text = "Regione"
        .AddItemMenu 10, "Scarica Elenco"
        .AddItemMenu 20, "Salva Elenco"
        .ColWidth(0) = .Width - 240
    End With
    With cmbProvincia
        .Style = Checkboxes
        .LoadFile Var(ProvinceCsv).Valore, True
        .Enabled = False
        .Text = "Provincia"
        .AddItemMenu 10, "Scarica Elenco"
        .AddItemMenu 20, "Salva Elenco"
        .ColWidth(0) = .Width - 240
    End With
    
    With cmbSitoCoordinate
        .AddItem "* ViaMichelin.it"
        .AddItem "* Mappe.Libero.it"
        .AddItem "* TuttoCittà.it"
        .AddItem "* Google.com"
        .AddItem "* MultiMap.com"
        .ListIndex = 2
    End With
    Call BloccaComboBox(cmbSitoCoordinate)
    Call CaricaScript
    
    WebBrowser1.Navigate2 "about:blank"
    WebBrowser2.Navigate2 "about:blank"
    WebBrowser3.Navigate2 "about:blank"
    WebBrowser4.Navigate2 "about:blank"
    ' Funzioni da finire
    'WebBrowser4.Navigate cReg.ValueEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Favorites", REG_SZ, "about:blank")
    'ExtractAll cReg.ValueEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Favorites", REG_SZ, "about:blank"), App.Path
    arrListViewVuoto = True
    
    txtIndNaviga.Text = HomePage
    cmdStop.Enabled = False
    bStop = False
    
    If CommandLineFile <> "" Then
        FileDaAprire = CommandLineFile
        Call ClickPopUp(20, FileDaAprire)
        CommandLineFile = ""
        FileDaAprire = ""
    End If

    Call ControllaModalità("Edit")
    
    ' Cancello evenutali valori rimasti assegnati alla variabile
    SetupDescrizione = ""
    
    frmMain.Visible = False
    
    Exit Sub
    
Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione frmWeb.Form_Load."

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
    
    If ListView1.ListItems.count > 0 Then
        Cancel = ConfermaChiusuraForm
    End If
    
    If Cancel = 0 Then
        bStop = True
        
        Set mnu = Nothing
        Set cReg = Nothing
        
        WebBrowser1.Navigate2 "about:blank"
        WebBrowser2.Navigate2 "about:blank"
        WebBrowser3.Navigate2 "about:blank"
        WebBrowser4.Navigate2 "about:blank"
            
        DoEvents
        
        NomeFileAperto = ""
        If SecondaIstanza = True Then
            Unload frmMain
        Else
            frmMain.Visible = True
            frmMain.WindowState = vbNormal
            frmMain.ZOrder
            frmMain.SetFocus
        End If
    End If
    
    DoEvents
    
End Sub

Private Sub Form_Resize()
   ' Width = Larghezza  Height = Altezza
    On Error Resume Next
    
    ' Se la form non è minimizzata procedo al resize
    If Me.WindowState <> vbMinimized Then
        If Me.Width < (cmdCosaFai.Width + cmdEsci.Width + 500) Then Me.Width = cmdCosaFai.Width + cmdEsci.Width + 500
        If Me.Height < 5000 Then Me.Height = 5000

        picSlider.Height = 55
        picShadow.Height = 35
        picShadow.Top = picSlider.Top
        
        ' Mi assicuro che i controlli abbiano le solte misure
        picNomi.Height = pictPos.Height
        picNomi.Width = pictPos.Width
        picNav.Height = pictPos.Height
        picNav.Width = pictPos.Width
    
        ' Maintain a minimum height and width in order to not set a negative width or height
        If Me.Height < 1200 Or Me.Width < 1200 Then Exit Sub

        ' Centro i controlli nella form
        picNomi.Move (Me.ScaleWidth - picNomi.Width) / 2
        pictPos.Move ((Me.ScaleWidth - pictPos.Width) / 2)
        picNav.Move ((Me.ScaleWidth - pictPos.Width) / 2)
        picIntestazione.Move (Me.ScaleWidth - picIntestazione.Width) / 2
        
        ' Imposto la larghezza dei controlli
        ListView1.Width = Me.ScaleWidth
        WebBrowser1.Width = Me.ScaleWidth
        WebBrowser2.Width = Me.ScaleWidth
        WebBrowser3.Width = Me.ScaleWidth
        WebBrowser4.Width = Me.ScaleWidth
        picShadow.Width = Me.ScaleWidth
        picSlider.Width = Me.ScaleWidth
    
        ' Make sure we don't give any of the textboxes a negative value
        If Me.picSlider.Top > Me.ScaleHeight Or Me.picSlider.Top - 300 < 100 Then
            Me.picSlider.Top = Me.ScaleHeight
            Me.picShadow.Top = Me.picSlider.Top
        End If
        
        If CosaFai = "Edit" Then
            picSlider.Move 0, Me.Height
            picShadow.Move 0, Me.Height
        End If
        
        cmbRegione.Move ckRegione.Left + ckRegione.Width + 25, txtCognome.Top, cmbRegione.Width, txtCognome.Height
        cmbProvincia.Move ckProvincia.Left + ckProvincia.Width + 25, txtNome.Top, cmbProvincia.Width, txtNome.Height

        Call SliderMove
    End If

End Sub

Private Sub SliderMove()
    On Error Resume Next
    
    ' Turn off dragging and hide the shadow
    Dragging = False
    picShadow.Visible = False
    
    ' Make sure the shadow was not moved too far
    If picShadow.Top + 70 > Me.ScaleHeight Then picShadow.Top = Me.ScaleHeight - 70
    If picShadow.Top < 600 Then picShadow.Top = 600
    
    ' Move picSlider and resize the controls
    Me.picSlider.Top = Me.picShadow.Top
    Me.ListView1.Height = Me.picSlider.Top - Me.ListView1.Top
    picNomi.Top = Me.picSlider.Top + Me.picSlider.Height
    pictPos.Top = Me.picSlider.Top + Me.picSlider.Height
    picNav.Top = Me.picSlider.Top + Me.picSlider.Height
    Me.WebBrowser1.Top = (Me.picSlider.Top + Me.picSlider.Height) + picNomi.Height
    Me.WebBrowser2.Top = (Me.picSlider.Top + Me.picSlider.Height) + picNomi.Height
    Me.WebBrowser3.Top = (Me.picSlider.Top + Me.picSlider.Height)
    Me.WebBrowser4.Top = (Me.picSlider.Top + Me.picSlider.Height) + picNav.Height
    Me.WebBrowser1.Height = (Me.ScaleHeight - Me.picSlider.Top + Me.picSlider.Height - picNomi.Height) - 100
    Me.WebBrowser2.Height = (Me.ScaleHeight - Me.picSlider.Top + Me.picSlider.Height - picNomi.Height) - 100
    Me.WebBrowser3.Height = (Me.ScaleHeight - Me.picSlider.Top + Me.picSlider.Height) - 100
    Me.WebBrowser4.Height = (Me.ScaleHeight - Me.picSlider.Top + Me.picSlider.Height - picNav.Height) - 100

End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ControllaRigaChecked(ListView1, Item)
End Sub

Private Sub picSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True ' Set the dragging flag
End Sub

Private Sub picSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging = True Then
        ' Make sure the slider shadow is visible
        If Not picShadow.Visible Then picShadow.Visible = True
        ' Move the shadow
        picShadow.ZOrder 0
        picShadow.Top = Me.picSlider.Top + Y
    End If
End Sub

Private Sub picSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SliderMove
End Sub

Private Sub VisualizzaSuMappa()
    Dim tmpLon As String
    Dim tmpLat As String
    Dim longitude As Double
    Dim latitude As Double
    
    If (FormIsLoad("frmImmagine") = True) And (CurrRow > 0) And (ListView1.ListItems.count > 0) Then
        tmpLon = GetValoreCellaByNome(ListView1, CurrRow, "Longitudine")
        tmpLat = GetValoreCellaByNome(ListView1, CurrRow, "Latitudine")
        If IsNumeric(tmpLon) And IsNumeric(tmpLat) = True Then
            longitude = CDbl(tmpLon)
            latitude = CDbl(tmpLat)
            frmImmagine.andrMapCtl1.SetLonLat longitude, latitude
        End If
    End If
    
End Sub

Private Sub ListView1_DblClick()
    Call EditListView
End Sub

Private Sub EditListView(Optional EditaSempre As Boolean = False, Optional RigaSel As Long = -1, Optional ColonnaSel As Long = -1)
    Dim lvhti As LVHITTESTINFO
    Dim RC As RECT
    Dim li As ListItem
    
    ' If a left button double-click... (change to suit)
    If (GetKeyState(vbKeyLButton) And &H8000) Or EditaSempre = True Then
    
      If ListView1.FullRowSelect = True Then ListView1.FullRowSelect = False

      Call GetCursorPos(lvhti.pt)
      Call ScreenToClient(m_hwndLV, lvhti.pt)
      
        If EditaSempre = False Then
            RigaSel = ListView_SubItemHitTest(m_hwndLV, lvhti)
        Else
            ' Elaboro la riga selezionata con il mouse o passata dalla funzione
            If RigaSel <= -1 Then
              RigaSel = ListView1.ListItems.count - 1
            Else
              'If ((ListView_SubItemHitTest(m_hwndLV, lvhti) <> LVI_NOITEM And ListView_SubItemHitTest(m_hwndLV, lvhti) >= 0) And (ListView_SubItemHitTest(m_hwndLV, lvhti) <= ListView1.ListItems.Count)) _
                 ' And RigaSel < ListView1.ListItems.Count Then
              If (RigaSel < ListView1.ListItems.count) Then
                  RigaSel = RigaSel
              Else
                  RigaSel = 0
              End If
            End If
            
            ' Elaboro la colonna selezionata con il mouse o passata dalla funzione
            If ColonnaSel > 0 And ColonnaSel <= ListView1.ColumnHeaders.count - 1 Then
              lvhti.iSubItem = ColonnaSel
            Else
              ' mi sposto sulla prima colonna o sull'ultima colonna
              If ColonnaSel <> -1 Then
                    
                    If ColonnaSel = 0 Then ' Mi sposto alla riga precedente
                        If RigaSel <= ListView1.ListItems.count - 1 And RigaSel > 0 Then
                            RigaSel = RigaSel - 1
                        Else
                            RigaSel = ListView1.ListItems.count - 1
                        End If
                        ' mi sposto sull'ultima colonna
                        ColonnaSel = ListView1.ColumnHeaders.count - 1
                        lvhti.iSubItem = ColonnaSel
 
                    Else ' mi sposto alla riga successiva
                        If RigaSel < ListView1.ListItems.count - 1 Then
                            RigaSel = RigaSel + 1
                        Else
                            RigaSel = 0
                        End If
                        ' mi sposto sulla prima colonna
                        ColonnaSel = 1
                        lvhti.iSubItem = ColonnaSel
                    
                    End If
              End If
            End If
        End If

        lvhti.iItem = RigaSel

      ' Mi assicuro che la riga selezionata esista
      If (RigaSel <> LVI_NOITEM And RigaSel >= 0) And (RigaSel <= ListView1.ListItems.count) Then
        ' Mi assicuro che la riga sia selezionata e visibile
        ListView1.SelectedItem = ListView1.ListItems(RigaSel + 1)
        ListView1.SelectedItem.EnsureVisible
        
        If lvhti.iSubItem Then
          
          ' Get the SubItem's label (and icon) rect.
          If ListView_GetSubItemRect(m_hwndLV, lvhti.iItem, lvhti.iSubItem, LVIR_LABEL, RC) = True Then
            
            ' Either set the ListView as the TextBox parent window in order to
            ' have the TextBox Move method use ListView client coords, or just
            ' map the ListView client coords to the TextBox's paent Form
            'Call SetParent(m_hwndTB, m_hwndLV)
            Call MapWindowPoints(m_hwndLV, hwnd, RC, 2)
            Text1.Move (RC.Left + 4) * Screen.TwipsPerPixelX, RC.Top * Screen.TwipsPerPixelY, (RC.Right - RC.Left) * Screen.TwipsPerPixelX, (RC.Bottom - RC.Top) * Screen.TwipsPerPixelY
            
            ' Save the one-based index of the ListItem and the zero-based index
            ' of the SubItem(if the ListView is sorted via the  API, then ListItem.Index
            ' will be different than lvhti.iItem +1...)
            m_iItem = lvhti.iItem + 1
            m_iSubItem = lvhti.iSubItem
             
            ' Put the SubItem's text in the TextBox, save the SubItem's text,
            ' and clear the SubItem's text.
            Text1 = ListView1.ListItems(m_iItem).SubItems(m_iSubItem)
            Text1.Tag = Text1
            ListView1.ListItems(m_iItem).SubItems(m_iSubItem) = ""
            
            ' Make the TextBox the topmost Form control, make the it visible, select
            ' its text, give it the focus, and subclass it.
            Text1.ZOrder 0
            Text1.Visible = True
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1)
            Text1.SetFocus
            Call SubClass(m_hwndTB, AddressOf WndProc)
            
          End If   ' ListView_GetSubItemRect
        End If   ' lvhti.iSubItem
      End If   ' ListView_SubItemHitTest
    End If   ' GetKeyState(vbKeyLButton)

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call OrdinaColonnaByTag(ListView1, ColumnHeader)
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
    If Button = vbRightButton Or Button = vbLeftButton Then
        CurrRow = GetRigaSelezionata(ListView1, X, Y)
    End If
    
End Sub

Private Sub ListView1_Click()

    If CosaFai = "Navigazione Web" Then
        ' Qua si dovrebbe inserire il codice che permette di modificare la lista senza fare doppio click
    End If
        
    If ListView1.FullRowSelect = False Then ListView1.FullRowSelect = True
    
    ' Mi assicuro che la riga sia selezionata sia visibile
    If CurrRow > 0 And ListView1.ListItems.count > 0 Then
        If CurrRow <= ListView1.ListItems.count Then
            CurrRow = CurrRow
        Else
            CurrRow = ListView1.ListItems.count
        End If
        ListView1.SelectedItem = ListView1.ListItems(CurrRow)
        ListView1.SelectedItem.EnsureVisible
        Call VisualizzaSuMappa
    End If
    
    cmdImpQuestaPag.Enabled = False
    
    NascondiScrollBar ListView1, False
    
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If ListView1.ListItems.count > 0 Then
        CurrRow = CInt(ListView1.SelectedItem.index)
    End If

    ' Per cancellare le righe
    If KeyCode = 46 Then 'Se il tasto premuto è Canc
        ' Mi assicuro che la ListView non è vuota
        If Not IsNull(ListView1.SelectedItem) And Not ListView1.SelectedItem Is Nothing Then
            ListView1.ListItems.remove (ListView1.SelectedItem.index)
        End If
        
    ElseIf KeyCode = 93 Then ' il tasto per il menu del tato destro del mouse
        If CurrRow > 0 Then MouseMove Me, (ListView1.ListItems.Item(CurrRow).Left) - 25, (ListView1.ListItems.Item(CurrRow).Top / 15) + 55
        Call ListView1_MouseUp(vbRightButton, 0, 0, 0)
        
    End If

    If FormIsLoad("frmSetupDescrizione") = True Then
        frmSetupDescrizione.lblCambiaRigaCorrente.Caption = CurrRow
    End If

    Call VisualizzaSuMappa

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        ' Mi assicuro che la ListView non è vuota
        If ListView1.ListItems.count > 0 Then
            If ListView1.FullRowSelect = True Then ListView1.FullRowSelect = False
            ListView1.Refresh
            DoEvents
            EditListView True, ListView1.SelectedItem - 1, 1
        End If
        
    ElseIf KeyAscii = 22 Then ' Ctrl + v
        ClickPopUp 81
        
    End If

End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ListaVuota As Boolean
    Dim NoCoordinate As Boolean
    Dim m_cClip As New cCustomClipboard
    Dim bGrey As Boolean
    
    If ListView1.ListItems.count = 0 Then
        ListaVuota = True
    Else
        ListaVuota = False
    End If
    
    If ListaVuota = True Then
        NoCoordinate = True
    Else
        If (ControllaCella(ListView1, CurrRow, 8) And ControllaCella(ListView1, CurrRow, 9)) = True _
        And (GetValoreCella(ListView1, CurrRow, 8) <> "0,0" And GetValoreCella(ListView1, CurrRow, 9) <> "0,0") _
        Then
            NoCoordinate = False
        Else
            NoCoordinate = True
        End If
    End If
    
    ' Per il PopUp menu
    If Button = vbRightButton Then
        Set mnu = New clsMenu
        With mnu
            If ListaVuota = False And NoCoordinate = False Then
                Dim submnu1 As clsMenu: Set submnu1 = New clsMenu
                With submnu1
                    .Caption = "Mostra sulla cartina"
                    .AddItem 101, "www.multimap.com (" & CurrRow & ")", , , , NoCoordinate
                    .AddItem 102, "www.mapquest.com (" & CurrRow & ")", , , , NoCoordinate
                    .AddItem 103, "www.it.map24.com (" & CurrRow & ")", , , , NoCoordinate
                    .AddItem 104, "www16.mappy.com (" & CurrRow & ")", , , , NoCoordinate
                    .AddItem 105, "GoogleMaps.it (" & CurrRow & ")", , , , NoCoordinate
                End With
                '
                .AddItem 10, submnu1
            End If
            If ListaVuota = False Then
                .AddItem 11, "Controlla errori nelle coordinate", , , , ListaVuota
                .AddItem 12, "Inverti le due colonne delle coordinate", , , , ListaVuota
                .AddItem 0, "-"
            End If
            '
            .AddItem 20, "Apri...."
            .AddItem 30, "Salva... " & NomeFileAperto, , , , (Not ListaVuota Imp Not FileExists(PatchNomeFileAperto))
            .AddItem 31, "Salva con nome...", , , , ListaVuota
            .AddItem 0, "-"
            '
            .AddItem 40, "Setup Descrizione", , , , ListaVuota
            .AddItem 41, "Rimuovi Duplicati", , , , ListaVuota
            .AddItem 0, "-"
            '
            .AddItem 48, "Numera le righe della lista", , , , ListaVuota
            .AddItem 0, "-"
            '
            .AddItem 50, "AutoSize larghezza colonne", , , , ListaVuota
            .AddItem 51, "Selezione righe multiple", , Var(SelezMultipla).Valore, , ListaVuota
            .AddItem 52, "Backup automatico file in importazione coordinate ogni " & Var(SalvaOgni).Valore & " righe", , Var(SalvataggioBackup).Valore, , ListaVuota
            .AddItem 0, "-"
            '
            .AddItem 60, "Elimina riga selezionata", , , , ListaVuota
            .AddItem 70, "Elimina righe selezionate", , , , ListaVuota
            .AddItem 80, "Aggiungi nuova riga"
            .AddItem 0, "-"
            '
            ' Apro la ClipBoard
            If ((m_cClip.ClipboardOpen(Me.hwnd)) = True) And (m_cClip.GetTextData(1, strClipBoard) = True And (strClipBoard <> "" And InStr(1, strClipBoard, vbTab, vbTextCompare) <> 0)) Then bGrey = True
            m_cClip.ClipboardClose
            .AddItem 81, "Incolla da foglio di calcolo", , , , Not (bGrey)
            .AddItem 82, "Copia Riga negli Appunti", , , , ListaVuota
            .AddItem 83, "Copia Tutte le Righe negli Appunti", , , , ListaVuota
            '
            ClickPopUp .TrackPopup(Me.hwnd)
        End With

    End If

    If FormIsLoad("frmSetupDescrizione") = True Then
        frmSetupDescrizione.lblCambiaRigaCorrente.Caption = GetRigaSelezionata(ListView1, X, Y)
    End If

End Sub

Private Sub ClickPopUp(ValoreCliccato As Long, Optional FileDaAprire As String = "")
    Dim Result
    Dim ret
    Dim PatchNomeFile As String
    Dim CancellaListView As Boolean
    Dim SaltaRighe As Integer
    Dim cnt As Long
    Dim cntCol As Long
    Dim strTmp As String
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    If ckAggiungi.value = 1 Then
        CancellaListView = False
    Else
        CancellaListView = True
    End If
    
    PatchNomeFile = ""
    
    Select Case Left(ValoreCliccato, 2)
    
        Case Is = 10 ' Mostra posizione sulla cartina
                Screen.MousePointer = vbHourglass
                DoEvents
                If ControllaCella(ListView1, CurrRow, 8) = False Then Exit Sub
                If ControllaCella(ListView1, CurrRow, 9) = False Then Exit Sub
                WebBrowser3.Navigate2 "about:blank"
                DoEvents
                Call ControllaModalità("Importa Posizione")
                WebBrowser1.Visible = False
                WebBrowser2.Visible = False
                WebBrowser3.Visible = True
                picNomi.Visible = False
                pictPos.Visible = False
            If ValoreCliccato = 101 Then
                WebBrowser3.Navigate2 MostraPosizione(ListView1, "multimap", CurrRow)
            ElseIf ValoreCliccato = 102 Then
                WebBrowser3.Navigate2 MostraPosizione(ListView1, "mapquest", CurrRow)
            ElseIf ValoreCliccato = 103 Then
                WebBrowser3.Navigate2 MostraPosizione(ListView1, "map24", CurrRow)
            ElseIf ValoreCliccato = 104 Then
                WebBrowser3.Navigate2 MostraPosizione(ListView1, "www16.mappy", CurrRow)
            ElseIf ValoreCliccato = 105 Then
                WebBrowser3.Navigate2 MostraPosizione(ListView1, "GoogleMaps", CurrRow)
            End If
                Screen.MousePointer = vbDefault

        Case Is = 11 ' Controllo errori nelle coordinate
            Call VerificaCoordinate(ListView1, 8, 9, SplitOne(Var(LimiteLat).Valore, "|", 0), SplitOne(Var(LimiteLat).Valore, "|", 1), SplitOne(Var(LimiteLon).Valore, "|", 0), SplitOne(Var(LimiteLon).Valore, "|", 1))
            
        Case Is = 12 ' Inversione delle coordinate
            Call InvertiColonne(ListView1, 8, 9)
            
        Case Is = 20 ' Apri
            ' Prevent the ListView control from updating on screen - this is to hide the changes being made to the listitems and also to speed up the sort
            LockWindowUpdate ListView1.hwnd
            If ImportaDati(Me.hwnd, ListView1, FileDaAprire, , True, CancellaListView, Var(UltimoFilePOI).Valore, "*" & Right$(Var(UltimoFilePOI).Valore, 4), , "FileDatiPOI", False, ckAggiungi.value) = False Then
                Me.Caption = "Crea e Modifica"
                Result = MsgBox("Non è stato possibile importare il file!", vbOKOnly)
                ckAggiungi.Visible = CBool(ListView1.ListItems.count)
            Else
                If ckAggiungi.value = 0 Then
                    Me.Caption = FileDaAprire
                    ' Scrivo i dati nel file .xml
                    lVar(UltimoFilePOI) = FileDaAprire
                End If
                If CancellaListView = False Then Call NumeraListView(ListView1)
                Call ControllaCheck(ListView1)
                ckAggiungi.Visible = True
                If Var(AutoVerificaCoordinate).Valore > 0 Then
                    Call VerificaCoordinate(ListView1, 8, 9, SplitOne(Var(LimiteLat).Valore, "|", 0), SplitOne(Var(LimiteLat).Valore, "|", 1), SplitOne(Var(LimiteLon).Valore, "|", 0), SplitOne(Var(LimiteLon).Valore, "|", 1), Var(AutoVerificaCoordinate).Valore, False)
                End If
            End If
            
            If ckAggiungi.Visible = False Then ckAggiungi.value = 0
            
            ' Unlock the list window so that the OCX can update it
            LockWindowUpdate 0&
            
            Call CaricaImmagine(picBmp, NomeFileAperto, "bmp")
            
            Screen.MousePointer = vbDefault

        Case Is = 30 ' Salva
            If ListView1.ListItems.count <> 0 Then
                If Var(AutoVerificaCoordinate).Valore > 0 Then
                    Call VerificaCoordinate(ListView1, 8, 9, CDbl(SplitOne(Var(LimiteLat).Valore, "|", 0)), CDbl(SplitOne(Var(LimiteLat).Valore, "|", 1)), CDbl(SplitOne(Var(LimiteLon).Valore, "|", 0)), CDbl(SplitOne(Var(LimiteLon).Valore, "|", 1)), Var(AutoVerificaCoordinate).Valore, False)
                End If
                strTmp = Var(CampiRmkFile).Valore & "FileDatiPOI" & ";;" & Now & ";" & SetupDescrizione
                strTmp = ExportaDati(Me.hwnd, PatchNomeFileAperto, , , True, , strTmp, ListView1)
                If strTmp = 0 Then MsgBox "Sono stati esportati " & strTmp & " record." & vbNewLine & "Forse stavi salvando in un file .ov2 record senza i dati delle coordinate! ", vbInformation, App.ProductName
            End If
            Screen.MousePointer = vbDefault
            
        Case Is = 31 ' Salva con nome
            If ListView1.ListItems.count <> 0 Then
                If Var(AutoVerificaCoordinate).Valore > 0 Then
                    Call VerificaCoordinate(ListView1, 8, 9, CDbl(SplitOne(Var(LimiteLat).Valore, "|", 0)), CDbl(SplitOne(Var(LimiteLat).Valore, "|", 1)), CDbl(SplitOne(Var(LimiteLon).Valore, "|", 0)), CDbl(SplitOne(Var(LimiteLon).Valore, "|", 1)), Var(AutoVerificaCoordinate).Valore, False)
                End If
                strTmp = Var(CampiRmkFile).Valore & "FileDatiPOI" & ";;" & Now & ";" & SetupDescrizione
                strTmp = ExportaDati(Me.hwnd, , NomeFileAperto, , True, , strTmp, ListView1)
                If strTmp = 0 Then MsgBox "Sono stati esportati " & strTmp & " record." & vbNewLine & "Forse stavi salvando in un file .ov2 record senza i dati delle coordinate!" & vbNewLine & "In questo caso utilizza il formato .rmk.", vbInformation, App.ProductName
            End If
            Screen.MousePointer = vbDefault
            
        Case Is = 40 ' Apre form SetupDescrizione
            frmSetupDescrizione.Show , Me
            frmSetupDescrizione.lblCambiaRigaCorrente.Caption = CurrRow
        
        Case Is = 41 ' Apre form RimuoviDuplicati
            Screen.MousePointer = vbHourglass
            frmRimuoviDuplicati.Show , Me
            Screen.MousePointer = vbDefault

        Case Is = 48 ' Numera le righe della lista
            Call NumeraListView(ListView1)

        Case Is = 50
            Call AutoSizeColonne(ListView1)
        
        Case Is = 51
            Call ChangeSelezMultipla(ListView1)
        
        Case Is = 52
            lVar(SalvataggioBackup) = Not Var(SalvataggioBackup).Valore

        Case Is = 60
            ret = MsgBox("Vuoi davvero cancellare la riga selezionata?", vbYesNo + vbExclamation + vbDefaultButton1)
            If ret = vbYes Then
                Call CancellaRiga(ListView1)
            End If
        
        Case Is = 70
            ret = MsgBox("Vuoi davvero cancellare le righe selezionate?", vbYesNo + vbExclamation + vbDefaultButton1)
            If ret = vbYes Then
                Call CancellaRiga(ListView1, , True)
            End If

        Case Is = 80 ' Aggiunge nuova riga
            Dim itmX As Variant
            If CurrRow = 0 Then CurrRow = 1
            
            If ListView1.ListItems.count = 0 Then
                Set itmX = ListView1.ListItems.Add
            Else
                Set itmX = ListView1.ListItems.Add(CurrRow + 1)
            End If
            itmX.SubItems(1) = "Doppio click per scrivere......"
            Call NumeraListView(ListView1)
            Call SelezionaRigaListView(ListView1, CurrRow + 1, True)
            Call ControllaCheck(ListView1)
            
        Case Is = 81 ' Incolla da foglio di calcolo
            Dim RowClip As Variant
            Dim ColClip As Variant
            Dim NumColl As Long
            
            RowClip = Split(strClipBoard, vbCrLf)
            
            ReDim arrListView(UBound(RowClip), ListView1.ColumnHeaders.count - 2)
            
            ' Scrivo le intestazioni delle colonne
            For cntCol = 0 To UBound(arrListView, 2)
                arrListView(0, cntCol) = Trim$(ListView1.ColumnHeaders(cntCol + 2).Text)
            Next
            
            ' Scorro tutte le righe
            For cnt = 0 To UBound(RowClip)
                ColClip = Split(RowClip(cnt), vbTab)
                If UBound(arrListView, 2) > UBound(ColClip) Then
                    NumColl = UBound(ColClip)
                Else
                    NumColl = UBound(arrListView, 2)
                End If
                ' Scorro tutte le colonne
                For cntCol = 0 To NumColl
                    On Error Resume Next
                    arrListView(cnt + 1, cntCol) = ColClip(cntCol)
                    Err.Clear
                Next
            Next
            
            arrListViewVuoto = False
            Call CaricaListViewDaArray(ListView1, False)

        Case Is = 82 ' Copia Riga
            strTmp = GetValoreRiga(ListView1, CurrRow)
            If strTmp <> "" Then
                Clipboard.Clear
                Clipboard.SetText Replace(strTmp, vbTab & "+", vbTab & "'+") & vbCrLf, vbCFText
                Call ControllaCheck(ListView1)
            End If
            
            
        Case Is = 83 ' Copia Tutte le Righe
            Dim strTabella As String
            
            ' Scorro tutte le righe
            For cnt = 1 To ListView1.ListItems.count
                strTmp = Replace(GetValoreRiga(ListView1, cnt), vbTab & "+", vbTab & "'+")
                If cnt > 1 Then strTabella = strTabella & vbCrLf
                strTabella = strTabella & strTmp
            Next
            
            If strTabella <> "" Then
                Clipboard.Clear
                Clipboard.SetText strTabella, vbCFText
                Call ControllaCheck(ListView1)
            End If

    End Select

    TotaleRigheListView ListView1
    Set itmX = Nothing

    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione ClickPopUp."

End Sub

Private Sub ListView1_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result
    Dim i As Integer
    Dim nCampi As Integer
    Dim CancellaListView As Boolean
    
    If ckAggiungi.value = 1 Then
        CancellaListView = False
    Else
        CancellaListView = True
    End If
    
    If data.GetFormat(vbCFFiles) Then
        For i = 1 To 1 'Data.Files.count
            FileDaAprire = (data.Files(i))
            Screen.MousePointer = vbHourglass
            DoEvents
            If ImportaDati(Me.hwnd, ListView1, FileDaAprire, , True, CancellaListView, , , , "FileDatiPOI") = False Then Result = MsgBox("Errore di importazione del file!", vbOKOnly)
            Screen.MousePointer = vbDefault
        Next
    End If

End Sub

Private Sub cmdImportaPosizione_Click()
    Dim itemo As ListItem
    Set itemo = ListView1.SelectedItem
    Dim strTmp As String
    
    cmdImportaPosizione.Enabled = False
    
    If ck1.value = 1 Then ' Solo riga corrente
        If itemo Is Nothing Then
        Else
            CurrRow = itemo.index
            strTmp = WebBrowser1URL
            WebBrowser1.Navigate2 strTmp
            If Var(DebugMode).Valore = 1 Then WriteLog "WebBrowser1.Navigate2: " & strTmp, "Debug"
        End If
        
    Else ' Tutte le righe
        If ckDaCorrRiga.value = 1 Then
            CurrRow = itemo.index
        Else
            CurrRow = 1
        End If
        sURL = "about:blank"
        sURL = WebBrowser1URL
        WebBrowser1.Navigate2 sURL
        If Var(DebugMode).Valore = 1 Then WriteLog "WebBrowser1.Navigate2: " & sURL, "Debug"
    End If

    cmdImportaPosizione.Enabled = True
    
    Set itemo = Nothing

End Sub

Private Function WebBrowser1URL()
    Dim NomeImport As String
    
    '"ViaMichelin.it"
    '"Mappe.Libero.it"
    '"TuttoCittà.it"
    '"Google.com"
    '"MultiMap.com"
    
    CicliDocComplete = 1
    cmdImpQuestaPag.Enabled = False
    NomeImport = Trim$(cmbSitoCoordinate.Text)
    ValScript.oOK = False
    
    Select Case NomeImport
        Case "* ViaMichelin.it"
            SitoImportazione = "http://www.viamichelin.it"
            sURL = "http://www.viamichelin.it/viamichelin/ita/dyn/controller/mapPerformPage?"
            sURL = sURL & "strAddress=" & Replace(ListView1.ListItems(CurrRow).SubItems(2), " ", "+") & "&"
            sURL = sURL & "strCP=" & Replace(ListView1.ListItems(CurrRow).SubItems(3), " ", "+") & "&"
            sURL = sURL & "strLocation=" & Replace(ListView1.ListItems(CurrRow).SubItems(4), " ", "+")
        
        Case "* Mappe.Libero.it"
            SitoImportazione = "http://mappe.libero.it"
            'http://mappe.libero.it/tcolnew/index_libero.html#sez=1015&com=Milano%20(MI)&ind=Via%20Dante&nc=2
            sURL = "http://mappe.libero.it/mappe/search.jsp?"
            sURL = sURL & "country=IT&"
            sURL = sURL & "city=" & Replace(ListView1.ListItems(CurrRow).SubItems(4), " ", "+") & "&"
            sURL = sURL & "zipCode=" & Replace(ListView1.ListItems(CurrRow).SubItems(3), " ", "+") & "&"
            sURL = sURL & "street=" & Replace(ListView1.ListItems(CurrRow).SubItems(2), " ", "+")
            
        Case "* TuttoCittà.it"
            SitoImportazione = "http://www.tuttocitta.it"
            'http://www.tuttocitta.it/tcoln/action?msez=500&com=Milano&in=via%20Dante&nc=2
            sURL = "http://www.tuttocitta.it/tcoln/action?"
            sURL = sURL & "msez=500&"
            sURL = sURL & "com=" & Replace(ListView1.ListItems(CurrRow).SubItems(4), " ", "%20") & "&"
            sURL = sURL & "in=" & Replace(ListView1.ListItems(CurrRow).SubItems(2), " ", "%20") & "&"
            sURL = sURL & "nc=&"
            
        Case "* Google.com"
            CicliDocComplete = 2
            SitoImportazione = "http://www.google.com"
            sURL = "http://www.google.com/maps?f=q&hl=en&q="
            sURL = sURL & Replace(ListView1.ListItems(CurrRow).SubItems(2), " ", "+")
            sURL = sURL & "," & Replace(ListView1.ListItems(CurrRow).SubItems(3), " ", "+")
            sURL = sURL & "," & Replace(ListView1.ListItems(CurrRow).SubItems(4), " ", "+")
            sURL = sURL & ",italy"
            
        Case "* MultiMap.com"
            SitoImportazione = "http://www.multimap.com"
            'http://www.multimap.com/map/places.cgi?client=public&lang=&advanced=&db=IT&overviewmap=&keepicon=true&addr2=via+Dante+2&addr3=Milano&pc=20123
            sURL = "http://www.multimap.com/map/places.cgi?client=public&lang=&advanced=&db=IT&overviewmap=&keepicon=true&"
            sURL = sURL & "addr2=" & Replace(ListView1.ListItems(CurrRow).SubItems(2), " ", "+") & "&"
            sURL = sURL & "addr3=" & Replace(ListView1.ListItems(CurrRow).SubItems(4), " ", "+") & "&"
            sURL = sURL & "pc=" & Replace(ListView1.ListItems(CurrRow).SubItems(3), " ", "+")
        
        Case Else
            ' Se non c'è l'asterisco allora si sta utilizzando uno script
            If Left$(NomeImport, 1) <> "*" Then
                ValScript.iIndirizzo = ListView1.ListItems(CurrRow).SubItems(2)
                ValScript.iCitta = ListView1.ListItems(CurrRow).SubItems(4)
                ValScript.iCap = ListView1.ListItems(CurrRow).SubItems(3)
                Set ValScript.iWebBrowser = WebBrowser1
                
                LeggiScript Var(CartellaScript).Valore & "\" & NomeImport & ".txt"
                
                sURL = ValScript.iURL
                
            End If
            
    End Select
    
Esci:
    If sURL = "" Then sURL = "about:blank"
    
    LastWebNav = sURL
    WebBrowser1URL = sURL

End Function

Private Sub txtCognome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdRicercaSequenziale_Click
    End If
End Sub

Private Sub txtIndNaviga_GotFocus()
    
    ' Seleziono tutto il testo
    txtIndNaviga.SelStart = 0
    txtIndNaviga.SelLength = Len(txtIndNaviga.Text)
    
End Sub

Private Sub txtIndNaviga_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        cmdNaviga_Click
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdRicercaSequenziale_Click
    End If
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Static LastCurrRow As Long
    Dim tmpCurrRow As Long
    Dim strPatchNomeFileAperto As String
    Dim strTmp As String
    
    tmpCurrRow = CurrRow - 1
    
    ' Per il salvataggio automatico del file
    If Var(SalvataggioBackup).Valore = True Then
    If ((tmpCurrRow <> LastCurrRow And tmpCurrRow > 0 And (0 = tmpCurrRow Mod Var(SalvaOgni).Valore)) Or (tmpCurrRow = ListView1.ListItems.count)) And ListView1.ListItems.count > 0 Then
        Screen.MousePointer = vbHourglass
        
        If PatchNomeFileAperto <> "" Then
            strPatchNomeFileAperto = PatchNomeFileAperto
        Else
            strPatchNomeFileAperto = App.path & "\SalvataggioAutomatico.bak.rmk"
        End If
        
        strTmp = Left$(strPatchNomeFileAperto, Len(strPatchNomeFileAperto) - 4) & ".bak.rmk"
        Call ExportaDati(Me.hwnd, strTmp, , , False, , Var(CampiRmkFile).Valore & "FileDatiPOI" & ";;" & Now, ListView1)
        PatchNomeFileAperto = strPatchNomeFileAperto
        Screen.MousePointer = vbDefault
        DoEvents
    End If
    End If
    
    LastCurrRow = tmpCurrRow
    
    cmdImportaPosizione.Enabled = False

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, url As Variant)
    Dim NomeImport As String
    Dim i As Integer
    Dim doc As Object
    Dim scratch As String
    Dim pos1, pos2 As Integer
    Dim Latitudine, Longitudine As String
    Dim Comodo1, Comodo2 As String
    Dim vGradi
    Dim trovato As Boolean
    Dim CurrRowCambiata As Boolean
    Dim cntTm As Integer
    Dim NavigaVerso As String
    Dim arrStr() As String
    Dim scriviCoo As Boolean
    
    If bStop = True Then GoTo Esci
    
    CurrRowCambiata = False
    scriviCoo = False
    
    'Debug.Print sURL
    'Debug.Print URL
    'Debug.Print "------"
    'MsgBox sURL & vbNewLine & "-------" & vbNewLine & URL
    
    '"ViaMichelin.it"
    '"Mappe.Libero.it"
    '"TuttoCittà.it"
    '"Google.com"
    '"MultiMap.com"

    NomeImport = Trim$(cmbSitoCoordinate.Text)

    If ((sURL = url Or ck1.value = 1) And _
        NomeImport = "* ViaMichelin.it") Or _
        NomeImport = "* Mappe.Libero.it" Or _
        ((Left$(sURL, 61) = Left$(url, 61) Or ck1.value = 1) And NomeImport = "* TuttoCittà.it") Or _
        NomeImport = "* Google.com" Or _
        (NomeImport = "* MultiMap.com" And Left$(url, 27) = Left$(sURL, 27)) Or _
        Left$(NomeImport, 1) <> "*" And InStr(1, url, ValScript.iInStrRetURL) <> 0 _
        Then
        
        ' Aggiunto per eliminare l'errore sull'importazione con TuttoCittà
        'If NomeImport = "* TuttoCittà.it" Then
        '    DoEvents
        '    DoEvents
        'End If
        '
         
        If CurrRow >= 1 Then
            Set doc = WebBrowser1.Document
            trovato = False
            
            ' Cerco su ViaMichelin -------------------------------------------------------------------------------------------------------------------
            If NomeImport = "* ViaMichelin.it" And url <> "about:blank" Then
                For i = 0 To doc.scripts.Length - 1
                    scratch = doc.scripts.Item(i).Text
                    pos1 = InStr(1, scratch, "lati = ")
                    If trovato = False Then
                            If pos1 > 0 Then
                                pos2 = InStr(pos1, scratch, ";")
                                If pos2 > pos1 Then
                                    Latitudine = Mid(scratch, pos1 + 7, pos2 - (pos1 + 7))
                                Else
                                    Latitudine = "0.0"
                                End If
                            Else
                                Latitudine = "0.0"
                            End If
                    
                            pos1 = InStr(1, scratch, "longi = ")
                    
                            If pos1 > 0 Then
                                pos2 = InStr(pos1, scratch, ";")
                                If pos2 > pos1 Then
                                    Longitudine = Mid(scratch, pos1 + 7, pos2 - (pos1 + 7))
                                    trovato = True
                                Else
                                    Longitudine = "0.0"
                                End If
                            Else
                                Longitudine = "0.0"
                            End If
                    End If
                Next

            ' Cerco su MappeLibero -------------------------------------------------------------------------------------------------------------------
            ElseIf NomeImport = "* Mappe.Libero.it" And url <> "about:blank" Then
                For i = 0 To doc.body.All.Length - 1
                    scratch = doc.body.All.Item(i).innerText
                    pos1 = InStr(1, scratch, "Lat.: ")
                    If trovato = False Then
                        If pos1 > 0 Then
                            pos2 = InStr(pos1 + 6, scratch, """ N")
                            Comodo1 = Mid(scratch, pos1 + 6, pos2 - (pos1 + 6))
                            pos1 = InStr(1, scratch, "Long.: ")
                            pos2 = InStr(pos1 + 7, scratch, """ E")
                            ' Controllo che i calcoli utilizzati per trovare Comodo 2 non producano valori negativi che mandano in errore la funzione Mid
                            If (pos1 + 7) > 0 And (pos2 - (pos1 + 7)) > 0 Then
                                Comodo2 = Mid(scratch, pos1 + 7, pos2 - (pos1 + 7))
                                Comodo1 = Replace$(Comodo1, "°", "'")
                                vGradi = Split(Comodo1, "'")
                                Latitudine = Replace$(CStr(Round(vGradi(0) + vGradi(1) / 60 + vGradi(2) / 3600, 6)), ",", ".")
                                Comodo2 = Replace$(Comodo2, "°", "'")
                                vGradi = Split(Comodo2, "'")
                                Longitudine = Replace$(CStr(Round(vGradi(0) + vGradi(1) / 60 + vGradi(2) / 3600, 6)), ",", ".")
                            Else
                                Latitudine = "0.0"
                                Longitudine = "0.0"
                            End If
                            trovato = True
                        Else
                            Latitudine = "0.0"
                            Longitudine = "0.0"
                        End If
                    End If
                Next
            
            ' Cerco su TuttoCittà -------------------------------------------------------------------------------------------------------------------
            ElseIf NomeImport = "* TuttoCittà.it" And url <> "about:blank" Then
                If trovato = False Then
                    Dim X As New DOMDocument
                    X.async = False
                    X.setProperty "SelectionLanguage", "XPath"
                    
                    ' Prendo il codice html della pagina
                    X.loadXML (doc.documentElement.outerHTML)
                    
                    If X.xml <> "" And InStr(X.xml, "CODE") <> 0 Then
                        Comodo1 = X.selectNodes("HTML/BODY").Item(0).selectSingleNode("CODE").Text
                        'dove <code>0</code>
                        '0 = ricerca esatta
                        '5 = centro via
                        '3 = non trovato
                        'etc...
                        If Comodo1 = 0 Or Comodo1 = 5 Then
                            Longitudine = X.selectNodes("/HTML/BODY").Item(0).selectSingleNode("X").Text
                            Latitudine = X.selectNodes("/HTML/BODY").Item(0).selectSingleNode("Y").Text
                            'Comodo1 = x.selectNodes("/HTML/BODY").Item(0).selectSingleNode("Z").Text
                            'Comodo1 = x.selectNodes("/HTML/BODY").Item(0).selectSingleNode("FRAZ").Text
                            'Comodo1 = x.selectNodes("/HTML/BODY").Item(0).selectSingleNode("C").Text
                            'Comodo1 = x.selectNodes("/HTML/BODY").Item(0).selectSingleNode("COM").Text
                            'Comodo1 = x.selectNodes("/HTML/BODY").Item(0).selectSingleNode("PROV").Text
                            'Comodo1 = x.selectNodes("/HTML/BODY").Item(0).selectSingleNode("TOPO").Text
                            'Comodo1 = x.selectNodes("/HTML/BODY").Item(0).selectSingleNode("CIV").Text
                            'Comodo1 = x.selectNodes("/HTML/BODY").Item(0).selectSingleNode("CDLOC").Text
                            'Comodo1 = x.selectNodes("/HTML/BODY").Item(0).selectSingleNode("REG").Text
                            'Comodo1 = x.selectNodes("/html/body").Item(0).selectSingleNode("UL").Text
                            trovato = True
                            
                        ElseIf Comodo1 = 3 Or Comodo1 = 4 Then
                            Longitudine = "0.0"
                            Latitudine = "0.0"
                        
                        ElseIf Comodo1 = 10 Then
                            ' Ci possono essere alternative da cercare
                            Longitudine = ""
                            Latitudine = ""
                        
                        End If
                        
                    Else
                        ' La pagine web potrebbe non avere trovato il risultato e suggerire alternative
                        'MsgBox doc.documentElement.outerHTML
                        Longitudine = "0.0"
                        Latitudine = "0.0"
                    End If
                                    
                    Set X = Nothing
                End If

            ' Cerco su Google -------------------------------------------------------------------------------------------------------------------
            ElseIf NomeImport = "* Google.com" Then
                ' Le coordinate si trovano in un link contenuto nella pagina web
                ' http://www.google.com/maps?f=q&hl=en&q=Via+Zamenhof+62,54036,Marina+di+Carrara,italy&ie=UTF8&ll=44.051077,10.042019&spn=0.011721,0.049524&z=14&iwloc=addr&om=1
                DoEvents
                arrStr = EnumLink(WebBrowser1, "LEFT      " & "http://www.google.com/maps?", "&ll=")
                Comodo1 = GetValoreFromURL(arrStr(0), "ll")
                If Comodo1 <> "" Then
                    LatLonFromString Comodo1, Latitudine, Longitudine, "Latitudine,Longitudine"
                End If
            
            ' Cerco su MultiMap -------------------------------------------------------------------------------------------------------------------
            ElseIf NomeImport = "* MultiMap.com" Then
                ' Le coordinate si trovano nella URL basta solo prenderle :)
                ' http://www.multimap.com/maps/?&t=l&map=44.01485,10.12846|17|4&loc=IT:44.01485:10.12846:17
                LatLonFromString url, Latitudine, Longitudine, "&map=Latitudine,Longitudine|"
                DoEvents
                DoEvents
                
            ' Se non c'è l'asterisco allora si sta utilizzando uno script -------------------------------------------------------------------------------------------------------------------
            ElseIf Left$(NomeImport, 1) <> "*" Then
                ' Le coordinate si trovano nella URL basta solo prenderle :)
                ' http://www.multimap.com/maps/?&t=l&map=44.01485,10.12846|17|4&loc=IT:44.01485:10.12846:17
                LatLonFromString url, Latitudine, Longitudine, "&map=Latitudine,Longitudine|"
                DoEvents
                DoEvents
                
            End If
            
            
            If ck1.value = 0 And CurrRow > 1 Then
                CurrRow = CurrRow - 1
                CurrRowCambiata = True
            End If
            
            ' Aggiungo i dati
            ListView1.ListItems(CurrRow).SubItems(8) = Replace$(Latitudine, ".", ",")
            ListView1.ListItems(CurrRow).SubItems(9) = Replace$(Longitudine, ".", ",")
            If Latitudine <> "0.0" And Longitudine <> "0.0" Then
                ListView1.ListItems(CurrRow).SubItems(10) = SitoImportazione
            Else
                ListView1.ListItems(CurrRow).SubItems(10) = ""
            End If
            
            If ck1.value = 0 And CurrRowCambiata = True Then CurrRow = CurrRow + 1
        End If
        
        If CurrRow > ListView1.ListItems.count Or ListView1.ListItems.count = 0 Then
            cmdImportaPosizione.Enabled = True
            cmdImpQuestaPag.Enabled = True
            Exit Sub
        End If

        ' Carico solo i valori nelle celle con 0.0.......
        If ckSoloNulle.value = 1 And ck1.value = 0 And CurrRow <> 0 And ListView1.ListItems.count <> 0 Then
            For cntTm = 1 To ListView1.ListItems.count
                If CurrRow > ListView1.ListItems.count Then Exit For
                If ListView1.ListItems(CurrRow).SubItems(8) = "" Or ListView1.ListItems(CurrRow).SubItems(8) = "" Then Exit For
                Do While (ListView1.ListItems(CurrRow).SubItems(8) <> "0,0" And ListView1.ListItems(CurrRow).SubItems(8) <> "0,0")
                    If CurrRow >= ListView1.ListItems.count Then Exit Do
                    CurrRow = CurrRow + 1
                Loop
            Next
        End If
        ' Carico solo i valori nelle celle vuote.......
        If ckSoloVuote.value = 1 And ck1.value = 0 And CurrRow <> 0 And ListView1.ListItems.count <> 0 Then
            For cntTm = 1 To ListView1.ListItems.count
                If CurrRow > ListView1.ListItems.count Then Exit For
                If ListView1.ListItems(CurrRow).SubItems(8) = "0,0" Or ListView1.ListItems(CurrRow).SubItems(8) = "0,0" Then Exit For
                Do While (ListView1.ListItems(CurrRow).SubItems(8) <> "" And ListView1.ListItems(CurrRow).SubItems(8) <> "")
                    If CurrRow >= ListView1.ListItems.count Then Exit Do
                    CurrRow = CurrRow + 1
                Loop
            Next
        End If

        If CurrRow <> 0 And CurrRow <= ListView1.ListItems.count And ck1.value = 0 Then
        
            If NomeImport = "* MultiMap.com" Then
                CicliDocComplete = 1
            End If
            
            If CicliDocComplete <= 1 Then
                NavigaVerso = WebBrowser1URL
                ' Seleziono la riga nella ListView
                Call SelezionaRigaListView(ListView1, CurrRow - 1)
                CurrRow = CurrRow + 1
                WebBrowser1.Navigate2 NavigaVerso
            Else
                CicliDocComplete = CicliDocComplete - 1
            End If
        End If
        
    End If

Esci:
    
    Set doc = Nothing
    cmdImpQuestaPag.Enabled = True
    cmdImportaPosizione.Enabled = True
    DoEvents

End Sub

Private Sub cmdRicercaSequenziale_Click()
    Dim Cognome As String
    Dim Nome As String
    Dim Dove As String
        
    If txtCognome.Text <> "Cognome o Nome Azienda" Then Cognome = txtCognome.Text
    If txtNome.Text <> "Nome" And optSito(4).value = True Then Nome = txtNome.Text
    If txtDove.Text <> "Dove" Then Dove = txtDove.Text
    
    If Cognome = "" Then
        txtCognome.Text = ""
        txtCognome.SetFocus
        Exit Sub
    End If
    
    If optSito(3).value = True Then
        ' Pagine Gialle
        WebBrowser2.Navigate2 ("http://www.paginegialle.it/pg/cgi/pgsearch.cgi?btt=1&ts=1&l=1&cb=0&ind=&nc=&qs=" & Cognome & "&dv=" & Dove & "&x=0&y=0")
        DoEvents
    ElseIf optSito(4).value = True Then
        ' Pagine Bianche
        WebBrowser2.Navigate2 ("http://www.paginebianche.it/execute.cgi?btt=1&tl=2&tr=101&tc=&cb=&x=0&y=0&tq=2&qs=" & Cognome & "&qsn=" & Nome & "&dv=" & Dove & "&ind=&nc=")
        DoEvents
    End If
    
End Sub

Private Sub cmdImporta_Click()
    Dim itmX As ListItem
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Y As Integer
    Dim cntDel As Long
    Dim cntRiga As Long
    Dim scratch As String
    Dim vWord, vWord2, vWord3, vWord4
    Dim sDescr As String
    Dim sCap As String
    Dim sNum
    Dim sAddr As String
    Dim sCity As String
    Dim sProv As String
    Dim sTel As String
    Dim sCat As String
    Dim pos1 As Long
    Dim pos2 As Long
    Dim GoToWeb As String
    Dim RicercaPerQuartiere As Boolean
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    RicercaPerQuartiere = False
    cmdImporta.Enabled = False
    cmdStop.Enabled = True

    sLinkNext = ""
    
    If doc Is Nothing Then
        MsgBox "Documento Web non impostato.", vbInformation, App.ProductName
        GoTo Esci
    End If
    
    For i = 0 To doc.body.All.Length - 1
        scratch = doc.body.All.Item(i).innerHTML
        
        If sLinkNext = "" Then
        If optSito(3) = True Then
        pos1 = InStr(1, scratch, "succ »")
            If pos1 > 0 Then
            pos1 = InStr(pos1, scratch, "href=")
            pos2 = InStr(pos1 + 6, scratch, """")
                If pos2 > 0 Then
                    sLinkNext = Replace$(Mid(scratch, pos1 + 6, pos2 - (pos1 + 6)), "&amp;", "&")
                End If
            End If
        Else
        pos1 = InStr(1, scratch, "succ »")
        If pos1 > 0 Then
        pos1 = InStr(pos1 - 250, scratch, "/execute.cgi?be")
            If pos1 > 0 Then
            pos2 = InStr(pos1, scratch, """")
                If pos2 > 0 Then
                    sLinkNext = Replace$(Mid(scratch, pos1, pos2 - pos1), "&amp;", "&")
                End If
            End If
            End If
        End If
        End If

        If (Left(scratch, 23) = "<TD width=25 rowSpan=8>" And optSito(3) = True) Or (Left(scratch, 24) = "<TR>" & vbCrLf & "<TD class=pallino>" And optSito(4) = True) Then
            scratch = doc.body.All.Item(i).innerText
            
            ' Tolgo i caratteri iniziali che si trovano nella stringa che si ottiene nella ricerca per quartiere
            While ASC(Left$(scratch, 1)) = 32 Or ASC(Left$(scratch, 1)) = 13 Or ASC(Left$(scratch, 1)) = 10
                scratch = Right$(scratch, Len(scratch) - 1)
                RicercaPerQuartiere = True
            Wend
            
            vWord = Split(scratch, vbCrLf)
            
            ' Se..... cancello la seconda riga dell'array
            If RicercaPerQuartiere = True Then
                For Y = 1 To UBound(vWord) - 1
                    vWord(Y) = vWord(Y + 1)
                Next
                ReDim Preserve vWord(UBound(vWord) - 1)
            End If
            
            ' La prima riga corrisponde alla descrizione
            sDescr = vWord(0)
            
CercaCap:
            ' Scorro tutto l'array per eliminare le righe che non iniziano per il CAP dopo la prima
            ' Se trovo il CAP lascio l'array come sta ed esco
            cntDel = 0
            For Y = 1 To UBound(vWord) - 1
                If (IsNumeric(Left$(vWord(Y), 5)) = False) Or (cntDel > 0) Then
                    ' Se l'inizio della riga non è il CAP cancello la riga
                    vWord(Y) = vWord(Y + 1)
                    cntDel = cntDel + 1
                Else
                    Exit For
                End If
            Next
            If cntDel > 0 Then
                ReDim Preserve vWord(UBound(vWord) - 1)
                ' Ripeto il controllo
                GoTo CercaCap
            End If
            '
            vWord2 = Split(vWord(1), "-")
            vWord3 = Split(vWord2(0), " ")
            sCap = vWord3(0)
            sCity = ""
            
            For j = 1 To UBound(vWord3) - 2
                sCity = sCity & vWord3(j) & " "
            Next
        
             sCity = Trim(sCity)
             scratch = vWord3(UBound(vWord3) - 1)
             scratch = Replace(scratch, "(", " ")
             scratch = Replace(scratch, ")", " ")
             sProv = Trim(scratch)
             sNum = ""
             sAddr = "non trovato"
             
             If UBound(vWord2) > 0 Then
                vWord4 = Split(vWord2(1), ",")
                If UBound(vWord4) > 0 Then
                If optSito(3) = True Then
                sNum = vWord4(0)
                sAddr = vWord4(1) & " " & sNum
                Else
                sNum = vWord4(1)
                sAddr = vWord4(0) & " " & sNum
                End If
                Else
                sAddr = vWord4(0)
                End If
             End If
             
             sTel = ""
             
             If UBound(vWord) > 1 Then
                 If optSito(4).value = True Then
                     sTel = vWord(UBound(vWord) - 1)
                 Else
                     sTel = vWord(2)
                 End If
             End If
             
             If sTel <> "" Then
                sTel = Replace$(sTel, "•", ",")
                sTel = Replace$(sTel, "-", ",")
                vWord4 = Split(sTel, ",")
                sTel = vWord4(0)
                sTel = Replace$(sTel, "fax:", "")
                sTel = Replace$(sTel, "tel:", "")
                vWord4 = Split(Trim(sTel), " ")
                
                If UBound(vWord4) > 0 Then
                    sTel = "" & Trim(vWord4(0)) & Trim(vWord4(1))
                Else
                    sTel = "" & Trim(vWord4(0))
                End If
             End If
             
             sCat = ""
             
             If optSito(3) = True Then
                k = 1
                While k < 99
                    scratch = doc.body.All.Item(i + k).innerText
                    If InStr(1, scratch, "Categoria") > 0 Then
                    vWord4 = Split(scratch, "»")
                    
                    'If UBound(vWord4) > 0 Then
                        sCat = Trim$(vWord4(1))
                    'Else
                    '    sCat = "Non trovato"
                    'End If
                    
                    k = 100
                    Else
                    k = k + 1
                    End If
                Wend
             End If
             
             ' Se nell'indirizzo mancano alcuni dati correggo i dati importati male
             If Not IsNumeric(sCap) And sCity = "" Then
                sCity = sCap
                sCap = ""
             End If
             
            Set itmX = ListView1.ListItems.Add(, , ElaboraNumeroRiga(ListView1.ListItems.count + 1))
             itmX.SubItems(1) = ElaboraNome(sDescr)
             itmX.SubItems(2) = ElaboraNome(sAddr)
             itmX.SubItems(3) = sCap
             itmX.SubItems(4) = ElaboraNome(sCity)
             itmX.SubItems(5) = UCase(sProv)
             itmX.SubItems(6) = sTel
             itmX.SubItems(7) = ElaboraNome(sCat)
             cntRiga = cntRiga + 1
         End If
    Next
    
    Call SelezionaRigaListView(ListView1, ListView1.ListItems.count, True)
    
    ' Fermo il ciclo
    If bStop = True Then
        cmdStop.Enabled = False
        bStop = False
        GoTo Esci
    End If
    
    If sLinkNext <> "" And cekPagine.value = 1 Then
        If optSito(3) = True Then GoToWeb = "http://www.paginegialle.it" & sLinkNext
        If optSito(4) = True Then GoToWeb = "http://www.paginebianche.it" & sLinkNext
        WebBrowser2.Navigate2 GoToWeb
        
    ElseIf sLinkNext = "" And cekPagine.value = 1 And cmdRicercaSequenziale.Visible = True Then
        ' Il ciclo di importazione di tutte le pagine è finito e quindi cambio Regione - Provincia
        If cmbRegione.Enabled = True And cmbRegione.ListIndex < cmbRegione.ListCount - 1 Then
            cmbRegione.ListIndex = (cmbRegione.ListIndex + 1)
            DoEvents
            cmdRicercaSequenziale_Click
            
        ElseIf cmbProvincia.Enabled = True And cmbProvincia.ListIndex < cmbProvincia.ListCount - 1 Then
            cmbProvincia.ListIndex = (cmbProvincia.ListIndex + 1)
            DoEvents
            cmdRicercaSequenziale_Click
        End If
        
    End If

Esci:
    cmdImporta.Enabled = True
    Set itmX = Nothing
    Set doc = Nothing
    
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione cmdImporta_Click." & vbNewLine & scratch & vbNewLine & vbNewLine
    GoTo Esci

End Sub

Private Sub WebBrowser2_BeforeNavigate2(ByVal pDisp As Object, url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    
    cmdImporta.Enabled = False

End Sub

Private Sub WebBrowser2_DocumentComplete(ByVal pDisp As Object, url As Variant)
    Set doc = WebBrowser2.Document
    
    If (sLinkNext <> "" And cekPagine.value = 1) And _
        ((Left(url, 46) = "http://www.paginegialle.it/pg/cgi/pgsearch.cgi" And optSito(3) = True) Or _
        (Left(url, 42) = "http://www.paginebianche.it/execute.cgi?be" And optSito(4) = True)) Then
        DoEvents
        cmdImporta_Click
        
    ElseIf (sLinkNext = "" And cekPagine.value = 1 And cmdRicercaSequenziale.Visible = True) And _
        ((Left(url, 50) = "http://www.paginegialle.it/pg/cgi/pgsearch.cgi?btt" And optSito(3) = True) Or _
        (Left(url, 43) = "http://www.paginebianche.it/execute.cgi?btt" And optSito(4) = True)) Then
        DoEvents
        cmdImporta_Click

    End If

    cmdImporta.Enabled = True

End Sub

Private Sub optSito_Click(index As Integer)

    Select Case index
    
        Case Is = 3
            UserForm_Activate
            cmdImporta.Caption = "Importa da: " & optSito(3).Caption
            ckRegione.Enabled = True
            ckProvincia.Enabled = True
            txtNome.Visible = False
            
        Case Is = 4
            UserForm_Activate
            cmdImporta.Caption = "Importa da: " & optSito(4).Caption
            ckRegione.Enabled = True
            ckProvincia.Enabled = True
            txtNome.Visible = True
    End Select
    
End Sub

Private Sub UserForm_Activate()
    Dim GoToWeb As String

    sLinkNext = ""
    
    If optSito(3) = True Then
        GoToWeb = "http://www.paginegialle.it"
    ElseIf optSito(4) = True Then
        GoToWeb = "http://www.paginebianche.it"
    End If
    
    WebBrowser2.Navigate2 GoToWeb
    
End Sub

Private Sub BloccaWebBrowser(Optional Azione As Boolean = True)
    
    If Azione = True Then
    
    Else
    
    End If
    
End Sub

Private Function ElaboraNome(StringaTesto As String) As String
    Dim cnt As Integer
    Dim LenStr As Long
    Dim car As String
    Dim strResult As String
    Dim strTmp As String
    Dim Splitted() As String
    
    LenStr = Len(StringaTesto)
    
    For cnt = 1 To LenStr Step 1
        car = Mid$(StringaTesto, cnt, 1)
        Select Case car
            Case Is = "-", "_"
                strResult = strResult & " "
            Case Is = "|"
                strResult = strResult & " "
            Case Else
                strResult = strResult & car
        End Select
    Next
    
    ' Divido la stringa in più stringhe in base al delimitatore spazio
    Splitted = Split(strResult)
    strResult = ""
    
    For cnt = 0 To UBound(Splitted)
        ' La prima lettera maiuscola
        strTmp = PrimaMaiuscola(Trim$(Splitted(cnt)))
        If strTmp <> "" Then strResult = strResult & " " & strTmp
    Next
    
    ' Tolgo gli eventuali spazi di troppo
    strResult = Trim$(strResult)
    
    ElaboraNome = strResult
    
End Function

Private Sub Text1_GotFocus()
    Dim i As Long
    
    ' Selects the ListItem whose SubItem is being edited...
    ListView1.ListItems(m_iItem).Selected = True
    cmdEsci.Visible = False
    
    ' When Text1 gets the focus, clear all TabStop properties on all
    ' controls on the form. Ignore all errors, in case a control does
    ' not have the TabStop property.
    On Error Resume Next
    For i = 0 To Controls.count - 1   ' Use the Controls collection
       Controls(i).TabStop = False
    Next
    
    ' Controllo se il numero è divisibile per 2.....
    If m_iItem Mod 2 = 0 Then
        Text1.BackColor = vbGreenLemon
    Else
        Text1.BackColor = vbWhite
    End If
    
    NascondiScrollBar ListView1

End Sub

Private Sub text1_LostFocus()
    Dim i As Long

    ' When Text1 loses the focus, make the TabStop property True for all
    ' controls on the form. That restores the ability to tab between
    ' controls. Ignore all errors, in case a control does not have the
    ' TabStop property.
    On Error Resume Next
    
    For i = 0 To Controls.count - 1   ' Use the Controls collection
       Controls(i).TabStop = True
    Next
    
End Sub

Private Sub Text1_Change()
    ' If the TextBox is shown, size its width so that it's always a little
    ' longer than the length of its Text.
    If m_iItem And Text1.Width <= TextWidth(Text1) * 1.3 Then
        Text1.Width = TextWidth(Text1) * 1.35
    End If
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    TextKeyCodeZero = False
    
    If (KeyCode = vbKeyUp) Then
        KeyCode = 0
        TextKeyCodeZero = True
        Call HideTextBox(True)
        If ListView1.FullRowSelect = True Then ListView1.FullRowSelect = False
        EditListView True, m_iItem - 2, m_iSubItem

    ElseIf (KeyCode = vbKeyPageUp) Then
        KeyCode = 0
        TextKeyCodeZero = True
        Call HideTextBox(True)
        If ListView1.FullRowSelect = True Then ListView1.FullRowSelect = False
        EditListView True, m_iItem - 11, m_iSubItem

    ElseIf (KeyCode = vbKeyDown) Then
        KeyCode = 0
        TextKeyCodeZero = True
        Call HideTextBox(True)
        If ListView1.FullRowSelect = True Then ListView1.FullRowSelect = False
        EditListView True, m_iItem, m_iSubItem
    
    ElseIf (KeyCode = vbKeyPageDown) Then
        KeyCode = 0
        TextKeyCodeZero = True
        Call HideTextBox(True)
        If ListView1.FullRowSelect = True Then ListView1.FullRowSelect = False
        EditListView True, m_iItem + 9, m_iSubItem
    
    ElseIf (KeyCode = vbKeyTab) And (Shift = 0) Then
        KeyCode = 0
        TextKeyCodeZero = True
        Call HideTextBox(True)
        If ListView1.FullRowSelect = True Then ListView1.FullRowSelect = False
        EditListView True, m_iItem - 1, m_iSubItem + 1
    
    ElseIf (KeyCode = vbKeyTab) And (Shift = 1) Then
        KeyCode = 0
        TextKeyCodeZero = True
        Call HideTextBox(True)
        If ListView1.FullRowSelect = True Then ListView1.FullRowSelect = False
        EditListView True, m_iItem - 1, m_iSubItem - 1

    End If
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ' Update the SubItem text on the Enter key, cancel on the Escape Key.
    
    If (KeyAscii = vbKeyReturn) Then
        ListView1.FullRowSelect = True
        NascondiScrollBar ListView1, False
        Call HideTextBox(True)
        KeyAscii = 0
        
    ElseIf (KeyAscii = vbKeyEscape) Then
        ListView1.FullRowSelect = True
        NascondiScrollBar ListView1, False
        Call HideTextBox(False)
        KeyAscii = 0
                                
    ElseIf (KeyAscii = 124) Then ' Il simbolo |
        KeyAscii = 32 ' Lo sostituisco con lo spazio
        
    End If
    
    If TextKeyCodeZero = True Then KeyAscii = 0
      
End Sub

Friend Sub HideTextBox(fApplyChanges As Boolean)

    If fApplyChanges Then
        bEditet = True
        ListView1.ListItems(m_iItem).SubItems(m_iSubItem) = Trim$(Text1.Text)
        ' Cambio lo stato delle lablel nella form
        If frmSetupDescrizione.Visible = True Then
            frmSetupDescrizione.lblCambiaRigaCorrente.Caption = ""
            frmSetupDescrizione.lblCambiaRigaCorrente.Caption = ListView1.ListItems(m_iItem)
        End If
    Else
      ListView1.ListItems(m_iItem).SubItems(m_iSubItem) = Text1.Tag
    End If
    
    Call UnSubClass(m_hwndTB)
    Text1.Visible = False
    Text1 = ""
    CurrRow = m_iItem
    ListView1.SetFocus
    ListView1_Click
    'm_iItem = 0
    cmdEsci.Visible = True
  
End Sub

Public Sub CaricaScript()
    Dim sFile
    Dim cnt As Integer
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    sFile = fso.GetFolderFiles(Var(CartellaScript).Valore, ".txt", False)
    
    If IsArrayVuoto(sFile) = True Then GoTo Esci
    
    For cnt = 0 To UBound(sFile)
        cmbSitoCoordinate.AddItem " " & fso.RimuoviExt(sFile(cnt))
    Next

Esci:
    
    Exit Sub
    
Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione CaricaScript."

End Sub
