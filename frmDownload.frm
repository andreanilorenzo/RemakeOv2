VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDownload 
   Caption         =   "Download Manager"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtListBox 
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
      Height          =   345
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "txtListBox"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdUpdateFile 
      Caption         =   "&Aggiorna file"
      Height          =   375
      Left            =   2040
      OLEDropMode     =   1  'Manual
      TabIndex        =   35
      Top             =   60
      Width           =   1455
   End
   Begin MSComctlLib.ImageList imglRunSearch 
      Left            =   2160
      Top             =   2520
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
            Picture         =   "frmDownload.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":30F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":4DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":5114
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":526E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":53C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":5522
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":583C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":668E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":74E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":838A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":91DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":B98E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":BCA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":CAFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":CE14
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":E46E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":E788
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":EAA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":EDBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":F0D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":F9B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":1028A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":110DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":11F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":12D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":1365A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":144AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":161B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":164D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":17322
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":19AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":1C286
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":1DF90
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":20742
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2244C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":22D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":23B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":251D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":25AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":26386
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":271D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":28832
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":29684
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2999E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2A278
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2A592
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2A8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2ABC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2AD20
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2B03A
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2B354
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2B66E
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2B988
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2BCA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2BFBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2C116
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2C270
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":2C3CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSlider 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   135
      ScaleWidth      =   12855
      TabIndex        =   17
      Top             =   5280
      Width           =   12855
   End
   Begin VB.PictureBox picProxy 
      Height          =   975
      Left            =   3600
      ScaleHeight     =   915
      ScaleWidth      =   6435
      TabIndex        =   23
      Top             =   120
      Width           =   6495
      Begin VB.CheckBox chkUseProxy 
         Caption         =   "U&sa Proxy"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtProxyServer 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   27
         Text            =   "txtProxyServer"
         Top             =   120
         Width           =   3975
      End
      Begin VB.TextBox txtProxyPort 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   26
         Text            =   "txtProxyPort"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txtProxyUserName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Text            =   "txtProxyUserName"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtProxyPassword 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   24
         Text            =   "txtProxyPassword"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Proxy Ser&ver: "
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Porta:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5280
         TabIndex        =   30
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Us&er Name:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1320
         TabIndex        =   29
         Top             =   540
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "P&assword:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3960
         TabIndex        =   28
         Top             =   540
         Width           =   735
      End
   End
   Begin VB.PictureBox picComm 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      ScaleHeight     =   330
      ScaleWidth      =   9975
      TabIndex        =   18
      Top             =   1320
      Width           =   9975
      Begin Remakeov2.RMComboView cmbTipoPDI 
         Height          =   315
         Left            =   2640
         TabIndex        =   36
         Top             =   0
         Width           =   2055
         _ExtentX        =   4471
         _ExtentY        =   661
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
         BorderStyle     =   3
         ColumnResize    =   -1  'True
         ColumnSort      =   -1  'True
         DisplayEllipsis =   -1  'True
         DropDownAutoWidth=   -1  'True
         DropDownItemsVisible=   20
         HotButtonBackColor=   0
      End
      Begin VB.CheckBox cekFiltraPerMappa 
         Height          =   195
         Left            =   9120
         TabIndex        =   33
         ToolTipText     =   $"frmDownload.frx":2C524
         Top             =   60
         Value           =   1  'Checked
         Width           =   135
      End
      Begin VB.CommandButton cmdScaricaElenco 
         Caption         =   "< Scarica elenco file XML       "
         Height          =   315
         Index           =   0
         Left            =   7080
         TabIndex        =   22
         ToolTipText     =   "Scarica l'elenco dei PDI disponibili su XML"
         Top             =   0
         Width           =   2415
      End
      Begin VB.CommandButton cmdScaricaElenco 
         Caption         =   "< Scarica elenco file WEB"
         Height          =   315
         Index           =   1
         Left            =   4800
         TabIndex        =   21
         ToolTipText     =   "Scarica l'elenco dei PDI della categoria selezionata"
         Top             =   0
         Width           =   2175
      End
      Begin VB.TextBox txtFiltra 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2160
         TabIndex        =   20
         Text            =   "PDI"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdScaricaTipiPDI 
         Caption         =   "Aggiorna categorie PDI >"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Aggiorna le categorie dei PDI disponibili sul sito www.poigps.com"
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox picShadow 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   8175
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.CheckBox ckBrowser 
      Caption         =   "Visualizza Browser"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox picCommand 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   120
      ScaleHeight     =   405
      ScaleWidth      =   10065
      TabIndex        =   8
      Top             =   4800
      Width           =   10065
      Begin VB.CheckBox cekRemake 
         Caption         =   "cekRemake"
         Height          =   195
         Left            =   2400
         TabIndex        =   34
         ToolTipText     =   "Dopo aver scaricato i file apri automaticamente la finestra per effettuare il trattamento"
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton cmdScaricaFile 
         Caption         =   "&Scarica file selezionati"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   25
         Width           =   3015
      End
      Begin VB.CommandButton cmdEsci 
         Cancel          =   -1  'True
         Caption         =   "&Esci  [Esc]"
         Height          =   375
         Left            =   8160
         TabIndex        =   11
         Top             =   25
         Width           =   1815
      End
      Begin VB.CommandButton cmdScaricaFile 
         Caption         =   "&Scarica tutti i file"
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   10
         Top             =   25
         Width           =   3015
      End
      Begin VB.CommandButton cmdApriCartellaFile 
         Caption         =   "&Apri cartella file"
         Height          =   375
         Left            =   6240
         TabIndex        =   9
         Top             =   25
         Width           =   1815
      End
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
      Left            =   1440
      ScaleHeight     =   990
      ScaleWidth      =   2295
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   2325
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "txtUserName"
      ToolTipText     =   "Utilizzata solo in modalità XML: User Name"
      Top             =   480
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3075
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   5424
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
   Begin SHDocVwCtl.WebBrowser wbrSource 
      Height          =   735
      Left            =   9120
      TabIndex        =   13
      Top             =   2280
      Width           =   1455
      ExtentX         =   2566
      ExtentY         =   1296
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   7815
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstLog 
      Height          =   2400
      Left            =   0
      TabIndex        =   6
      Top             =   5400
      Width           =   10215
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Text            =   "txtPassword"
      ToolTipText     =   "Utilizzata solo in modalità XML: Password"
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Pass&word:"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "User &Name:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   960
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tyDatiFile
    Descrizione As String
    url As String
    upURL As String
    data As String
    Nome As String
    Gruppo As String
    Formato As String
    Mappa As String
    Immagine As String
    Note As String
    Modalità As String
End Type

Private Type tyTipoPDI
    NomeMappa As String
    PaginaWeb As String
End Type

Public Enum scda
    Noo = 0
    Web = 1
    xml = 2
End Enum

Dim Dragging As Boolean ' Flag that tells us if the slider is moving

Public elencoDa As scda
Private arrElencoFile() As tyDatiFile
Private lbLog As clsListBox
Private UsaPassword As Boolean
Private FileScaricato As Boolean
Private SoloChecked As Boolean
Private RigaArray As Long
Private CurrRow As Long
Private FormPronta As Boolean
Private FileDaAprire As String
Private TipoPDI() As tyTipoPDI
Private txtHTML() As String
Const FormCaption As String = "Download Manager"

' Classe per il menu PopUp
Dim mnu As clsMenu

Private Sub chkUseProxy_LostFocus()
    Call ImpostazioniProxy(True)
End Sub

Private Sub chkUseProxy_Click()
    Dim blnEnabled As Boolean

    blnEnabled = chkUseProxy.value
    If blnEnabled = False Then
        picProxy.Appearance = 1
        picProxy.BackColor = &H8000000F
    Else
        picProxy.Appearance = 0
        picProxy.BackColor = &H8000000F
    End If
    Label2.Enabled = blnEnabled
    Label3.Enabled = blnEnabled
    Label6.Enabled = blnEnabled
    Label7.Enabled = blnEnabled
    txtProxyServer.Enabled = blnEnabled
    txtProxyPort.Enabled = blnEnabled
    txtProxyUserName.Enabled = blnEnabled
    txtProxyPassword.Enabled = blnEnabled
    
End Sub

Private Sub ckBrowser_Click()

    If ckBrowser.value = 1 Then
        wbrSource.ZOrder 0
        ckBrowser.ZOrder 0
        With wbrSource
            .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        End With

    Else
        wbrSource.ZOrder 1
        With wbrSource
            .Move 0, 0, 0, 0
        End With

    End If
    
End Sub

Private Sub cmbTipoPDI_ClickPopUp(ValoreCliccato As Long)
    
    Select Case ValoreCliccato
        Case 10
            cmbTipoPDI.SaveFile (Var(TipoPdiCsv).Valore)
    End Select
    
End Sub

Private Sub cmdApriCartellaFile_Click()
    Dim cart As String
    
    cart = Var(PoiScaricati).Valore
    'Shell "explorer.exe " & PoiScaricati, vbNormalFocus
    ShellExecute 0, "Open", cart, vbNullString, vbNullString, SW_SHOWNORMAL

End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Public Sub cmdScaricaFile_Click(index As Integer)
    Dim ret
    Dim Msg As String

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    If ListView1.ListItems.count > 0 Then
        Screen.MousePointer = vbHourglass
        DoEvents
        
        ' Cancello i file nella directory tmpDownFile
        ret = DeleteAllFiles(Var(tmpDownFile).Valore)
        
        ' Scrivo i dati della ListView1 nell'array
        Call ArrayDaListView
        
        lstLog.AddItem "Inizio operazioni di download " & Now
        lstLog.AddItem " "
        
        RigaArray = 0
        
        Select Case index
            Case 0 ' Solo selezionati
                SoloChecked = True
                
            Case 1 ' Tutti i file
                SoloChecked = False
                
        End Select
        
        If elencoDa = xml Then
            ' Scarico i file .ov2 ed i file .bmp
            ScaricaDaLink True, Var(PoiScaricati).Valore, Var(ScaricaBMP).Valore
            
        ElseIf elencoDa = Web Then
            ' Scarico il file .zip
            ScaricaDaLink True
        
            ' Scrivo nella StatusBar1
            StatusBar1.Panels(2).Text = "Inizio operazioni di decompressione...."
            DoEvents
            Call AvviaOperazioniUnZip
        
        End If
        
        StatusBar1.Panels(2).Text = "Operazioni di download terminate"
        lstLog.AddItem "Operazioni di download terminate " & Now
        
        If elencoDa = xml Then
            lstLog.AddItem "Inizio elaborazione dei nomi dei file scaricati"
            Call ElaboraNomeFileInCartella(Var(PoiScaricati).Valore, , True)
        End If

        ' Imposto tutti i file .bmp come file nascosti per evitare che vengano visualizzati nella galleria dei cellulari nokia
        Call NascondiFile("*.bmp", Var(PoiScaricati).Valore)

        lstLog.AddItem "Termine di tutte le operazioni"
        lstLog.AddItem " ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- "
        lstLog.AddItem " "
        StatusBar1.Panels(2).Text = ""
        
        Call ApriFormPOICartella
            
    Else
        Msg = "ATTENZIONE!" & vbNewLine & " Prima devi scaricare l'elenco dei file." & vbNewLine & " Premi il pulsante ''Scarica elenco file disponibili''" & vbNewLine & " Vuoi scaricare ora l'elenco dei file?         "
        ret = MsgBox(Msg, vbYesNo + vbExclamation + vbDefaultButton1)
        If ret = vbYes Then
           cmdScaricaElenco_Click (0)
        Else
           cmdScaricaElenco(0).SetFocus
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione frmDownload.cmdScaricaFile_Click." & vbNewLine & "Index: " & index
    
End Sub

Private Sub ScaricaDaLink(Optional ByVal SelezionaRiga As Boolean = False, Optional ByVal CartellaDest As String = "", Optional ByVal bScaricaBMP As Boolean = False)
    Dim cnt As Long
    Dim cnt1 As Long
    Dim ret
    Dim strTmp As String

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    ProgressBar1.Max = ListView1.ListItems.count
    
    If CartellaDest = "" Then CartellaDest = Var(tmpDownFile).Valore
    
    For cnt = 1 To ListView1.ListItems.count
        If ListViewChecked(ListView1, cnt) = True Or SoloChecked = False Then
            
            lstLog.AddItem "Scarico: " & arrElencoFile(cnt - 1).Descrizione & "." & arrElencoFile(cnt - 1).Formato
            DoEvents
            
            ret = DownloadFile(arrElencoFile(cnt - 1).upURL, CartellaDest & "\" & Replace$(arrElencoFile(cnt - 1).Nome, """", "'") & "." & arrElencoFile(cnt - 1).Formato)
            
            strTmp = CartellaDest & "\" & Replace$(arrElencoFile(cnt - 1).Nome, """", "'") & EstensioneFromFile(arrElencoFile(cnt - 1).Immagine)
            If (bScaricaBMP = True) Or ((bScaricaBMP = False) And (FileExists(strTmp) = False)) Then
                ret = DownloadFile(arrElencoFile(cnt - 1).Immagine, strTmp)
            End If
            
            If SelezionaRiga = True Then
                Call SelezionaRigaListView(ListView1, cnt + 1, True)
                lstLog.ListIndex = lstLog.ListCount - 1
            End If
            
            cnt1 = cnt1 + 1
            
        End If
        ProgressBar1.value = cnt
    Next
    
    lstLog.AddItem " "
    lstLog.AddItem "Scaricati " & cnt1 & " file."
    
    ProgressBar1.value = 0.01

    Exit Sub
    
Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione frmDownload.ScaricaDaLink." & vbNewLine & "Indice: " & cnt

End Sub

Private Sub ApriFormPOICartella()
    Dim ret

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    If cekRemake.value = 0 Then
        ret = MsgBox("I file sono stati scaricati." & vbNewLine & "Adesso i file dovrebbero essere trattati per l'inserimento degli Skipper Record." & vbNewLine & "Vuoi effettuare il trattamento dei file ora?" & vbNewLine & vbNewLine & "- Se scegli ""Si"" verrà aperta la finestra che permette il trattamento dei file." & vbNewLine & "- Se scegli ""No"" verrà aperta la cartella PoiScaricati." & vbNewLine & "- Se scegli ""Annulla"" ritorni alla finestra " & Me.Caption & ".", vbInformation + vbYesNoCancel, App.ProductName)
    End If
    
    If ret = vbYes Or cekRemake.value = 1 Then
        Load frmRemakeov2
        frmRemakeov2.Show
        Me.Visible = False
        DoEvents
        frmRemakeov2.cmdInserisciFile_Click
        
    ElseIf ret = vbNo Then
        ' Apro la cartella
        ShellExecute 0, "Open", Var(PoiScaricati).Valore, vbNullString, vbNullString, SW_SHOWNORMAL
    End If
    
    Exit Sub

Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione frmDownload.ApriFormPOICartella."

End Sub

Private Sub cmdScaricaTipiPDI_Click()
    Dim cnt As Long
    Dim avanz As Long
    Dim i As Long
    Dim id As String
    Dim PaginaWeb As String
    Dim strCerca As String
    Dim Comodo As String
    Dim strHTML As String
    
    Screen.MousePointer = vbHourglass
    DoEvents
    cmdScaricaTipiPDI.Enabled = False
    ProgressBar1.Max = 100
    avanz = 10
    
    PaginaWeb = "http://www.poigps.com/modules.php?name=Downloads"
    strCerca = LCase("http://www.poigps.com/modules.php?name=Downloads&d_op=viewdownload&cid=")
    
    cmbTipoPDI.Clear True

    ' Open the URL
    wbrSource.Navigate PaginaWeb
    
    ' Wait until the document is loaded
    ' Attenzione se la pagina contiene frames, la ruotine sballa pensando che la pagina sia caricata, quando invece mancano ancora altri frame
    Do While wbrSource.ReadyState <> READYSTATE_COMPLETE
        MsgWaitObj 1000
        DoEvents
        If ProgressBar1.value + avanz >= ProgressBar1.Max Then
            ProgressBar1.value = 1
        Else
            ProgressBar1.value = ProgressBar1.value + avanz
        End If
    Loop
    
    ' Prendo il codice html della pagina che ho navigato per ricercare dei valori....
    ' strHTML$ = wbrSource.Document.documentElement.innerHtml
    cnt = 0
    ' Cancello l'array
    ReDim arrElencoFile(0)
    ProgressBar1.value = 0.01
    ProgressBar1.Max = wbrSource.Document.links.Length - 1
    ' List the links available
    For i = 0 To wbrSource.Document.links.Length - 1
        If strCerca = LCase(Left(wbrSource.Document.links(i).href, Len(strCerca))) Or strCerca = "" Then
            ' Filtro per il valore nella TextBox
            If txtFiltra.Text = Left$(Trim$(wbrSource.Document.links(i).innerText), Len(txtFiltra.Text)) Or txtFiltra.Text = "" Then
                Comodo = Trim$(wbrSource.Document.links(i).innerText)
                Comodo = Comodo & vbTab & wbrSource.Document.links(i).href
                'id = Mid$(wbrSource.Document.links(i).href, Len(strCerca) + 1)
                cmbTipoPDI.AddItem Comodo
                cnt = cnt + 1
            End If
        End If
        ProgressBar1.value = i
    Next
    
    ' Seleziono il primo elemento
    If cmbTipoPDI.ListCount > 0 Then
        cmbTipoPDI.ListIndex = 0
        'AutoSizeLBHeight cmbTipoPDI
    Else
        ckBrowser.value = 1
        MsgBox "Controlla che il sito ti abbia riconosciuto. Probabilmente devi effettuare il LogIn.", vbInformation
    End If
    
    ProgressBar1.value = 0.01
    Screen.MousePointer = vbDefault
    cmdScaricaTipiPDI.Enabled = True

End Sub

Public Sub cmdScaricaElenco_Click(index As Integer)
    Dim cnt As Long
    Dim avanz As Long
    Dim strTmp() As String
    Dim strVar As String
    Dim Campo As String
    Dim pos1 As Long
    Dim pos2 As Long
    Dim PosTmp As Long
    Dim Comodo As String
    
    Screen.MousePointer = vbHourglass
    DoEvents
    cmdScaricaElenco(index).Enabled = False
    ProgressBar1.Max = 100
    avanz = 5

    ' Estraggo il nome della mappa e la pagina web selezionati
    ReDim TipoPDI(0)
    TipoPDI(0).NomeMappa = Trim$(Mid$(cmbTipoPDI.ItemText(cmbTipoPDI.ListIndex, 0), 5))
    TipoPDI(0).PaginaWeb = cmbTipoPDI.ItemText(cmbTipoPDI.ListIndex, 1)

    Select Case index
    Case 0 ' Modalità XML  ------------------------------------------------------------------------------------
        Dim xmlDoc As Object
        Dim xmlList As Object
        Dim xmlNode As Object
        
        Me.Caption = FormCaption & " - XML"
        
        Set xmlDoc = CreateObject("MSXML.DOMDocument")
        xmlDoc.async = False
        txtListBox.Text = " Download del file " & Var(PoiGpsXmlWeb).Valore & " in corso....."
        xmlDoc.Load (Var(PoiGpsXmlWeb).Valore)
        ' visualizzo la pagina Web nel browser
        wbrSource.Navigate Var(PoiGpsXmlWeb).Valore
        txtListBox.Text = ""
        
        If Not xmlDoc Is Nothing Then
            cnt = 0
            
            elencoDa = xml

            ' Cancello l'array
            ReDim arrElencoFile(0)
            Set xmlList = xmlDoc.getElementsByTagName("poi")
            
            ' Conto quanti nodi ci sono (serve per la progressbar)
            For Each xmlNode In xmlList
                cnt = cnt + 1
            Next
            ProgressBar1.Max = IIf(cnt > 0, cnt, 1)
            
            cnt = 0
            For Each xmlNode In xmlList
                ReDim Preserve arrElencoFile(cnt)
                
                Comodo = xmlNode.selectSingleNode("group").Text
                arrElencoFile(cnt).Gruppo = Comodo
                
                Comodo = xmlNode.selectSingleNode("note").Text
                arrElencoFile(cnt).Note = Comodo
                If IsDate(Right$(Comodo, 10)) = True Then arrElencoFile(cnt).data = Right$(Comodo, 10)

                Comodo = xmlNode.selectSingleNode("description").Text
                arrElencoFile(cnt).Descrizione = Comodo
                
                Comodo = xmlNode.selectSingleNode("url").Text
                Comodo = "http://" & Mid$(Comodo, InStr(1, LCase(Comodo), "www", vbTextCompare))
                If InStr(Comodo, "?") <> 0 Then Comodo = Mid$(Comodo, 1, InStr(1, LCase(Comodo), "?", vbTextCompare) - 1)
                arrElencoFile(cnt).url = Replace(Comodo, " ", "%20", , , vbTextCompare)
                
                Comodo = xmlNode.selectSingleNode("url").Text
                Comodo = Replace(Comodo, "%Username%", Var(PoiGpsXmlUserName).Valore)
                Comodo = Replace(Comodo, "%Password%", Var(PoiGpsXmlPsw).Valore)
                arrElencoFile(cnt).upURL = Comodo
                
                Comodo = xmlNode.selectSingleNode("format").Text
                arrElencoFile(cnt).Formato = Comodo
                
                Comodo = xmlNode.selectSingleNode("image").Text
                arrElencoFile(cnt).Immagine = Replace(Comodo, " ", "%20", , , vbTextCompare)
                
                Comodo = xmlNode.selectSingleNode("map").Text
                arrElencoFile(cnt).Mappa = Comodo
                
                arrElencoFile(cnt).Modalità = "xml"
    
                ProgressBar1.value = cnt
                cnt = cnt + 1
                DoEvents
            Next
        End If
        
        Set xmlDoc = Nothing
        
        Call ArrayInListView
        
    Case 1 ' Modalità WEB ------------------------------------------------------------------------------------
        Dim i As Integer
        Dim id As String
        Dim lenTxt0 As Long
        Dim PaginaWeb As String
        Dim strCerca As String
        
        Me.Caption = FormCaption & " - WEB"
        elencoDa = Web
        cekFiltraPerMappa.value = 0

        PaginaWeb = TipoPDI(0).PaginaWeb & "&show=2000"
        strCerca = Var(PoiGpsWebWeb).Valore
        
        ListView1.ListItems.Clear
        
        txtListBox.Text = " Download dati della pagina web....."
        ' Open the URL
        wbrSource.Navigate url:=PaginaWeb, Headers:="Authorization: Basic XXXXXX" & Chr$(13) & Chr$(10)
        'wbrSource.Navigate paginaWeb
        
        ' Wait until the document is loaded
        Do While wbrSource.ReadyState <> READYSTATE_COMPLETE
            MsgWaitObj 500
            DoEvents
            MsgWaitObj 500
            DoEvents
            
            If ProgressBar1.value + avanz >= ProgressBar1.Max Then
                ProgressBar1.value = 1
            Else
                ProgressBar1.value = ProgressBar1.value + avanz
            End If
        Loop
        
        txtListBox.Text = ""
        
        ' Prendo tutto il testo contenuto nella pagina
        ' per poter cercare la data di inserimento del file nel sito
        strTmp = Split(wbrSource.Document.body.All.Item.innerHTML, vbNewLine, , vbTextCompare)
        Campo = "<TD bgColor=#ecf0f7><IMG height=14 alt=Downloads src="
        
        cnt = 0
        ReDim txtHTML(wbrSource.Document.links.Length - 1, 2)
        
        For i = 0 To UBound(strTmp)
            If Left$(strTmp(i), Len(Campo)) = Campo Then
            
                pos1 = InStr(1, strTmp(i), "ttitle", vbTextCompare)
                pos2 = InStr(pos1, strTmp(i), """>", vbTextCompare)
                PosTmp = InStr(pos1, strTmp(i), """&gt", vbTextCompare)
                If PosTmp <> 0 Then
                    pos2 = PosTmp
                End If

                If pos1 <> 0 And pos2 <> 0 Then
                    ' Trovo il nome del file
                    Comodo = LCase$(Mid$(strTmp(i), pos1 + 7, pos2 - (pos1 + 7)))
                    'Debug.Print Comodo
                    Comodo = Replace$(Comodo, "&amp;g", "")
                    Comodo = Replace$(Comodo, "&amp;", "")
                    Comodo = Replace$(Comodo, "<", "")
                    Comodo = Replace$(Comodo, ">", "")
                    Comodo = Replace$(Comodo, "/", "")
                    Comodo = Replace$(Comodo, "_", "")
                    Comodo = Replace$(Comodo, """", "")
                    Comodo = Replace$(Comodo, "&", "")
                    Comodo = Replace$(Comodo, ";", "")
                    Comodo = Replace$(Comodo, "?", "")
                    Comodo = Replace$(Comodo, "b", "")
                    Comodo = Replace$(Comodo, " ", "")
                    'Debug.Print Comodo
                    txtHTML(cnt, 1) = Comodo
                End If
                
                pos1 = InStr(1, strTmp(i), "Aggiunto il:</B>", vbTextCompare)
                pos2 = pos1 + 17 + 11
                If pos1 <> 0 And pos2 <> 0 Then
                    ' Trovo la data di inserimento del file nel sito
                    Comodo = Trim$(Mid$(strTmp(i), pos1 + 17, pos2 - (pos1 + 17)))
                    txtHTML(cnt, 2) = Replace$(Comodo, "&", "")
                End If

                cnt = cnt + 1
            End If
            DoEvents
        Next

        cnt = 0
        ' Cancello l'array
        ReDim arrElencoFile(0)
        ' List the links available
        
        On Error Resume Next
        
        ProgressBar1.Max = wbrSource.Document.links.Length - 1
        For i = 0 To wbrSource.Document.links.Length - 1
            lenTxt0 = Len(strCerca)
            If strCerca = Left(wbrSource.Document.links(i).href, lenTxt0) Or strCerca = "" Then
                ReDim Preserve arrElencoFile(cnt)
                ' Prendo il nome del file
                arrElencoFile(cnt).Nome = Trim$(wbrSource.Document.links(i).innerText)
                
                Comodo = CercaDataAggiuntaFile(arrElencoFile(cnt).Nome)
                If IsDate(Comodo) = True Then arrElencoFile(cnt).data = Comodo
                
                arrElencoFile(cnt).Descrizione = Trim$(wbrSource.Document.links(i).innerText)
                arrElencoFile(cnt).Formato = "zip"
                arrElencoFile(cnt).Mappa = TipoPDI(0).NomeMappa
                
                ' Prendo il link del file
                ' arrElencoFile(cnt).URL = wbrSource.Document.Links.item(i) ' Da il solito link della riga più sotto
                arrElencoFile(cnt).url = wbrSource.Document.links(i).href
                arrElencoFile(cnt).upURL = wbrSource.Document.links(i).href
                'id = Mid$(wbrSource.Document.links(i).href, lenTxt0 + 1)
                
                arrElencoFile(cnt).Modalità = "web"
                
                cnt = cnt + 1
            End If
            ProgressBar1.value = i
            DoEvents
        Next
        
        ProgressBar1.value = 0.01
        
        Call ArrayInListView
        
    End Select
    
    ProgressBar1.value = 0.01
    
    If ListView1.ListItems.count > 0 Then
        Call ControllaCheck(ListView1)
        Call OrdinaColonnaByTag(ListView1, ListView1.ColumnHeaders.Item(3))
        Call SelezionaRigaListView(ListView1, 1)
    Else
        frmDownload.Visible = True
        strVar = "Non sono stati trovati PDI da scaricare." & vbNewLine
        
        Select Case elencoDa
            Case Is = Web
                ckBrowser.value = 1
                strVar = strVar & "Controlla che il sito ti abbia riconosciuto. Probabilmente devi effettuare il LogIn."
            Case Is = xml
                strVar = strVar & "Link per download: " & Var(PoiGpsXmlWeb).Valore & vbNewLine & "E' possibile che non sia disponibile nessun PDI per la categoria selezionata, riprova quindi con una altra categoria."
        End Select
        
        MsgBox strVar, vbExclamation, App.ProductName
    End If

    TotaleRigheListView ListView1
    
    Screen.MousePointer = vbDefault
    cmdScaricaElenco(index).Enabled = True

End Sub

Private Function CercaDataAggiuntaFile(ByVal NomeFile As String) As String
    Dim i As Long
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Esci
    
    NomeFile = Replace$(LCase$(NomeFile), "_", "")
    NomeFile = Replace$(NomeFile, "<", "")
    NomeFile = Replace$(NomeFile, ">", "")
    NomeFile = Replace$(NomeFile, "/", "")
    NomeFile = Replace$(NomeFile, """", "")
    NomeFile = Replace$(NomeFile, "&", "")
    NomeFile = Replace$(NomeFile, ";", "")
    NomeFile = Replace$(NomeFile, "?", "")
    NomeFile = Replace$(NomeFile, "b", "")
    NomeFile = Replace$(NomeFile, " ", "")
    
    For i = 0 To UBound(txtHTML)
        If txtHTML(i, 1) = NomeFile Then
            ' Cancello il campo
            txtHTML(i, 1) = ""
            ' Assegno la data al risultato della funzione
            CercaDataAggiuntaFile = txtHTML(i, 2)
            Exit For
        End If
    Next
    
    Exit Function

Esci:
    CercaDataAggiuntaFile = ""
    
End Function

Private Sub ArrayInListView()
    Dim cntRecord As Long
    Dim cnt As Long
    Dim itmX As Variant

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    ListView1.ListItems.Clear
        
    cntRecord = 0
    For cnt = 0 To UBound(arrElencoFile) ' Scorro tutte le righe dell'array
        If (arrElencoFile(cnt).url <> "" And arrElencoFile(cnt).Mappa = TipoPDI(0).NomeMappa) _
          Or (arrElencoFile(cnt).url <> "" And cekFiltraPerMappa.value = 0) Then
            Set itmX = ListView1.ListItems.Add(, , Format(cntRecord + 1, "00000"))
            itmX.SubItems(GetNumColDaIntestazione(ListView1, "Descrizione") - 1) = arrElencoFile(cnt).Descrizione
            itmX.SubItems(GetNumColDaIntestazione(ListView1, "Data") - 1) = arrElencoFile(cnt).data
            itmX.SubItems(GetNumColDaIntestazione(ListView1, "Url.") - 1) = arrElencoFile(cnt).url
            itmX.SubItems(GetNumColDaIntestazione(ListView1, "upUrl") - 1) = arrElencoFile(cnt).upURL
            itmX.SubItems(GetNumColDaIntestazione(ListView1, "Formato") - 1) = arrElencoFile(cnt).Formato
            itmX.SubItems(GetNumColDaIntestazione(ListView1, "Immagine") - 1) = arrElencoFile(cnt).Immagine
            itmX.SubItems(GetNumColDaIntestazione(ListView1, "Mappa") - 1) = arrElencoFile(cnt).Mappa
            itmX.SubItems(GetNumColDaIntestazione(ListView1, "Note") - 1) = arrElencoFile(cnt).Note
            itmX.SubItems(GetNumColDaIntestazione(ListView1, "Modalità") - 1) = arrElencoFile(cnt).Modalità
            cntRecord = cntRecord + 1
        End If
    Next
    
    ListView1.Refresh
    Call SelezionaRigaListView(ListView1, ListView1.ListItems.count, True)

    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione frmDownload.ArrayInListView." & vbNewLine & "Record: " & cnt

End Sub

Private Sub ArrayDaListView()
    Dim cnt As Long

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    ' Se non ci sono righe esco
    If ListView1.ListItems.count = 0 Then Exit Sub
    
    ' Dimensiono l'array
    ReDim arrElencoFile(ListView1.ListItems.count - 1)
    
    ' Scorro tutte le righe.....
    For cnt = 0 To ListView1.ListItems.count - 1
        arrElencoFile(cnt).Descrizione = Trim(ListView1.ListItems(cnt + 1).SubItems(1))
        arrElencoFile(cnt).Nome = Trim(ListView1.ListItems(cnt + 1).SubItems(1))
        arrElencoFile(cnt).data = Trim(ListView1.ListItems(cnt + 1).SubItems(2))
        arrElencoFile(cnt).url = Trim(ListView1.ListItems(cnt + 1).SubItems(3))
        arrElencoFile(cnt).upURL = Trim(ListView1.ListItems(cnt + 1).SubItems(4))
        arrElencoFile(cnt).Formato = Trim(ListView1.ListItems(cnt + 1).SubItems(5))
        arrElencoFile(cnt).Immagine = Trim(ListView1.ListItems(cnt + 1).SubItems(6))
        arrElencoFile(cnt).Mappa = Trim(ListView1.ListItems(cnt + 1).SubItems(7))
        arrElencoFile(cnt).Note = Trim(ListView1.ListItems(cnt + 1).SubItems(8))
    Next
    
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione frmDownload.ArrayDaListView." & vbNewLine & "Riga: " & cnt
        
End Sub

Private Sub cmdUpdateFile_Click()
    frmUpdateFile.Show
End Sub

Private Sub Form_Activate()
    If FormIsLoad("frmMain") = True Then frmMain.Visible = False
    FormPronta = True
    ListView1.SetFocus
End Sub

Private Sub Form_Initialize()
    FormPronta = False
    Me.Caption = FormCaption
End Sub

Private Sub Form_Load()
    Dim LarghezzaColonne As Variant
    Dim ColonneListView1 As Variant
    Dim TagColonne As Variant
    Dim cnt As Long
    Dim ctl As Object

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    ' Inizializzo la classe per la ListBox
    Set lbLog = New clsListBox
    lbLog.Attach lstLog
    
    ' Creo gli array con i dati delle colonne
    LarghezzaColonne = Array(1200, 3000, 1300, 3000, 20, 800, 20, 1000, 900, 650)
    ColonneListView1 = Array("", "Descrizione", "Data", "Url.", "upUrl", "Formato", "Immagine", "Mappa", "Note", "Modalità")
        ' Serve per la funzione di ordinamento dei dati nella colonna
          TagColonne = Array("NUMBER", "STRING", "DATE", "STRING", "STRING", "STRING", "STRING", "STRING", "STRING", "STRING")
    
    With ListView1
        .HideSelection = False
        .FullRowSelect = True
        .MultiSelect = Var(SelezMultipla).Valore
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

    ' Pulisco i TextBox
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next
        
    ' Dimensiono l'array ad un valore (cosi non si hanno errori nelle funzioni Ubound ecc..)
    ReDim arrElencoFile(0)
    ' Preparo le variabili
    UsaPassword = True
    FileScaricato = False
    SoloChecked = False
    RigaArray = -1
    elencoDa = Noo

    With StatusBar1
        .Panels(1).Width = 100
        .Panels.Add
        .Panels(2).Width = 4000
        .Panels.Add
        .Panels(3).Width = 3400
        .Panels(3).Text = IEversion
        .Panels.Add
        .Panels(4).AutoSize = sbrSpring
        .Height = 280
        .ZOrder 0
    End With
    
    With ProgressBar1
        .value = 0.01
        .BorderStyle = ccNone
        .ScrollIng = ccScrollingSmooth
        .ZOrder 0
    End With
    
    txtUserName.Text = Var(PoiGpsXmlUserName).Valore
    txtPassword.Text = Var(PoiGpsXmlPsw).Valore
    txtFiltra.Text = "PDI"
        
    ' Carico le impostazioni del Proxy
    Call ImpostazioniProxy
    
    With cmbTipoPDI
        'FormatString creates a Column for each item in the string.
        '> Right Justify  '< Left Justify  '^ Centre Justify
        .FormatString = "< Tipo PDI|< PaginaWeb"
        
        .ColWidth(0) = .Width - 240
        .ColWidth(1) = 0
        
        'ColType property allows Sort to process values correctly
        .ColType(0) = TypeString
        .ColType(1) = TypeString
        
        '.ImageList = imglRunSearch 'Set images
        
        'Override column alignment from FormatString
        .ColAlignment(0) = AlignLeftCenter
        .ColAlignment(1) = AlignCenterCenter
        
        '.AddItem "XD " & 10
        '.ItemImage(.NewIndex) = 61
        '.ItemText(.NewIndex, 1) = "B"
        '.ItemForeColor(.NewIndex) = vbRed
        
        .LoadFile Var(TipoPdiCsv).Valore
        .ListIndex = 0
        
        .AddItemMenu 10, "Salva Elenco"
    End With
    
    If CommandLineFile <> "" Then
        Call ClickPopUp(10, CommandLineFile)
        FileDaAprire = CommandLineFile
        CommandLineFile = ""
    End If
    
    Form_Resize
    wbrSource.Navigate ("about.blank")
    
    Exit Sub
    
Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione frmDownload.Form_Load."

End Sub

Private Sub ImpostazioniProxy(Optional Salva As Boolean = False)
    ' Formato della stringa ProxySet: UsaProxy,Server,Porta,User,Password
    Dim strTmp As String
    Dim vTmp
    
    If Salva = True Then ' Salvo le impostazioni
        strTmp = chkUseProxy.value
        strTmp = strTmp & "," & txtProxyServer.Text
        strTmp = strTmp & "," & txtProxyPort
        strTmp = strTmp & "," & txtProxyUserName
        strTmp = strTmp & "," & txtProxyPassword
        lVar(ProxySet) = strTmp

    ElseIf Salva = False Then ' Leggo le impostazioni
        If Var(ProxySet).Valore = "0" Then
            chkUseProxy.value = 0
        Else
            vTmp = Split(Var(ProxySet).Valore, ",")
            chkUseProxy.value = vTmp(0)
            txtProxyServer.Text = vTmp(1)
            txtProxyPort = vTmp(2)
            txtProxyUserName = vTmp(3)
            txtProxyPassword = vTmp(4)
        End If
    End If

End Sub

Private Sub Form_Paint()
    
    FormPronta = False
    With ProgressBar1
        .Height = 210
        .Top = StatusBar1.Top + 40
        .Left = StatusBar1.Panels(4).Left + 30
        .Width = StatusBar1.Panels(4).Width - 60
    End With
    FormPronta = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error Resume Next
    cmdEsci.SetFocus ' l'ho inserito per attivare il codice collegato agli eventi lostfocus....
    Err.Clear
    DoEvents
    
    FormPronta = False
    NomeFileAperto = ""
    
    If SecondaIstanza = True Then
        Unload frmMain
    Else
        frmMain.Visible = True
        frmMain.WindowState = vbNormal
        frmMain.ZOrder
        frmMain.SetFocus
    End If
    
    DoEvents

End Sub

Private Sub Form_Resize()
   ' Width = Larghezza  Height = Altezza
    Dim pos As Long
    Dim ckBrowserVal As Integer

    ckBrowserVal = ckBrowser.value
    
    FormPronta = False

    On Error Resume Next
    
    ckBrowser.value = 0
    
    picSlider.Height = 55
    picShadow.Height = 35
    picShadow.Top = picSlider.Top

    ' Imposto la larghezza dei controlli
    ListView1.Width = Me.ScaleWidth
    picShadow.Width = Me.ScaleWidth
    picSlider.Width = Me.ScaleWidth

    pos = picCommand.Top + picCommand.Height + (StatusBar1.Height * 3)
    If Me.Height < pos Then Me.Height = pos
    
    With lstLog
        .Move 0, .Top, Me.ScaleWidth, Me.ScaleHeight - StatusBar1.Height - .Top
    End With
    
    With wbrSource
        .Move 0, 0, 0, 0
        .ZOrder 1
    End With
    
    With picComm
        ' Centro il controllo nella form
        .Move (Me.Width - .Width) / 2
    End With
    
    With picCommand
        ' Centro il controllo nella form
        .Move (Me.Width - .Width) / 2
    End With
    
    With ListView1
        .Move 0, .Top, Me.ScaleWidth
        txtListBox.Move (.Left + (.Width / 2)) - txtListBox.Width / 2, (.Top + (.Height / 2)) - txtListBox.Height / 2
    End With
    
    Dragging = True
    picShadow.Top = Me.ScaleHeight / 10 * 8.5
    SliderMove

    ckBrowser.value = ckBrowserVal
    
    FormPronta = False

End Sub

Private Sub SliderMove()
    On Error Resume Next
    
    ' Turn off dragging and hide the shadow
    Dragging = False
    picShadow.Visible = False
    
    ' Make sure the shadow was not moved too far
    If picShadow.Top + StatusBar1.Height - picShadow.Height > Me.ScaleHeight Then picShadow.Top = Me.ScaleHeight - StatusBar1.Height - picShadow.Height
    If picShadow.Top <= ListView1.Top + picCommand.Height + 70 Then picShadow.Top = ListView1.Top + picCommand.Height + 70
    
    picCommand.ZOrder 0
    ' Move picSlider and resize the controls
    picSlider.Top = picShadow.Top
    picCommand.Top = picSlider.Top - picSlider.Height - picCommand.Height
    ListView1.Height = picCommand.Top - ListView1.Top
    lstLog.Top = picSlider.Top + picSlider.Height
    lstLog.Height = ScaleHeight - picSlider.Top + picSlider.Height - StatusBar1.Height

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call OrdinaColonnaByTag(ListView1, ColumnHeader)
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ControllaRigaChecked(ListView1, Item)
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 93 Then ' Tasto per il menu del tato destro del mouse
        MouseMove Me, 100, 100
        Call ListView1_MouseUp(vbRightButton, 0, 0, 0)
    ElseIf KeyCode = 32 Then ' Tasto Spazio
        KeyCode = 0
    ElseIf KeyCode = 46 Then  ' Tasto Canc
        txtListBox.Text = ""
    End If
    
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 8 Then ' Tasto Back
        If txtListBox.Text <> "" Then txtListBox.Text = Mid(txtListBox.Text, 1, Len(txtListBox.Text) - 1)
    ElseIf KeyAscii = 13 Then ' Tasto Invio
        txtListBox.Text = ""
    Else
        txtListBox.Text = txtListBox.Text & Chr(KeyAscii)
    End If
    
    txtListBox.SelStart = Len(txtListBox.Text)
End Sub

Private Sub ListView1_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    If data.GetFormat(vbCFFiles) Then
        For i = 1 To 1 'Data.Files.count
            FileDaAprire = (data.Files(i))
            Call ClickPopUp(10, FileDaAprire)
        Next
    End If

End Sub

Private Sub lstLog_DblClick()
    Dim strTmp As String
    Dim ret
    
    If lstLog.ListCount = 0 Then Exit Sub
    
    strTmp = lstLog.List(lstLog.ListIndex)
    'ret = MsgBox(" Vuoi copiare questa riga negli appunti? ", vbInformation + vbYesNo, App.ProductName)
    'If ret = vbYes Then
        Clipboard.Clear
        Clipboard.SetText strTmp
    'End If
    
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

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        CurrRow = GetRigaSelezionata(ListView1, X, Y)
    ElseIf Button = vbLeftButton Then
        '
    End If
    
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    Dim ListaVuota As Boolean
    Dim Colonna As Long

    If ListView1.ListItems.count = 0 Then
        ListaVuota = True
    Else
        ListaVuota = False
    End If
    
    ' Per il PopUp menu
    If Button = vbRightButton Then
        Set mnu = New clsMenu
        With mnu
            'Dim submnu1 As clsMenu: Set submnu1 = New clsMenu
            'With submnu1
            '    .Caption = "Mostra sulla cartina"
            '    .AddItem 101, "www.multimap.com (" & CurrRow & ")", , , , ListaVuota
            '    .AddItem 102, "www.mapquest.com (" & CurrRow & ")", , , , ListaVuota
            'End With
            .AddItem 10, "Apri..."
            .AddItem 20, "Salva... " & NomeFileAperto, , , , Not ListaVuota Imp Not FileExists(PatchNomeFileAperto)
            .AddItem 21, "Salva con nome...", , , , ListaVuota
            .AddItem 0, "-"
            '
            Colonna = GetNumColDaIntestazione(ListView1, "Descrizione") - 1
            .AddItem 30, "Scarica il file: " & (GetValoreCella(ListView1, CurrRow, Colonna)), , , , ListaVuota
            .AddItem 31, "Apri pagina web con tutti i poi"
            .AddItem 0, "-"
            '
            .AddItem 49, "Numera le righe della lista", , , , ListaVuota
            .AddItem 0, "-"
            '
            .AddItem 50, "AutoSize larghezza colonne", , , , ListaVuota
            .AddItem 59, "Selezione righe multiple", , Var(SelezMultipla).Valore, , ListaVuota
            .AddItem 60, "Cancella riga selezionata", , , , ListaVuota
            .AddItem 70, "Cancella righe selezionate", , , , ListaVuota
            .AddItem 80, "Seleziona tutto", , , , ListaVuota
            .AddItem 81, "Deseleziona tutto", , , , ListaVuota
            '
            '.AddItem 95, "Nuova Colonna", , , , , True
            ClickPopUp .TrackPopup(Me.hwnd)
        End With
    End If

    If FormIsLoad("frmSetupDescrizione") = True Then
        frmSetupDescrizione.lblCambiaRigaCorrente.Caption = GetRigaSelezionata(ListView1, X, Y)
    End If
    
    Exit Sub
    
Errore:
    GestErr Err, "Errore nella funzione ListView1_MouseUp."

End Sub

Private Sub ClickPopUp(ByVal ValoreCliccato As Long, Optional ByVal FileDaAprire As String = "")
    Dim Result
    Dim PatchNomeFile As String
    
    PatchNomeFile = ""

    Select Case Left(ValoreCliccato, 2)
        Case Is = 10 ' Apri
            If EstensioneFromFile(FileDaAprire) <> ".rmk" Then FileDaAprire = ""
            If ImportaDati(Me.hwnd, ListView1, FileDaAprire, "rmk", , True, Var(UltimoFileWEB).Valore, , True, "FileImpostazioniDownloadWeb") = True Then
                Me.Caption = FileDaAprire
                ' Scrivo i dati nel file .xml
                lVar(UltimoFileWEB) = FileDaAprire
                elencoDa = arrCampiRmkFile(1)
                Call CaricaArrElencoFile
                Call ControllaCheck(ListView1)
            Else
                Me.Caption = FormCaption
                Result = MsgBox("Non è stato possibile importare il file!", vbOKOnly)
            End If
            Screen.MousePointer = vbDefault

        Case Is = 20 ' Salva
            Call ExportaDati(Me.hwnd, PatchNomeFileAperto, , "rmk", True, , Var(CampiRmkFile).Valore & "FileImpostazioniDownloadWeb;" & elencoDa & ";" & Now, ListView1, True)
            
        Case Is = 21 ' Salva con nome
            Call ExportaDati(Me.hwnd, , NomeFileAperto, "rmk", True, , Var(CampiRmkFile).Valore & "FileImpostazioniDownloadWeb;" & elencoDa & ";" & Now, ListView1, True)
            
        Case Is = 30 ' Scarica il file selezionato
            Dim Colonna As Long
            Dim Riga As Long
            
            If ListView1.ListItems.count <> 0 Then
                Colonna = GetNumColDaIntestazione(ListView1, "upUrl") - 1
                Riga = ListView1.SelectedItem.index
                DownloadFileDialog (GetValoreCella(ListView1, Riga, Colonna))
                ' Imposto tutti i file .bmp come file nascosti per evitare che vengano visualizzati nella galleria dei cellulari nokia
                Call NascondiFile("*.bmp", Var(PoiScaricati).Valore)
            End If
            
        Case Is = 31 ' Apri pagina web con tutti i poi
            ShellExecute hwnd, "open", "http://www.poigps.com/modules.php?name=Downloads&d_op=search&query=%&min=0&orderby=dateD&show=2000", vbNullString, vbNullString, SW_SHOW

        Case Is = 49 ' Numera le righe della lista
            Call NumeraListView(ListView1)

        Case Is = 50
            Call AutoSizeColonne(ListView1)
        
        Case Is = 59
            Call ChangeSelezMultipla(ListView1)
        
        Case Is = 60
            Call CancellaRiga(ListView1)
        
        Case Is = 70
            Call CancellaRiga(ListView1, , True)
            
        Case Is = 80 ' Seleziona tutto
            SetCeckedListView ListView1, False, True
        
        Case Is = 81 ' Deseleziona tutto
            SetCeckedListView ListView1, False, False
            
    End Select

    TotaleRigheListView ListView1

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub txtListBox_Change()
    With txtListBox
        If .Text = "" Then
            .Visible = False
            On Error Resume Next
            ListView1.SetFocus
            Err.Clear
        Else
            .Visible = True
            .ZOrder
            If Left(.Text, 1) <> " " Then CercaEvidenzia ListView1, .Text
        End If
        .Width = (TextWidth(.Text) * 1.5) + 500
        .Move (ListView1.Left + (ListView1.Width / 2)) - .Width / 2, (ListView1.Top + (ListView1.Height / 2)) - .Height / 2
    End With
    DoEvents
End Sub

Private Sub txtListBox_GotFocus()
    ListView1.SetFocus
End Sub

Private Sub txtProxyPassword_LostFocus()
    Call ImpostazioniProxy(True)
End Sub

Private Sub txtProxyPort_LostFocus()
    Call ImpostazioniProxy(True)
End Sub

Private Sub txtProxyServer_LostFocus()
    Call ImpostazioniProxy(True)
End Sub

Private Sub txtProxyUserName_LostFocus()
    Call ImpostazioniProxy(True)
End Sub

Private Sub wbrSource_StatusTextChange(ByVal Text As String)
    On Error Resume Next
    If FormPronta = True Then StatusBar1.Panels(2).Text = Text
End Sub

Public Sub CaricaArrElencoFile()
    Dim cnt As Long
    Dim cntCol As Long
    Dim nColonne As Long
        
    ' Se non ci sono righe esco
    If ListView1.ListItems.count = 0 Then Exit Sub
    
    nColonne = ListView1.ColumnHeaders.count - 1

    ' Cancello l'array
    ReDim arrElencoFile(ListView1.ListItems.count)
    
    ' Scorro tutte le righe.....
    For cnt = 1 To ListView1.ListItems.count
        arrElencoFile(cnt - 1).Descrizione = Trim(ListView1.ListItems(cnt).SubItems(GetNumColDaIntestazione(ListView1, "Descrizione") - 1))
        arrElencoFile(cnt - 1).url = Trim(ListView1.ListItems(cnt).SubItems(GetNumColDaIntestazione(ListView1, "URL.") - 1))
        arrElencoFile(cnt - 1).upURL = Trim(ListView1.ListItems(cnt).SubItems(GetNumColDaIntestazione(ListView1, "upURL") - 1))
        arrElencoFile(cnt - 1).Formato = Trim(ListView1.ListItems(cnt).SubItems(GetNumColDaIntestazione(ListView1, "Formato") - 1))
        arrElencoFile(cnt - 1).Mappa = Trim(ListView1.ListItems(cnt).SubItems(GetNumColDaIntestazione(ListView1, "Mappa") - 1))
        arrElencoFile(cnt - 1).Immagine = Trim(ListView1.ListItems(cnt).SubItems(GetNumColDaIntestazione(ListView1, "Immagine") - 1))
        arrElencoFile(cnt - 1).Note = Trim(ListView1.ListItems(cnt).SubItems(GetNumColDaIntestazione(ListView1, "Note") - 1))
    Next
    
End Sub

