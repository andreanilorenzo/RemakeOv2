VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRimuoviDuplicati 
   Caption         =   "Rimuovi Duplicati"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   Icon            =   "frmRimuoviDuplicati.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Parametro4 
      Caption         =   "Parametro 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6000
      TabIndex        =   22
      Top             =   120
      Width           =   1815
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   1575
         TabIndex        =   23
         Top             =   240
         Width           =   1575
         Begin VB.ComboBox cmbColonneList 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Text            =   "cmbColonneList"
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox cekPar4 
            Caption         =   "Attiva"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin VB.CheckBox cekCaricaLista 
      Caption         =   "Carica Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8280
      TabIndex        =   21
      ToolTipText     =   "Quando attivato la lista viene caricata dalla lista nella finestra ""Crea e Modifica"" ogni volta che viene modificato un parametro"
      Top             =   720
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox cekEscludiDescrizione 
      Caption         =   "Escludi Descrizione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   12
      Top             =   1440
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin MSComctlLib.ImageList imglRunSearch 
      Left            =   2520
      Top             =   3240
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
            Picture         =   "frmRimuoviDuplicati.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":30F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":4DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":5114
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":526E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":53C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":5522
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":583C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":668E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":74E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":838A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":91DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":B98E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":BCA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":CAFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":CE14
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":E46E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":E788
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":EAA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":EDBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":F0D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":F9B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":1028A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":110DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":11F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":12D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":1365A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":144AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":161B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":164D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":17322
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":19AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":1C286
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":1DF90
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":20742
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2244C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":22D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":23B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":251D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":25AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":26386
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":271D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":28832
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":29684
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2999E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2A278
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2A592
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2A8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2ABC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2AD20
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2B03A
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2B354
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2B66E
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2B988
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2BCA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2BFBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2C116
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2C270
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRimuoviDuplicati.frx":2C3CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancellaDuplicati 
      Caption         =   "Cancella &duplicati"
      Height          =   495
      Left            =   8160
      TabIndex        =   20
      Top             =   1080
      Width           =   1695
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
      Left            =   1560
      ScaleHeight     =   990
      ScaleWidth      =   2295
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Frame Parametro3 
      Caption         =   "Parametro 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   1815
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   1575
         TabIndex        =   16
         Top             =   240
         Width           =   1575
         Begin VB.TextBox txtDistanza 
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
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Text            =   "txtDistanza"
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox cekDistanza 
            Caption         =   "Distanza < metri"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   17
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton cmdCercaDuplicati 
      Caption         =   "&Cerca duplicati"
      Height          =   495
      Left            =   8160
      TabIndex        =   13
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "&Chiudi  [Esc]"
      Height          =   1215
      Left            =   10080
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6675
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   11774
      View            =   2
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
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
   Begin VB.Frame Parametro2 
      Caption         =   "Parametro 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   1815
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   1575
         TabIndex        =   8
         Top             =   240
         Width           =   1575
         Begin VB.CheckBox cekLongitudine 
            Caption         =   "Longitudine"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox cekDescrizione2 
            Caption         =   "Descrizione"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox cekLatitudine 
            Caption         =   "Latitudine"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Parametro1 
      Caption         =   "Parametro 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      Begin VB.PictureBox Picture1b 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   1575
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         Begin VB.CheckBox cekCAP 
            Caption         =   "CAP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   1455
         End
         Begin VB.CheckBox cekDescrizione1 
            Caption         =   "Descrizione"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox cekIndirizzo 
            Caption         =   "Indirizzo"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1455
         End
      End
   End
   Begin VB.Label lblDup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   1920
      Width           =   2655
   End
End
Attribute VB_Name = "frmRimuoviDuplicati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim siEsce As Boolean
Dim CurrRow As Long
Dim Parametro As Integer
Dim cekEdit As Boolean

Dim ArrayOv2PoiRecTUTTI() As Ov2FileTy ' Array che contiene tutti i record da analizzare

' Classe per il menu PopUp
Dim mnu As clsMenu

Private Sub cekCAP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cekEdit = True
End Sub

Private Sub cekCAP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cekEdit = False
End Sub

Private Sub cekCAP_Click()

    If cekEdit = True Then
        If cekCAP.Value = 1 Then
            Parametro = 1
        Else
            Parametro = 2
        End If
        Call ControllaParametro
    End If

End Sub

Private Sub cekCaricaLista_Click()
    
    If cekCaricaLista.Value = 1 Then Call ControllaParametro
    
End Sub

Private Sub cekDistanza_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cekEdit = True
End Sub

Private Sub cekDistanza_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cekEdit = False
End Sub

Private Sub cekDistanza_Click()

    If cekEdit = True Then
        If cekDistanza.Value = 1 Then
            Parametro = 3
        Else
            Parametro = 1
        End If
        Call ControllaParametro
    End If

End Sub

Private Sub cekEscludiDescrizione_Click()

    cekDescrizione1.Value = Abs(Not cekEscludiDescrizione.Value)
    cekDescrizione2.Value = Abs(Not cekEscludiDescrizione.Value)
    Call ControllaParametro
    
End Sub

Private Sub cekLatitudine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cekEdit = True
End Sub

Private Sub cekLatitudine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cekEdit = False
End Sub
Private Sub cekLatitudine_Click()
    
    If cekEdit = True Then
         If cekLatitudine.Value = 1 Then
             Parametro = 2
         Else
             Parametro = 1
        End If
        Call ControllaParametro
    End If

End Sub

Private Sub cekPar4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cekEdit = True
End Sub

Private Sub cekPar4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cekEdit = False
End Sub

Private Sub cekPar4_Click()

    If cekEdit = True Then
        If cekPar4.Value = 1 Then
            Parametro = 4
        Else
            Parametro = 1
        End If
        Call ControllaParametro
    End If

End Sub

Private Sub cmbColonneList_GotFocus()
    cekEdit = True
End Sub

Private Sub cmbColonneList_LostFocus()
    cekEdit = False
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyZ Then
        Call CercaChecked(ListView1, False)
        
    ElseIf KeyCode = vbKeyX Then
        Call CercaChecked(ListView1, True)
        
    ' Per cancellare le righe
    ElseIf KeyCode = 46 Then 'Se il tasto premuto è Canc cancello la riga
        ' Mi assicuro che la ListView non è vuota
        If Not IsNull(ListView1.SelectedItem) And Not ListView1.SelectedItem Is Nothing Then
            ListView1.ListItems.Remove (ListView1.SelectedItem.index)
        End If
    
    ElseIf KeyCode = vbKeyC Then 'Se il tasto premuto è C cancello la riga se è ceckata
        ' Mi assicuro che la ListView non è vuota
        If Not IsNull(ListView1.SelectedItem) And Not ListView1.SelectedItem Is Nothing Then
            If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
                ListView1.ListItems.Remove (ListView1.SelectedItem.index)
            End If
        End If
    End If

End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result
    Dim i As Integer
    Dim nCampi As Integer
    
    Screen.MousePointer = vbHourglass
    
    ListView1.ListItems.Clear
    
    If Data.GetFormat(vbCFFiles) Then
        For i = 1 To 1 'Data.Files.count
            ClickPopUp 20, (Data.Files(i))
        Next
    End If

    Screen.MousePointer = vbDefault

End Sub

Private Sub txtDistanza_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Parametro = 3
        ControllaParametro
    End If
    
End Sub

Private Sub cmbColonneList_Click()
    
    If cekEdit = True Then
        DoEvents
        Parametro = 4
        ControllaParametro
    End If
    
End Sub

Private Sub ControllaParametro()
    
    Screen.MousePointer = vbHourglass

    cekEdit = False

    Select Case Parametro
        Case 1
            Parametro1.FontBold = True
            Parametro2.FontBold = False
            Parametro3.FontBold = False
            Parametro4.FontBold = False
            
            cekCAP.Value = 1
            cekDescrizione1.Value = Abs(Not cekEscludiDescrizione.Value)
            cekIndirizzo.Value = 1
            
            cekLatitudine.Value = 0
            cekLongitudine.Value = 0
            cekDescrizione2.Value = 0
            
            cekDistanza.Value = 0
            
            cekPar4.Value = 0
        
        Case 2
            Parametro1.FontBold = False
            Parametro2.FontBold = True
            Parametro3.FontBold = False
            Parametro4.FontBold = False

            cekCAP.Value = 0
            cekDescrizione1.Value = 0
            cekIndirizzo.Value = 0
            
            cekLatitudine.Value = 1
            cekLongitudine.Value = 1
            cekDescrizione2.Value = Abs(Not cekEscludiDescrizione.Value)
            
            cekDistanza.Value = 0
            
            cekPar4.Value = 0

        Case 3
            Parametro1.FontBold = False
            Parametro2.FontBold = False
            Parametro3.FontBold = True
            Parametro4.FontBold = False
            
            cekCAP.Value = 0
            cekDescrizione1.Value = 0
            cekIndirizzo.Value = 0
            
            cekLatitudine.Value = 0
            cekLongitudine.Value = 0
            cekDescrizione2.Value = 0
            
            cekDistanza.Value = 1

            cekPar4.Value = 0
            
        Case 4
            Parametro1.FontBold = False
            Parametro2.FontBold = False
            Parametro3.FontBold = False
            Parametro4.FontBold = True
            
            cekCAP.Value = 0
            cekDescrizione1.Value = 0
            cekIndirizzo.Value = 0
            
            cekLatitudine.Value = 0
            cekLongitudine.Value = 0
            cekDescrizione2.Value = 0
            
            cekDistanza.Value = 0
            
            cekPar4.Value = 1
            
    End Select
    
    cekDistanza.Caption = "Distanza < metri: " & txtDistanza.Text
    
    lblDup.Caption = "....................."
    
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Prima devi caricare il file....", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    ' Prevent the ListView control from updating on screen - this is to hide the changes being made to the listitems and also to speed up the sort
    LockWindowUpdate ListView1.hwnd
    
    If cekCaricaLista.Value = 1 Then
        ArrayOv2PoiRec = ArrayOv2PoiRecTUTTI
    Else
        ' Carico i dati dalla listview
        Call ArrayOv2PoiRecDaListView(ListView1, ArrayOv2PoiRec)
    End If
        
    If Parametro = 1 Or Parametro = 2 Then
        ' Ordino l'array in base al Parametro
        If QuickSortArrayOv2PoiRec(Parametro, ArrayOv2PoiRec) = True Then
            ' Carico l'array nella listview
            Call ArrayOv2PoiRecInListView(ListView1, ArrayOv2PoiRec, True)
        End If
        
    ElseIf Parametro = 3 Then
        Call SetCeckedListView(ListView1, True)
        Call OrdinaColonnaByNumeroCol(ListView1, 1, 1)
        If QuickSortArrayOv2PoiRec(Parametro, ArrayOv2PoiRec) = True Then
        End If
        
    ElseIf Parametro = 4 Then
        Call SetCeckedListView(ListView1, True)
        Call OrdinaColonna(ListView1, ListView1.ColumnHeaders.Item(cmbColonneList.ListIndex + 2), 1)
    End If

    Call AutoSizeColonne(ListView1, 8)
    Call AutoSizeColonne(ListView1, 9)
    
    ' Unlock the list window so that the OCX can update it
    LockWindowUpdate 0&

    cmdCercaDuplicati.Enabled = True
    cmdCancellaDuplicati.Enabled = True
    cekCaricaLista.Value = 0
    cekCaricaLista.Enabled = True
    
    Me.Refresh
    Screen.MousePointer = vbDefault
    
End Sub

Private Function QuickSortArrayOv2PoiRec(Parametro As Integer, ArrayOv2PoiRec1() As Ov2FileTy) As Boolean
    Dim i As Long
    Dim j As Long
    Dim gap As Long
    Dim maxrec As Long
    Dim arrTmp As Ov2FileTy
    Dim i_key As String
    Dim j_key As String
    Dim lstCountRow As Long
    Dim Distanza As Double
    Dim timAvvio As Date
    Dim slen As Integer
    
    'Il sistema per trovare duplicati si basa sul riordino e
    'su un ciclo di scasione sull'array ordinato per testare
    'se la riga corrente e uguale alla corrente + 1 .
    'Le righe uguali nelle colonne selezionate (le stesse prese in
    'considerazione per il sort) sono marcate per l'eventale eliminazione.
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Esci
    
    slen = 64
    maxrec = UBound(ArrayOv2PoiRec1)
    gap = maxrec / 2
    lblDup.Caption = "....................."
    lstCountRow = ListView1.ListItems.Count
    timAvvio = Now
    i = 0
    
    Select Case Parametro
        Case Is = 1 ' Cap - Indirizzo - Descrizione
            While gap > 0
                i = gap
                While i < maxrec 'Esegue il ciclo finché "i è minore di maxRec" dà come risultato True
                    j = i - gap
                    While j >= 0
                        i_key = Left$(ArrayOv2PoiRec1(j + gap).iCap & String$(5, 0), 5)
                        i_key = i_key & Left$(ArrayOv2PoiRec1(j + gap).iIndirizzo & Space$(40), 40)
                        If cekEscludiDescrizione.Value = 0 Then i_key = i_key & Left$(ArrayOv2PoiRec1(j + gap).hTy2descrizione & Space$(slen), slen)
                        
                        j_key = Left$(ArrayOv2PoiRec1(j).iCap & String$(5, 0), 5)
                        j_key = j_key & Left$(ArrayOv2PoiRec1(j).iIndirizzo & Space$(40), 40)
                        If cekEscludiDescrizione.Value = 0 Then j_key = j_key & Left$(ArrayOv2PoiRec1(j).hTy2descrizione & Space$(slen), slen)
                        
                        If j_key <= i_key Then
                            j = -1
                        End If
                        
                        DoEvents
                        
                        If j >= 0 Then
                            arrTmp = ArrayOv2PoiRec1(j)
                            ArrayOv2PoiRec1(j) = ArrayOv2PoiRec1(j + gap)
                            ArrayOv2PoiRec1(j + gap) = arrTmp
                            j = j - gap
                        End If
                        
                    Wend
                    
                    If i Mod 2000 = 0 Then DoEvents
                    If siEsce = True Then GoTo Esci

                    i = i + 1
                Wend
                
                Call ElaboraAvanzamentoRecord(timAvvio, maxrec, i)
                i = i + 1

                gap = gap / 2
            Wend
        
        Case Is = 2 ' Latitudine - Longitudine - Descrizione
            While gap > 0
                i = gap
                While i < maxrec 'Esegue il ciclo finché "i è minore di maxRec" dà come risultato True
                    j = i - gap
                    While j >= 0
                        i_key = FormatNumber(ArrayOv2PoiRec1(j + gap).fTy2PoiLatitude, 6)
                        i_key = i_key & FormatNumber(ArrayOv2PoiRec1(j + gap).gTy2PoiLongitude, 6)
                        If cekEscludiDescrizione.Value = 0 Then i_key = i_key & Left$(ArrayOv2PoiRec1(j + gap).hTy2descrizione & Space$(slen), slen)
                        
                        j_key = FormatNumber(ArrayOv2PoiRec1(j).fTy2PoiLatitude, 6)
                        j_key = j_key & FormatNumber(ArrayOv2PoiRec1(j).gTy2PoiLongitude, 6)
                        If cekEscludiDescrizione.Value = 0 Then j_key = j_key & Left$(ArrayOv2PoiRec1(j).hTy2descrizione & Space$(slen), slen)
                        
                        If j_key <= i_key Then
                            j = -1
                        End If
                        
                        If j >= 0 Then
                            arrTmp = ArrayOv2PoiRec1(j)
                            ArrayOv2PoiRec1(j) = ArrayOv2PoiRec1(j + gap)
                            ArrayOv2PoiRec1(j + gap) = arrTmp
                            j = j - gap
                        End If
                        
                    Wend
                    
                    If i Mod 2000 = 0 Then DoEvents
                    If siEsce = True Then GoTo Esci

                    i = i + 1
                Wend
                
                Call ElaboraAvanzamentoRecord(timAvvio, maxrec, i)
                i = i + 1

                gap = gap / 2
            Wend
            

        Case Is = 3 ' Distanza Metri
            ' La funzione per trovare le coordinate che si trovano
            ' in una zona inferiore a metri xxx l'ho implementata con un doppio ciclo.
            ' Per ogni POI calcolo la distanza su tutti gli altri con il
            ' ciclo secondario se la distanza e' <= ai metri previsti
            ' li marco il numero ordinale del poi che sto' analizzando.
            Call SetCeckedListView(ListView1)
            
            ' Aggiungo la colonna (se non esiste)
            If GetNumColDaIntestazione(ListView1, " 12 SubOrdine") = 0 Then ListView1.ColumnHeaders.Add , "x13", " 12 SubOrdine"

            ' Prevent the ListView control from updating on screen - this is to hide the changes being made to the listitems and also to speed up the sort
            LockWindowUpdate ListView1.hwnd

            For i = 1 To lstCountRow
                For j = 1 To lstCountRow
                    If j <> i Then
                        Distanza = CalculateDistance(ListView1.ListItems(i).SubItems(8), ListView1.ListItems(i).SubItems(9), ListView1.ListItems(j).SubItems(8), ListView1.ListItems(j).SubItems(9))
                        If Distanza <= CDbl(txtDistanza.Text) Then
                            ListView1.ListItems(j).SubItems(12) = Format$(i, "000000")
                            ListView1.ListItems(i).SubItems(12) = Format$(i, "000000")
                        Else
                            If ListView1.ListItems(i).SubItems(12) = "" Then ListView1.ListItems(i).SubItems(12) = Format$(0, "000000")
                        End If
                    End If
                    
                    If j Mod 1000 = 0 Then DoEvents
                    
                    If siEsce = True Then
                        lblDup.Caption = "Ricerca annullata dall'utente....."
                        GoTo Esci
                    End If
                    
                Next
                
                Call ElaboraAvanzamentoRecord(timAvvio, lstCountRow, i + 1, 30)
            Next
            
            ' Ordino per l'elemento che contiene la marcatura
            Call OrdinaColonnaByNumeroCol(ListView1, 13, 1)

    End Select
    
    QuickSortArrayOv2PoiRec = True
    
    Exit Function

Esci:
    QuickSortArrayOv2PoiRec = False

End Function


Private Sub ElaboraAvanzamentoRecord(ByVal timAvvio As Date, ByVal TotaleRecord As Long, ByVal i As Long, Optional StepTime As Integer = 70, Optional QuckSort As Boolean = False)
    Dim secTrascorsi As Double
    Static UltimaStima As String

    If i >= TotaleRecord Then
        UltimaStima = "calcolo in corso....."
        lblDup.Caption = i & " di " & TotaleRecord & " - Tempo totale: " & ConvertiSecInGiorni(TempoTrascorso(timAvvio, Now, True))
        lblDup.Refresh

    ElseIf i Mod StepTime = 0 Then
        secTrascorsi = TempoTrascorso(timAvvio, Now, True)
        UltimaStima = StimaTempoRestante(secTrascorsi, i, TotaleRecord)
        lblDup.Caption = i & " di " & TotaleRecord & " - Tempo restante: " & UltimaStima
        DoEvents
    
    ElseIf i Mod 5 = 0 Then
        lblDup.Caption = i & " di " & TotaleRecord & " - Tempo restante: " & UltimaStima
        lblDup.Refresh
    
    ElseIf i = 0 Or i = 1 Or i = 2 Then
        UltimaStima = "calcolo in corso....."
        lblDup.Caption = i & " di " & TotaleRecord & " - Tempo restante: " & UltimaStima
        lblDup.Refresh
                
    End If

End Sub

Private Sub cmdCercaDuplicati_Click()
    Dim R As Long
    Dim lstCountRow As Long
    Dim r2 As Long
    Dim dup As Integer
    Dim timAvvio As Date
    Dim secTrascorsi As Double
    Dim UltimaStima As String
    
    timAvvio = Now
    
    dup = 0
    lblDup.Caption = "....................."
    cmdCercaDuplicati.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
            
    lstCountRow = ListView1.ListItems.Count
    
    Select Case Parametro
        Case Is = 1 ' Cap - Descrizione - Indirizzo
            ' Prevent the ListView control from updating on screen - this is to hide the changes being made to the listitems and also to speed up the sort
            LockWindowUpdate ListView1.hwnd

            For R = 1 To lstCountRow - 1
                    ' 1 = descrizione - 2 = Indirirzzo - 3 = Cap
                If (ListView1.ListItems(R + 1).SubItems(1) = ListView1.ListItems(R).SubItems(1) _
                    And ListView1.ListItems(R + 1).SubItems(2) = ListView1.ListItems(R).SubItems(2) _
                    And ListView1.ListItems(R + 1).SubItems(3) = ListView1.ListItems(R).SubItems(3) _
                    And cekEscludiDescrizione.Value = 0) Or _
                    (ListView1.ListItems(R + 1).SubItems(2) = ListView1.ListItems(R).SubItems(2) _
                    And ListView1.ListItems(R + 1).SubItems(3) = ListView1.ListItems(R).SubItems(3) _
                    And cekEscludiDescrizione.Value = 1) Then
                        ' Imposto il cekBox
                        ListView1.ListItems.Item(R + 1).Checked = True
                        dup = dup + 1
                End If
                Call ElaboraAvanzamentoRecord(timAvvio, lstCountRow, R)
            Next
        
        Case Is = 2 ' Latitudine - Longitudine - Descrizione
            ' Prevent the ListView control from updating on screen - this is to hide the changes being made to the listitems and also to speed up the sort
            LockWindowUpdate ListView1.hwnd
            
            For R = 1 To lstCountRow - 1
                    ' 1 = descrizione - 8 = Lat. - 9 = Long.
                If (ListView1.ListItems(R + 1).SubItems(1) = ListView1.ListItems(R).SubItems(1) _
                    And ListView1.ListItems(R + 1).SubItems(8) = ListView1.ListItems(R).SubItems(8) _
                    And ListView1.ListItems(R + 1).SubItems(9) = ListView1.ListItems(R).SubItems(9) _
                    And cekEscludiDescrizione.Value = 0) Or _
                    (ListView1.ListItems(R + 1).SubItems(8) = ListView1.ListItems(R).SubItems(8) _
                    And ListView1.ListItems(R + 1).SubItems(9) = ListView1.ListItems(R).SubItems(9) _
                    And cekEscludiDescrizione.Value = 1) Then
                        ' Imposto il cekBox
                        ListView1.ListItems.Item(R + 1).Checked = True
                        dup = dup + 1
                End If
                Call ElaboraAvanzamentoRecord(timAvvio, lstCountRow, R)
            Next
        
        Case Is = 3 ' Distanza Metri
            ' Procedo con la segnalazione dei doppi
            For R = 1 To ListView1.ListItems.Count - 1
                If ListView1.ListItems(R).SubItems(12) <> 0 And (ListView1.ListItems(R + 1).SubItems(12) = ListView1.ListItems(R).SubItems(12)) Then
                    ListView1.ListItems.Item(R + 1).Checked = True
                    dup = dup + 1
                End If
                Call ElaboraAvanzamentoRecord(timAvvio, lstCountRow, R)
            Next

        Case Is = 4 ' Colonna
            ' Prevent the ListView control from updating on screen - this is to hide the changes being made to the listitems and also to speed up the sort
            LockWindowUpdate ListView1.hwnd
            
            For R = 1 To lstCountRow - 1
                If ListView1.ListItems(R + 1).SubItems(cmbColonneList.ListIndex + 1) = ListView1.ListItems(R).SubItems(cmbColonneList.ListIndex + 1) Then
                    ' Imposto il cekBox
                    ListView1.ListItems.Item(R + 1).Checked = True
                    dup = dup + 1
                End If
                Call ElaboraAvanzamentoRecord(timAvvio, lstCountRow, R)
            Next
            
    End Select

    Call ControllaCheck(ListView1)
    
    If dup <> 0 Then
        lblDup.Caption = "Trovati n°" & dup & " duplicati - Utilizza i tasti ""z"" e ""x"" per cercare i duplicati nella lista e ""c"" per cancellare"
    Else
        lblDup.Caption = "Non è stato trovato nessun duplicato"
    End If
    
    ListView1.SetFocus

Esci:
    ' Unlock the list window so that the OCX can update it
    LockWindowUpdate 0&
    
    cmdCercaDuplicati.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdCancellaDuplicati_Click()
    Dim ret As Integer
    
    ret = CancellaCeckedListView(ListView1)
    
    If ret = 0 Then
        MsgBox "Non sono stati trovati record selezionati.", vbInformation, App.ProductName
    Else
        MsgBox "Sono stati rimossi n°" & ret & " record duplicati.", vbInformation, App.ProductName
    End If
    
End Sub

Private Sub cmdEsci_Click()
    siEsce = True
    DoEvents
    DoEvents
    Unload Me
End Sub

Private Sub Form_Load()
    Dim LarghezzaColonne As Variant
    Dim cnt As Integer

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    LarghezzaColonne = Array(0, 1100, 2800, 2200, 980, 1600, 650, 1200, 1200, 980, 980, 3000, 1200)
    
    With ListView1
        .HideSelection = False
        .MultiSelect = Var(SelezMultipla).Valore
        .FullRowSelect = True
        .View = lvwReport
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
    
    cmbColonneList.Text = ""
    Call BloccaComboBox(cmbColonneList)

    cmdCercaDuplicati.Enabled = False
    cmdCancellaDuplicati.Enabled = False
    cekCaricaLista.Enabled = False
    
    txtDistanza.Text = 50
    siEsce = False

    
    If FormIsLoad("frmWeb") = True Then
        ' Carico i dati dalla listview della form frmWeb
        Call ArrayOv2PoiRecDaListView(frmWeb.ListView1, ArrayOv2PoiRecTUTTI)
    
        For cnt = 1 To ListView1.ColumnHeaders.Count - 1
            cmbColonneList.AddItem Trim$(ListView1.ColumnHeaders(cnt + 1).Text)
        Next
        
        cekEdit = False
        cmbColonneList.ListIndex = 0
        
        LockWindowUpdate ListView1.hwnd
        Call ArrayOv2PoiRecInListView(ListView1, ArrayOv2PoiRecTUTTI, True)
        TotaleRigheListView ListView1
        LockWindowUpdate 0&
        
        frmWeb.Visible = False
    End If
    
    Exit Sub
    
Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione frmRimuoviDuplicati.Form_Load."

End Sub

Private Sub Form_Activate()
    frmMain.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    siEsce = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    ' Make sure the form is not minimized
    If Me.WindowState <> vbMinimized Then
        ' Maintain a minimum height and width in order to not set a negative width or height
        If Me.Height < 9500 Or Me.Width < 12500 Then
            Me.Height = 9500
            Me.Width = 12500
        End If
        
        ' Centro i controlli nella form
        'lblDup.Move (Me.ScaleWidth - lblDup.Width) / 2
        
        lblDup.Move 0, lblDup.Top, Me.ScaleWidth, lblDup.Height
        ListView1.Move 0, ListView1.Top, Me.ScaleWidth, Me.ScaleHeight - ListView1.Top

    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    
    If FormIsLoad("frmWeb") = True Then
        frmWeb.Visible = True
        frmWeb.SetFocus
    Else
        frmMain.Visible = True
        frmMain.SetFocus
    End If
    
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call OrdinaColonna(ListView1, ColumnHeader)
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ControllaRigaChecked(ListView1, Item)
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        CurrRow = GetRigaSelezionata(ListView1, X, Y)
    End If

End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ListaVuota As Boolean
    
    If ListView1.ListItems.Count = 0 Then
        ListaVuota = True
    Else
        ListaVuota = False
    End If
    
    ' Per il PopUp menu
    If Button = vbRightButton Then
        Set mnu = New clsMenu
        With mnu
            If ListaVuota = False Then
                Dim submnu1 As clsMenu: Set submnu1 = New clsMenu
                With submnu1
                    .Caption = "Mostra sulla cartina"
                    .AddItem 101, "www.multimap.com (" & CurrRow & ")", , , , ListaVuota
                    .AddItem 102, "www.mapquest.com (" & CurrRow & ")", , , , ListaVuota
                End With
                '
                .AddItem 10, submnu1
                .AddItem 0, "-"
            End If
            '
            .AddItem 20, "Apri...."
            .AddItem 30, "Salva...", , , , ListaVuota
            .AddItem 0, "-"
            '
            .AddItem 40, "Numera le righe della lista", , , , ListaVuota
            .AddItem 0, "-"
            '
            .AddItem 50, "AutoSize larghezza colonne", , , , ListaVuota
            .AddItem 59, "Selezione righe multiple", , Var(SelezMultipla).Valore, , ListaVuota
            .AddItem 60, "Cancella riga selezionata", , , , ListaVuota
            .AddItem 70, "Cancella righe selezionate", , , , ListaVuota
            '
            '.AddItem 80, "Nuova Colonna", , , , , True
            ClickPopUp .TrackPopup(Me.hwnd)
        End With

    End If

End Sub

Public Sub ClickPopUp(ValoreCliccato As Long, Optional FileDaAprire As String = "")
    Dim ret As Variant
    Dim cnt As Integer
    Dim strTmp As String
    Dim PatchNomeFile As String
    Dim CancellaListView As Boolean
    Dim SaltaRighe As Integer
    
    PatchNomeFile = ""
    
    Select Case Left(ValoreCliccato, 2)
        Case Is = 10 ' Mostra posizione sulla cartina
                Screen.MousePointer = vbHourglass
                DoEvents
            If ValoreCliccato = 101 Then
                strTmp = MostraPosizione(ListView1, "multimap", CurrRow)
                If strTmp <> "" Then ShellExecute hwnd, "open", strTmp, vbNullString, vbNullString, SW_SHOW
            ElseIf ValoreCliccato = 102 Then
                strTmp = MostraPosizione(ListView1, "mapquest", CurrRow)
                If strTmp <> "" Then ShellExecute hwnd, "open", strTmp, vbNullString, vbNullString, SW_SHOW
            ElseIf ValoreCliccato = 102 Then
                strTmp = MostraPosizione(ListView1, "map24", CurrRow)
                If strTmp <> "" Then ShellExecute hwnd, "open", strTmp, vbNullString, vbNullString, SW_SHOW
            End If
                Screen.MousePointer = vbDefault
            
        Case Is = 20 ' Apri
            ' Prevent the ListView control from updating on screen - this is to hide the changes being made to the listitems and also to speed up the sort
            LockWindowUpdate ListView1.hwnd
            If ImportaDati(Me.hwnd, ListView1, FileDaAprire, , True, True, , , , "FileDatiPOI") = False Then
                Me.Caption = "Rimuovi Duplicati"
                ret = MsgBox("Non è stato possibile importare il file!", vbOKOnly)
            Else
                Me.Caption = "Rimuovi Duplicati: " & FileDaAprire
            End If
            ' Unlock the list window so that the OCX can update it
            LockWindowUpdate 0&
            DoEvents
            Call ArrayOv2PoiRecDaListView(ListView1, ArrayOv2PoiRecTUTTI)
        
            For cnt = 1 To ListView1.ColumnHeaders.Count - 1
                cmbColonneList.AddItem Trim$(ListView1.ColumnHeaders(cnt + 1).Text)
            Next
            
            cmbColonneList.ListIndex = 0
            cekCaricaLista.Value = 0
            lblDup.Caption = "....................."

            Screen.MousePointer = vbDefault

        Case Is = 30 ' Salva
            ' Se esiste, cancello la colonna provvisoria
            ret = GetNumColDaIntestazione(ListView1, " 12 SubOrdine")
            If ret <> 0 Then
                ListView1.ColumnHeaders.Remove (ret)
            End If
            Call ExportaDati(Me.hwnd, , NomeFileAperto, , True, , Var(CampiRmkFile).Valore & "FileDatiPOI" & ";;" & Now, ListView1)
            Screen.MousePointer = vbDefault
            
        Case Is = 40 ' Numera le righe della lista
            Call NumeraListView(ListView1)
        
        Case Is = 50
            Call AutoSizeColonne(ListView1)
        
        Case Is = 59
            Call ChangeSelezMultipla(ListView1)
        
        Case Is = 60
            Call CancellaRiga(ListView1)
        
        Case Is = 70
            Call CancellaRiga(ListView1, , True)
            
    End Select

    TotaleRigheListView ListView1
    
End Sub
