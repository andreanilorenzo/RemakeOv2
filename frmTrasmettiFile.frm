VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrasmettiFile 
   Caption         =   "Aggiorna POI sul telefono"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   Icon            =   "frmTrasmettiFile.frx":0000
   LinkTopic       =   "frmTrasmettiFile"
   ScaleHeight     =   6165
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   5790
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitH 
      BackColor       =   &H80000002&
      Height          =   75
      Index           =   0
      Left            =   360
      ScaleHeight     =   15
      ScaleWidth      =   9615
      TabIndex        =   18
      Top             =   3240
      Width           =   9675
   End
   Begin VB.TextBox txtTmp 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   7800
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "TxtTmp"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOutputs 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   855
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3120
      Width           =   10095
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
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.PictureBox picComandi 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   480
      ScaleHeight     =   1215
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   4440
      Width           =   9435
      Begin VB.CommandButton cmdVediCartella 
         Caption         =   "&<--"
         Height          =   315
         Left            =   9000
         TabIndex        =   17
         ToolTipText     =   "Visualizza il contenuto della cartella"
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtCerca 
         Alignment       =   2  'Center
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "txtCerca"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtDrive 
         Alignment       =   2  'Center
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   4200
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "E:\"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdSalva 
         Caption         =   "&Salva"
         Height          =   315
         Left            =   2880
         TabIndex        =   13
         Top             =   120
         Width           =   615
      End
      Begin VB.ComboBox cmbTelefoni 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   12
         Text            =   "cmbTelefoni"
         Top             =   120
         Width           =   2775
      End
      Begin VB.ComboBox cmbDirMap 
         Height          =   315
         Left            =   5760
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   3255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "&Elenco cartelle telefono"
         Height          =   315
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "Seleziona i file che sono presenti sul sito ma non sono presenti nella cartella del computer"
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdEsci 
         Cancel          =   -1  'True
         Caption         =   "&Esci  [Esc]"
         Height          =   375
         Left            =   6480
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox cekTrasfBMP 
         Caption         =   "Trasferisci anche BMP"
         Height          =   195
         Left            =   2640
         TabIndex        =   2
         ToolTipText     =   "Dopo aver scaricato i file apri automaticamente la finestra per effettuare il trattamento"
         Top             =   840
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton cmdTrasferisci 
         Caption         =   "&Trasferisci i file selezionati"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   1
         ToolTipText     =   "Seleziona i file che sono presenti sul sito ma non sono presenti nella cartella del computer"
         Top             =   480
         Width           =   2295
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
            Picture         =   "frmTrasmettiFile.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":30F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":4DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":5114
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":526E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":53C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":5522
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":583C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":668E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":74E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":838A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":91DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":B98E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":BCA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":CAFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":CE14
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":E46E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":E788
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":EAA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":EDBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":F0D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":F9B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":1028A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":110DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":11F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":12D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":1365A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":144AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":161B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":164D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":17322
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":19AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":1C286
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":1DF90
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":20742
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2244C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":22D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":23B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":251D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":25AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":26386
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":271D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":28832
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":29684
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2999E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2A278
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2A592
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2A8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2ABC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2AD20
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2B03A
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2B354
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2B66E
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2B988
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2BCA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2BFBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2C116
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2C270
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasmettiFile.frx":2C3CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2475
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   4366
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
   Begin VB.PictureBox pctNoFile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   840
      Picture         =   "frmTrasmettiFile.frx":2C524
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblDir 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "lblDir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5010
      MouseIcon       =   "frmTrasmettiFile.frx":4E1D7
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   360
      Width           =   465
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
      TabIndex        =   7
      Top             =   120
      Width           =   10215
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
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   10335
   End
End
Attribute VB_Name = "frmTrasmettiFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tyFile
    Nome As String
    Data As String
    Note As String
    File As String
    Estensione As String
    Immagine As String
End Type
Private ListaFile() As tyFile

Private NessunTelefono As String
Private FileTrovati As Integer
Private MancaNomeTel As Boolean
Private CartellaLavoro As String

Private WithEvents objPhnFile As cPhoneFile
Attribute objPhnFile.VB_VarHelpID = -1
Private WithEvents m_cSplitH As cSplitter
Attribute m_cSplitH.VB_VarHelpID = -1

Private Sub cmbTelefoni_Change()
    objPhnFile.NomeTelefono = cmbTelefoni.Text
End Sub

Private Sub cmbTelefoni_Click()
    cmbTelefoni_Change
End Sub

Private Sub cmdDir_Click()
    Dim strComand As String

    If cmbTelefoni.Text = NessunTelefono Then
        MsgBox "Prima devi inserire il nome del telefono...." & vbNewLine & "Vuota la casella del nome per connetterti all'ultimo telefono.", vbInformation, App.ProductName
        cmbTelefoni.SetFocus
    
    Else
        txtOutputs(0).Text = ""
        cmdDir.Enabled = False
        cmbDirMap.Clear
        DoEvents
        
        If cmbTelefoni.Text = "" Then
            MsgBox "Nome del telefono mancante." & vbNewLine & "RemakeOv2 tenterà di connettersi all'ultimo nome utilizzato.", vbInformation, App.ProductName
            MancaNomeTel = True
        End If

        ' Leggo le cartelle dal telefono
        objPhnFile.LeggiDir txtDrive.Text
        
        CercaDirMappa txtDrive.Text, txtOutputs(0).Text, "Italia"
        
        If GetNumeroRigheChecked(ListView1(0)) >= 1 And cmbDirMap.ListCount >= 1 Then
            cmdTrasferisci.Enabled = True
        Else
            cmdTrasferisci.Enabled = False
        End If
        cmdDir.Enabled = True
    End If
    
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdSalva_Click()
    Dim ret
    Dim cnt As Integer
    Dim strTmp As String
    Dim arrTmp() As String
    
    If cmbTelefoni.Text = NessunTelefono Then
        MsgBox "Non ci sono nomi da salvare nell'elenco dei telefoni.", vbInformation, App.ProductName
        cmbTelefoni.SetFocus
        Exit Sub
    End If
    
    ret = MsgBox("Vuoi salvare l'elenco dei telefoni?", vbInformation + vbYesNo, App.ProductName)
    If ret = vbYes Then
        ReDim arrTmp(cmbTelefoni.ListCount - 1 + 1)
        
        ' Scorro tutto il cmbTelefoni ed inserisco i dati nell'array arrTmp
        For cnt = 0 To cmbTelefoni.ListCount - 1
            If cmbTelefoni.List(cnt) <> "" Then arrTmp(cnt) = cmbTelefoni.List(cnt)
        Next
        
        ' Inserisco il valore scritto nella cmbTelefoni.Text
        If cmbTelefoni.Text <> NessunTelefono Then
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(cnt) = cmbTelefoni.Text
        End If
        
        ' Elimino dall'array eventuali valori duplicati
        Call FilterDuplicates(arrTmp)
        
        ' Scorro tutto l'array e preparo la stringa
        For cnt = 0 To UBound(arrTmp)
            If arrTmp(cnt) <> "" Then strTmp = strTmp & "|" & arrTmp(cnt)
        Next
        
        ' Tolgo il carattere | a sinistra della stringa
        If Left$(strTmp, 1) = "|" Then strTmp = Right$(strTmp, Len(strTmp) - 1)
        ' Scrivo i dati
        lVar(PhoneFileTele) = strTmp
    End If
    
    ' Carico i nuovi dati
    Call CaricaElencoTelefoni
    
End Sub

Private Sub cmdTrasferisci_Click()
    Dim cntRiga As Integer
    Dim Comando As String
    Dim FilePc As String
    Dim FileTel As String
    Dim NomeFile As String
    Dim strPhoneFile As String
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    Comando = "-write"
    txtOutputs(0).Text = "Inizio operazioni di trasmissione file " & Now
    cmdTrasferisci.Enabled = False
    
    ' Scorro tutte le righe.....
    For cntRiga = 1 To ListView1(0).ListItems.Count
        If ListView1(0).ListItems.Item(cntRiga).Checked = True Then
            ' Preparo le variabili
            NomeFile = FileNameFromPath(GetValoreCella(ListView1(0), cntRiga, GetNumColDaIntestazione(ListView1(0), "File") - 1))
            FilePc = """" & GetValoreCella(ListView1(0), cntRiga, GetNumColDaIntestazione(ListView1(0), "File") - 1) & """"
            FileTel = """" & cmbDirMap.Text & "\" & NomeFile & """"
            
            txtOutputs(0).Text = txtOutputs(0).Text & vbNewLine & "Invio del file " & NomeFile
            txtOutputs(0).Refresh
            
            ' Invio i file. ATTENZIONE: a questo comando PhoneFile non da alcun testo di risposta
            objPhnFile.ScriviFile FileTel, FilePc
            
            txtOutputs(0).Text = txtOutputs(0).Text & vbTab & vbTab & vbTab & " ....inviato"
            
            If cekTrasfBMP.Value = 1 Then
                ' Preparo le variabili
                NomeFile = FileNameFromPath(GetValoreCella(ListView1(0), cntRiga, GetNumColDaIntestazione(ListView1(0), "Immagine") - 1))
                FilePc = """" & GetValoreCella(ListView1(0), cntRiga, GetNumColDaIntestazione(ListView1(0), "Immagine") - 1) & """"
                FileTel = """" & cmbDirMap.Text & "\" & NomeFile & """"
                
                txtOutputs(0).Text = txtOutputs(0).Text & vbNewLine & "Invio del file " & NomeFile
                txtOutputs(0).Refresh
    
                ' Invio i file. ATTENZIONE: a questo comando PhoneFile non da alcun testo di risposta
                objPhnFile.ScriviFile FileTel, FilePc
                
                txtOutputs(0).Text = txtOutputs(0).Text & vbTab & vbTab & vbTab & " ....inviato"
            End If
        
        End If
    Next
    
    txtOutputs(0).Text = txtOutputs(0).Text & vbNewLine & "Fine operazioni di trasmissione file " & Now

    cmdTrasferisci.Enabled = True
    cmdTrasferisci.SetFocus
        
    Exit Sub
    
Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione ExportaDati."
    
End Sub

Private Sub cmdVediCartella_Click()
    
    cmdVediCartella.Enabled = False
    If cmbDirMap.Text <> "" Then
        txtOutputs(0).Text = ""
        txtOutputs(0).Refresh
        ' Leggo le cartelle dal telefono
        objPhnFile.LeggiDir cmbDirMap.Text
    End If
    cmdVediCartella.Enabled = True
    
End Sub

Private Sub Form_Activate()
    If FormIsLoad("frmMain") = True Then frmMain.Visible = False
End Sub

Private Sub Form_Load()
    Dim ret
    Dim sPhoneFileExe As String
    Dim okPhoneFileExe As String
    Dim ExeTest
    Dim LarghezzaColonne As Variant
    Dim ColonneListView1 As Variant
    Dim TagColonne As Variant
    Dim strTmp As String
    Dim cnt As Long
    
    Me.Caption = Me.Caption & " - Symbian S60 (1.x - 2.x)"
    sPhoneFileExe = Var(PhoneFileExe).Valore

    If FileExists(sPhoneFileExe) = False Then
        ExeTest = Split(Replace$(LCase$(sPhoneFileExe), "programmi", "program files") & "|" & Replace$(LCase$(sPhoneFileExe), "program files", "programmi"), "|", , vbTextCompare)
        For cnt = 0 To UBound(ExeTest)
            If FileExists(ExeTest(cnt)) = True Then
                lVar(PhoneFileExe) = ExeTest(cnt)
                Exit For
            End If
        Next
        If cnt = UBound(ExeTest) + 1 Then
            With pctNoFile
                .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
                .Visible = True
                .ZOrder
            End With
            ret = MsgBox("File " & Var(PhoneFileExe).Valore & " non trovato!" & vbNewLine & "Il programma PhoneFile non è stato trovato sul PC. Verrà aperta la pagina web del programma." & vbNewLine & vbNewLine & "Vuoi cercare ora il file?", vbInformation + vbYesNo, App.ProductName)
            
            ' Apro la pagine web del programma
            ShellExecute hwnd, "open", Var(PhoneFileWeb).Valore, vbNullString, vbNullString, SW_SHOW
    
            If ret = vbYes Then
                MsgBox "Funzione non ancora disponibile.", vbInformation, App.ProductName
            End If
            Exit Sub
        End If
    End If
    
    MsgBox "ATTENZIONE! " & vbNewLine & "Questa finestra è ancora in fase di test. Non utilizzarla se non sai cosa fare." & vbNewLine & "", vbExclamation, App.ProductName

    Set objPhnFile = New cPhoneFile
    objPhnFile.PhoneFile = Var(PhoneFileExe).Valore
    objPhnFile.GetPhoneFileInfo

    Set m_cSplitH = New cSplitter
    m_cSplitH.Initialise picSplitH, Me, cSPLTOrientationHorizontal, 40
    m_cSplitH.OffsetBasso = lblInfo
    
    ' Creo gli array con i dati delle colonne
    LarghezzaColonne = Array(1200, 4000, 1300, 2500, 10, 0)
    ColonneListView1 = Array("", "Nome", "Data", "File", "Estensione", "Immagine")
    ' Serve per la funzione di ordinamento dei dati nella colonna
    TagColonne = Array("NUMBER", "STRING", "DATE", "STRING", "STRING", "STRING")
    
    With ListView1(0)
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
    Call SetListViewColor(ListView1(0), Picture1, 1, vbWhite, vbGreenLemon)
    Call AutoSizeUltimaColonna(ListView1(0))
    
    strTmp = "Da questa finestra si possono trasferire sul telefono i file dei POI presenti nella cartella: "
    lblHelp.Caption = strTmp
    CartellaLavoro = Var(PoiScaricati).Valore
    lblDir.Caption = CartellaLavoro
    
    txtCerca.Text = Var(PhoneFileMapDir).Valore
    txtCerca.ForeColor = vbGrayText
        
    Call ElencoFile
        
    NessunTelefono = "Scrivi qua il nome bluetooth del telefono"
    cmbTelefoni.Text = NessunTelefono
    cmbTelefoni.ToolTipText = NessunTelefono
    MancaNomeTel = True
    Call CaricaElencoTelefoni
        
End Sub

Private Sub ElencoFile()
    ' Legge i file contenuti nella cartella..............
    Dim cnt As Long
    Dim StmpFile As String
    Dim strTmp As String

    cnt = 0
    StmpFile = Dir(CartellaLavoro & "\*.ov2")
    ListView1(0).ListItems.Clear
    FileTrovati = 0
    
    While StmpFile <> ""
        ReDim Preserve ListaFile(cnt)
        ListaFile(cnt).Nome = Left$(StmpFile, Len(StmpFile) - 4)
        ListaFile(cnt).Data = GetDataFile(CartellaLavoro & "\" & StmpFile, "M")
        ListaFile(cnt).File = CartellaLavoro & "\" & StmpFile
        ListaFile(cnt).Estensione = EstensioneFromFile(StmpFile)
        
        ' Inserisco i dati del file .bmp
        strTmp = CartellaLavoro & "\" & ListaFile(cnt).Nome & ".bmp"
        If FileExists(strTmp) = True Then
            ListaFile(cnt).Immagine = strTmp
        Else
            ListaFile(cnt).Immagine = "" ' Il file bmp non esiste
        End If
        
        StmpFile = Dir()
        cnt = cnt + 1
        FileTrovati = cnt
    Wend

    If FileTrovati > 0 Then
        lblInfo.Caption = "Inserisci il Nome del telefono, quindi premi il tasto """ & cmdDir.Caption & """"
        Call ElencoFileInListView
    Else
        lblInfo.Caption = "Non sono stati trovati file nella cartella. " & "Le sottocartelle non sono state scansionate."
        cmdTrasferisci.Enabled = False
    End If
    
End Sub

Private Sub ElencoFileInListView()
    Dim cntRecord As Long
    Dim cnt As Long
    Dim itmX As Variant

    ListView1(0).ListItems.Clear
        
    cntRecord = 0
    For cnt = 0 To UBound(ListaFile) ' Scorro tutte le righe dell'array
        Set itmX = ListView1(0).ListItems.Add(, , Format(cntRecord + 1, "00000"))
        itmX.SubItems(GetNumColDaIntestazione(ListView1(0), "Nome") - 1) = ListaFile(cnt).Nome
        itmX.SubItems(GetNumColDaIntestazione(ListView1(0), "Data") - 1) = ListaFile(cnt).Data
        itmX.SubItems(GetNumColDaIntestazione(ListView1(0), "File") - 1) = ListaFile(cnt).File
        itmX.SubItems(GetNumColDaIntestazione(ListView1(0), "Estensione") - 1) = ListaFile(cnt).Estensione
        itmX.SubItems(GetNumColDaIntestazione(ListView1(0), "Immagine") - 1) = ListaFile(cnt).Immagine
        cntRecord = cntRecord + 1
    Next
    
    TotaleRigheListView ListView1(0)
    Call ControllaCheck(ListView1(0))

    ListView1(0).Refresh

End Sub

Private Sub CercaDirMappa(ByVal DirectoryCell As String, Optional ByVal TestoElencoDir As String = "", Optional ByVal Predefinito As String = "")
    ' Guarda nella root del cellulare per cercare le cartelle delle mappe installate
    Dim arrTmpDir
    Dim arrDirMap() As String
    Dim cnt As Integer
    Dim cntDirMap As Integer
    Dim strTmp As String
    
    If TestoElencoDir = "" Then
        txtOutputs(0).Text = ""
        objPhnFile.LeggiDir DirectoryCell
        TestoElencoDir = txtOutputs(0).Text
    End If
    
    cmbDirMap.Clear
    
    ' Carico l'lenco delle cartelle nell'array
    arrTmpDir = Split(TestoElencoDir, vbCrLf, , vbTextCompare)
    ReDim arrDirMap(UBound(arrTmpDir))
    
    For cnt = 0 To UBound(arrTmpDir)
        If InStr(1, UCase(arrTmpDir(cnt)), "<DIR>", vbTextCompare) <> 0 And InStr(1, UCase(arrTmpDir(cnt)), UCase(Var(PhoneFileMapDir).Valore), vbTextCompare) <> 0 Then
            strTmp = Trim(Right$(arrTmpDir(cnt), InStr(1, UCase(arrTmpDir(cnt)), "<DIR>", vbTextCompare)))
            strTmp = Replace$(strTmp, vbCr, "", , , vbTextCompare)
            strTmp = Replace$(strTmp, vbLf, "", , , vbTextCompare)
            arrDirMap(cntDirMap) = DirectoryCell & strTmp
            cntDirMap = cntDirMap + 1
        End If
    Next
    cntDirMap = cntDirMap - 1
    
    If cntDirMap >= 0 Then
        ReDim Preserve arrDirMap(cntDirMap)
        For cnt = 0 To cntDirMap
            cmbDirMap.AddItem (arrDirMap(cnt))
        Next
    End If

    If cmbDirMap.ListCount >= 1 Then
        cmbDirMap.ListIndex = 0
    Else
        MsgBox "Nessuna cartella trovata con il riferimento: " & Chr(34) & txtCerca & Chr(34), vbInformation, App.ProductName
    End If
    
    If Predefinito <> "" Then
        CercaInComboBox cmbDirMap, Predefinito
    End If

End Sub

Private Sub CaricaElencoTelefoni()
    Dim arrTmp
    Dim cnt As Integer
    
    cmbTelefoni.Clear
    arrTmp = Split(Var(PhoneFileTele).Valore, "|", , vbTextCompare)
    
    For cnt = 0 To UBound(arrTmp)
        If arrTmp(cnt) <> "" Then cmbTelefoni.AddItem (arrTmp(cnt))
    Next
    
    If cmbTelefoni.ListCount >= 1 Then
        cmbTelefoni.ListIndex = 0
        MancaNomeTel = False
    Else
        cmbTelefoni.Text = NessunTelefono
        MancaNomeTel = True
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   m_cSplitH.MouseMove X, Y
   ElaboraLabelLink
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   m_cSplitH.MouseUp X, Y
End Sub

Private Sub Form_Resize()
   ' Width = Larghezza  Height = Altezza
    On Error Resume Next

    If Me.Width < 10500 Then Me.Width = 10500
    If Me.Height < 6200 Then Me.Height = 6200

    With pctNoFile
        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End With

    With lblHelp
        .Move 0, .Top, Me.ScaleWidth, .Height
    End With
    
    With lblDir
        .Move 0, .Top, Me.ScaleWidth, .Height
    End With
    
    With picComandi
        .Move (Me.ScaleWidth - .Width) / 2, Me.ScaleHeight - .Height, .Width, .Height
    End With
    
    With lblInfo
        .Move 0, picComandi.Top - .Height, Me.ScaleWidth, .Height
    End With
    
    m_cSplitH.Resize ListView1, txtOutputs

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objPhnFile = Nothing
    Set m_cSplitH = Nothing
    
    frmMain.Visible = True
    frmMain.WindowState = vbNormal
    frmMain.ZOrder
    frmMain.SetFocus

End Sub

Private Sub lblDir_Click()
    Dim strTmp As String
    
    strTmp = BrosweForFolder(frmMain, "Seleziona la cartella contenente i file.")
    
    If strTmp <> "" Then
        CartellaLavoro = strTmp
        lblDir.Caption = CartellaLavoro
        Call ElencoFile
    End If
    
End Sub

Private Sub lblDir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call ElaboraLabelLink(lblDir)
    lblDir.ToolTipText = "Click per selezionare un'altra cartella."

End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ElaboraLabelLink
End Sub

Private Sub ListView1_ItemCheck(index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim ret
    
    ret = GetNumeroRigheChecked(ListView1(0))
    
    If ret >= 1 And cmbDirMap.ListCount >= 1 Then
        cmdTrasferisci.Enabled = True
    Else
        cmdTrasferisci.Enabled = False
    End If

    If cmbDirMap.Text = "" Then cmdTrasferisci.Enabled = False
    
    ListView1(0).SetFocus
    
End Sub

Private Function NomeTelefono() As String

    NomeTelefono = cmbTelefoni.Text
    ' Se il nome contiene degli spazi aggiungo "
    If InStr(1, NomeTelefono, " ", vbTextCompare) > 0 Then
        NomeTelefono = Chr(34) & NomeTelefono & Chr(34)
    End If
    
End Function

Private Sub objDOS_ReceiveOutputs(CommandOutputs As String)
    txtOutputs(0).Text = txtOutputs(0).Text & CommandOutputs
End Sub

Private Sub ListView1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ElaboraLabelLink
End Sub

Private Sub objPhnFile_RispostaPhoneFile(CommandOutputs As String)
    txtOutputs(0).Text = txtOutputs(0).Text & CommandOutputs
End Sub

Private Sub txtCerca_DblClick()
    txtCerca.Locked = False
    txtCerca.ForeColor = vbInfoText
End Sub

Private Sub txtCerca_LostFocus()
    If txtCerca.Locked = False Then lVar(PhoneFileMapDir) = txtCerca.Text
    txtCerca.Locked = True
    txtCerca.ForeColor = vbGrayText
End Sub

Private Sub txtOutputs_Change(index As Integer)
    ' Imposto la posizione di scrittura dei nuovi caratteri
    txtOutputs(index).SelStart = Len(txtOutputs(index).Text)
End Sub

Private Sub ElaboraLabelLink(Optional LabelLink As Object)
    On Error Resume Next

    Dim ctl As Control
    
    Screen.MousePointer = vbDefault
    
    For Each ctl In Controls
        If TypeOf ctl Is Label Then
            If ctl.Name <> LabelLink.Name Then
                ctl.ForeColor = vbBlack
            Else
                ctl.ForeColor = vbBlue
            End If
            
        End If
    Next
    
    DoEvents

End Sub

' Inizio funzioni per lo split -----------------------------------------------------------------------------------------

Private Sub picSplitH_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   m_cSplitH.MouseDown index, X, Y
End Sub

Private Sub m_cSplitH_DoSplit(bSplit As Boolean)
   ' Can cancell split here
End Sub

Private Sub m_cSplitH_SplitComplete()
   Form_Resize
End Sub

