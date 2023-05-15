VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSetupDescrizione 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Descrizione"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   Icon            =   "frmSetupDescrizione.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRipSetup 
      Caption         =   "&Ripristina setup predefinto"
      Height          =   495
      Left            =   8760
      TabIndex        =   70
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtSetupDescrizione 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8760
      TabIndex        =   69
      Text            =   "txtSetupDescrizione"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CheckBox ckDescrizione 
      Caption         =   "Categoria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   120
      TabIndex        =   67
      Top             =   2280
      Width           =   1400
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   15
      Left            =   7080
      TabIndex        =   66
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   14
      Left            =   1560
      TabIndex        =   65
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox ckIncludiDelim 
      Appearance      =   0  'Flat
      Caption         =   "Includi sempre i delimitatori di campo"
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   8760
      TabIndex        =   58
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtTelInt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   56
      Text            =   "39"
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox ckTelInt 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1200
      TabIndex        =   21
      Top             =   1920
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "&Chiudi  [Esc]"
      Height          =   495
      Left            =   8760
      TabIndex        =   55
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdConcatena 
      Height          =   315
      Index           =   1
      Left            =   3270
      TabIndex        =   14
      Top             =   1215
      Width           =   315
   End
   Begin VB.CommandButton cmdConcatena 
      Height          =   315
      Index           =   0
      Left            =   3270
      TabIndex        =   5
      Top             =   855
      Width           =   315
   End
   Begin VB.CheckBox ckDescrizione 
      Caption         =   "Città"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   720
      TabIndex        =   11
      Top             =   1275
      Width           =   675
   End
   Begin VB.CommandButton cmdMuoviList 
      Caption         =   " Muovi &giù"
      Height          =   615
      Index           =   1
      Left            =   9720
      TabIndex        =   29
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdMuoviList 
      Caption         =   "Muovi  &su"
      Height          =   615
      Index           =   0
      Left            =   8880
      TabIndex        =   28
      Top             =   1920
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   5640
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.ListBox lstOrdine 
      Height          =   2205
      Left            =   8760
      TabIndex        =   53
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   13
      Left            =   3000
      TabIndex        =   13
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   12
      Left            =   1560
      TabIndex        =   12
      Text            =   "("
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   11
      Left            =   3000
      TabIndex        =   4
      Text            =   "]"
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   10
      Left            =   1560
      TabIndex        =   3
      Text            =   "["
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   9
      Left            =   7080
      TabIndex        =   23
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   1560
      TabIndex        =   22
      Text            =   ">"
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   7080
      TabIndex        =   19
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   1560
      TabIndex        =   18
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   7080
      TabIndex        =   16
      Text            =   ")"
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   3600
      TabIndex        =   15
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   7080
      TabIndex        =   9
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   3600
      TabIndex        =   7
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   7080
      TabIndex        =   2
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtAggiunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox ckDescrizione 
      Appearance      =   0  'Flat
      Caption         =   "Cap"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1275
      Width           =   735
   End
   Begin VB.ComboBox cmbTaglia 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmSetupDescrizione.frx":030A
      Left            =   7440
      List            =   "frmSetupDescrizione.frx":030C
      TabIndex        =   27
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox ckAvanzaAutomatico 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5520
      TabIndex        =   49
      Top             =   5310
      Width           =   255
   End
   Begin VB.CommandButton cmdCreaDescrizioni 
      Caption         =   "Salva &corrente"
      Height          =   525
      Index           =   1
      Left            =   5400
      TabIndex        =   48
      ToolTipText     =   "ATTENZIONE! Premendo questo pulsante verranno eseguite operazioni sui dati."
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreaDescrizioni 
      BackColor       =   &H8000000D&
      Caption         =   "Salva &tutti"
      Height          =   525
      Index           =   0
      Left            =   6960
      TabIndex        =   47
      ToolTipText     =   "ATTENZIONE! Premendo questo pulsante verranno eseguite operazioni sui dati."
      Top             =   5040
      Width           =   1455
   End
   Begin VB.PictureBox pctCarMax 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   760
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   1665
      TabIndex        =   44
      Top             =   4800
      Width           =   1695
      Begin VB.TextBox txtCarMax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -120
         TabIndex        =   45
         Text            =   "63"
         ToolTipText     =   "Indicare quanti caratteri massimi può avere la Descrizione finale"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblCarMax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Caratteri Max:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -120
         TabIndex        =   46
         Top             =   0
         Width           =   1875
      End
   End
   Begin VB.PictureBox pctElabora 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   8265
      TabIndex        =   37
      Top             =   3720
      Width           =   8295
      Begin VB.VScrollBar VScroll1 
         Height          =   330
         Left            =   5880
         Max             =   10
         Min             =   -10
         TabIndex        =   59
         Top             =   480
         Width           =   200
      End
      Begin VB.TextBox txtSostCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "0"
         ToolTipText     =   "(0: sostituisce tutto)  (x sostituisce x volte partendo dall'inizio)  (-x sostituisce x volte partendo dalla fine)"
         Top             =   480
         Width           =   470
      End
      Begin VB.CommandButton cmdAvviaOperazione 
         Caption         =   "&Avvia Operazione"
         Height          =   700
         Left            =   6120
         TabIndex        =   43
         ToolTipText     =   "ATTENZIONE! Premendo questo pulsante verranno eseguite operazioni sui dati."
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox txtSostituisci 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         TabIndex        =   42
         Text            =   "Sostituisci"
         ToolTipText     =   "Scrivi qua il testo che verrà sostituito con il testo nella casella Cerca"
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtCerca 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         TabIndex        =   41
         Text            =   "Cerca"
         ToolTipText     =   "Scrivi qua il testo da cercare"
         Top             =   120
         Width           =   3615
      End
      Begin VB.ComboBox cmbDove 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSetupDescrizione.frx":030E
         Left            =   120
         List            =   "frmSetupDescrizione.frx":0310
         TabIndex        =   40
         Text            =   "cmbDove"
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox cmbOperazione 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSetupDescrizione.frx":0312
         Left            =   120
         List            =   "frmSetupDescrizione.frx":0314
         TabIndex        =   38
         Text            =   "cmbOperazione"
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox pctMuovi 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   760
      Left            =   2040
      ScaleHeight     =   735
      ScaleWidth      =   3105
      TabIndex        =   35
      Top             =   4800
      Width           =   3135
      Begin VB.CommandButton cmdMuovi 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   120
         TabIndex        =   63
         ToolTipText     =   "Ir a la Primera Página"
         Top             =   120
         Width           =   465
      End
      Begin VB.CommandButton cmdMuovi 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   600
         TabIndex        =   62
         ToolTipText     =   "Ir a la Página Anterior"
         Top             =   120
         Width           =   345
      End
      Begin VB.CommandButton cmdMuovi 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   2
         Left            =   2160
         TabIndex        =   61
         ToolTipText     =   "Ir a la Siguiente Página"
         Top             =   120
         Width           =   345
      End
      Begin VB.CommandButton cmdMuovi 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         Left            =   2520
         TabIndex        =   60
         ToolTipText     =   "Ir a la Ultima Página"
         Top             =   120
         Width           =   465
      End
      Begin VB.Label lblRecord 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   960
         TabIndex        =   36
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.TextBox txtLen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   7440
      TabIndex        =   26
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtLen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   7440
      TabIndex        =   25
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtLen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   7440
      TabIndex        =   24
      ToolTipText     =   "Indicare quanti caratteri utilizzare per questo campo. Se lasciato vuoto, la funzione è automatica."
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox ckDescrizione 
      Caption         =   "Descrizione"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Value           =   1  'Checked
      Width           =   1400
   End
   Begin VB.CheckBox ckDescrizione 
      Caption         =   "Indirizzo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1400
   End
   Begin VB.CheckBox ckDescrizione 
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1400
   End
   Begin VB.CheckBox ckDescrizione 
      Caption         =   "Telefono"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.Label lblStruttura 
      Alignment       =   2  'Center
      Caption         =   "lblStruttura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   68
      Top             =   30
      Width           =   10695
   End
   Begin VB.Label lblDescrizione 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " lblDes(7)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   7
      Left            =   1800
      TabIndex        =   64
      Top             =   2280
      Width           =   5295
   End
   Begin VB.Label lblDescrizione 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " lblDes(6)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   6
      Left            =   1800
      TabIndex        =   52
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblDescrizione 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " lblDes(5)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   5
      Left            =   1800
      TabIndex        =   51
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblEsempio 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblEsempio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   1
      Left            =   120
      TabIndex        =   50
      ToolTipText     =   $"frmSetupDescrizione.frx":0316
      Top             =   2880
      Width           =   8295
   End
   Begin VB.Label lblCambiaRigaCorrente 
      BackColor       =   &H0080FFFF&
      Caption         =   "lblCambiaRigaCorrente"
      Height          =   255
      Left            =   7440
      TabIndex        =   39
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblEsempio 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblEsempio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   2625
      Width           =   8295
   End
   Begin VB.Label lblDescrizione 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " lblDes(4)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   4
      Left            =   2160
      TabIndex        =   33
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Label lblDescrizione 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " lblDes(3)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   3
      Left            =   1800
      TabIndex        =   32
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label lblDescrizione 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " lblDes(2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   2
      Left            =   3840
      TabIndex        =   31
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblDescrizione 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " lblDes(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   3840
      TabIndex        =   8
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblDescrizione 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " lblDes(0)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   1800
      TabIndex        =   30
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frmSetupDescrizione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
' Utilizzati nella funzione visualizza
Dim Des As String, Indi As String, NumCiv As String
Dim Cap As String, Citta As String, Prov As String
Dim Tel As String, Cat As String

Private Type txtOv2
    iOrdine As Long
    iLabelIndex As Long
    iTxtPrimaIndex As Long
    iTxtDopoIndex As Long
    iPrima As String
    iTesto As String
    iDopo As String
    iCompleto As String
End Type

Private arrAggiunta() As txtOv2
'-------------------------------------------------------------------------------

Private Enum eCampi
    eDescrizione = 0
    eIndirizzo = 1
    eCittà = 2
    eProvincia = 3
    eTelefono = 4
    eNumeroCivico = 5
    eCap = 6
    eCat = 7
End Enum

' Per il concatena della ListBox
' aListIndex contiene:
' (0) Descrizione
' (1) Primo
' (2) Secondo
Dim aListIndex() As String

Dim Campi As eCampi
Dim Concatena(1) As Long

Dim AltezzaForm As Long
Dim AltezzaFormBar As Long
Dim RigaCorrente As Long
Dim FormCaricata As Boolean

Private LBHS As clsListBox

Private Sub ckAvanzaAutomatico_Click()

    If ckAvanzaAutomatico.value = 0 Then cmdCreaDescrizioni(1).Caption = "Salva &Corrente" & vbNewLine & " "
    If ckAvanzaAutomatico.value = 1 Then cmdCreaDescrizioni(1).Caption = "Salva &Corrente" & vbNewLine & " ed avanza"
    
End Sub

Private Sub ckIncludiDelim_Click()
    Call Visualizza(RigaCorrente)
End Sub

Private Sub ckTelInt_Click()
    Call Visualizza(RigaCorrente)
End Sub

Private Sub cmbTaglia_Click()
    If FormCaricata = True Then
        Call Visualizza(RigaCorrente)
    End If
End Sub

Private Sub cmdCreaDescrizioni_Click(index As Integer)
    Dim cnt As Long
    
    Select Case index
        Case Is = 0
            If MsgBoxx(Me.hwnd, "ATTENZIONE! Verranno modificati e salvati i dati delle descrizione di tutti i record." & vbNewLine & "Continuo l'operazione?") = vbYes Then
                Call VuotaLabel
                Call FormBarResize(GetNumeroRighe(frmWeb.ListView1), True)
                For cnt = 1 To GetNumeroRighe(frmWeb.ListView1)
                    RigaCorrente = cnt
                    Call Visualizza(RigaCorrente, True, True, False)
                    ProgressBar1.value = cnt
                Next
                Call FormBarResize(, False)
                RigaCorrente = 1
                Call Visualizza(RigaCorrente, False, True, True)
            End If
            
        Case Is = 1
            Call Visualizza(RigaCorrente, True, True)
            If ckAvanzaAutomatico = 1 Then
                If RigaCorrente >= GetNumeroRighe(frmWeb.ListView1) Then
                    cmdMuovi_Click (0)
                Else
                    cmdMuovi_Click (2)
                End If
            End If
    End Select
    
    Call AutoSizeUltimaColonna(frmWeb.ListView1)
    
End Sub

Private Sub cmdCreaDescrizioni_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ckAvanzaAutomatico.Visible = False
End Sub
Private Sub cmdCreaDescrizioni_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ckAvanzaAutomatico.Visible = True
End Sub

Private Sub VuotaLabel()
    Dim cnt As Long
    
    For cnt = 0 To lblDescrizione.count - 1
        lblDescrizione(cnt).Caption = ""
        lblDescrizione(cnt).ToolTipText = ""
    Next
   
     For cnt = 0 To lblEsempio.count - 1
        lblEsempio(cnt).Caption = ""
        lblEsempio(cnt).ToolTipText = ""
    Next
  
    lblRecord.Caption = ""
    DoEvents
End Sub

Private Sub FormBarResize(Optional ProgressBarMax As Long = 100, Optional BarVisibile As Boolean = True)
    If BarVisibile = True Then
        MousePointer = vbHourglass
        ProgressBar1.value = 0
        ProgressBar1.Max = ProgressBarMax
        frmSetupDescrizione.Height = AltezzaFormBar
    Else
        frmSetupDescrizione.Height = AltezzaForm
        MousePointer = vbDefault
    End If
    DoEvents
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdMuoviList_Click(index As Integer)
    Dim strTemp1 As String  'Hold the selected index data temporarily for move
    Dim iCnt As Integer     'Holds the index of the item to be moved
    
    Select Case index
        Case Is = 0 ' Muovi Su
            iCnt = lstOrdine.ListIndex
            If iCnt = -1 Then iCnt = lstOrdine.ListCount - 1
            If iCnt > 0 Then
                strTemp1 = lstOrdine.List(iCnt)
                ' Add the item selected to one position above the current position
                lstOrdine.AddItem strTemp1, (iCnt - 1)
                ' Remove it from the current position. Note the current position has changed because the add has moved everything down by 1
                lstOrdine.RemoveItem (iCnt + 1)
                ' Reselect the item that was moved.
                lstOrdine.Selected(iCnt - 1) = True
            End If
        Case Is = 1 ' Muovi Giù
            iCnt = lstOrdine.ListIndex 'Assign the first index
            If iCnt = -1 Then iCnt = 0
            If iCnt > -1 And iCnt < lstOrdine.ListCount - 1 Then
                strTemp1 = lstOrdine.List(iCnt)
                ' Add the item selected to below the current position
                lstOrdine.AddItem strTemp1, (iCnt + 2)
                lstOrdine.RemoveItem (iCnt)
                ' Reselect the item that was moved.
                lstOrdine.Selected(iCnt + 1) = True
            End If
    End Select
    
    Call Visualizza(RigaCorrente, False, True, True)

End Sub

Private Sub cmdRipSetup_Click()
    ' Reimposto i valori di defualt
    PreparaSetupDescrizione 1, Var(SetupDescPred).Valore
End Sub

Private Sub Form_Activate()
    frmMain.Visible = False
End Sub

Private Sub Form_Initialize()
    FormCaricata = False
End Sub

Private Sub Form_Load()
    
    FormCaricata = False
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    RigaCorrente = 1
    
    AltezzaFormBar = Me.Height
    AltezzaForm = Me.Height - ProgressBar1.Height - 50
    ProgressBar1.Top = Me.Height - ProgressBar1.Height - 480
    ProgressBar1.Left = -10
    ProgressBar1.Width = Me.Width
    Me.Height = AltezzaForm
    
    cmbOperazione.AddItem (" Cerca e sostituisci")
    Call BloccaComboBox(cmbOperazione)
    ' Seleziono il primo elemento
    cmbOperazione.ListIndex = 0
    
    cmbDove.AddItem ("   Descrizione")
    cmbDove.AddItem ("   Indirizzo")
    cmbDove.AddItem ("   Cap")
    cmbDove.AddItem ("   Città")
    cmbDove.AddItem ("   Pr")
    cmbDove.AddItem ("   Telefono")
    cmbDove.AddItem ("   Categoria")
    cmbDove.AddItem ("   Desc.ov2")
    cmbDove.AddItem ("   Tutte le colonne")
    Call BloccaComboBox(cmbDove)
    ' Seleziono il primo elemento
    cmbDove.ListIndex = 0
    cmdAvviaOperazione.Caption = "&Avvia Operazione" & vbNewLine & ("su tutte le righe")
    
    cmbTaglia.AddItem (ckDescrizione(0).Caption)
    cmbTaglia.ListIndex = 0
    cmbTaglia.ToolTipText = "Selezionare il campo dal quale verranno eliminati dei dati se la descrizione supera il numero di caratteri massimi."
    Call BloccaComboBox(cmbTaglia)
    
    ckTelInt.ToolTipText = "Aggiunge il prefisso " & txtTelInt.Text & " al numero di telefono."
    
    cmdRipSetup.ToolTipText = Var(SetupDescPred).Valore
    
    ' Inserisco i primi valori nella ListBox
    Call ListBoxOrdine(-1)

    ckAvanzaAutomatico.value = 1
    txtSostCount.Text = VScroll1.value
    lblStruttura.Caption = "Descrizione consigliata: ""Descrizione[NumeroCivico]Indirizzo(CAP Città)Provincia>Telefono"""
    
    ' Carico le impostazioni dal file .rmk oppure dalla variabile SetupDescrizione
    If SetupDescrizione = "" Then
        PreparaSetupDescrizione (1)
    Else
        PreparaSetupDescrizione (1), SetupDescrizione
    End If
    
    FormCaricata = True
    
    Exit Sub
    
Errore:
    FormCaricata = True
    GestErr Err, "Errore nella funzione frmSetupDescrizione.Form_Load."

End Sub

Private Sub Form_Resize()
    VScroll1.Move txtSostCount.Left + txtSostCount.Width, txtSostCount.Top, VScroll1.Width, txtSostCount.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' Scrivo le impostazioni della variabile SetupDescrizione
    PreparaSetupDescrizione (0)
    frmWeb.SetFocus
End Sub

Private Sub lblCambiaRigaCorrente_Change()
    
    If lblCambiaRigaCorrente.Caption <> "" Then
        RigaCorrente = lblCambiaRigaCorrente.Caption
        Call Visualizza(RigaCorrente)
    Else
        Call VuotaLabel
    End If
    
End Sub

Private Sub lblCarMax_Click()
    txtCarMax.SetFocus
End Sub

Private Sub txtAggiunta_Change(index As Integer)
    Call Visualizza(RigaCorrente)
End Sub

Private Sub txtAggiunta_KeyPress(index As Integer, KeyAscii As Integer)
      KeyAscii = KeyValidate(KeyAscii, AlphaNumeric, txtAggiunta(index), 6)
End Sub

Private Sub txtLen_KeyPress(index As Integer, KeyAscii As Integer)
      KeyAscii = KeyValidate(KeyAscii, Numeric, txtLen(index), 2, , True)
End Sub

Private Sub txtLen_Change(index As Integer)
    Call Visualizza(RigaCorrente)
End Sub

Private Sub txtCarMax_KeyPress(KeyAscii As Integer)
      KeyAscii = KeyValidate(KeyAscii, Numeric, txtCarMax, 3, , True)
End Sub

Private Sub txtCarMax_Change()
    If txtCarMax.Text <> "" Then
        Call Visualizza(RigaCorrente)
    End If
End Sub

Private Sub txtCarMax_GotFocus()
    txtCarMax.BackColor = &H80000005
    lblCarMax.BackColor = &H80000005
End Sub

Private Sub txtCarMax_LostFocus()
    txtCarMax.BackColor = &H8000000F
    lblCarMax.BackColor = &H8000000F
End Sub

Private Sub ckDescrizione_Click(index As Integer)

    ' Pulisco il ComboBox
    cmbTaglia.Clear

    If ckDescrizione(0).value = 1 Then
        cmbTaglia.AddItem (ckDescrizione(0).Caption)
    End If
    If ckDescrizione(1).value = 1 Then
        cmbTaglia.AddItem (ckDescrizione(1).Caption)
    End If
    If ckDescrizione(2).value = 1 Then
        cmbTaglia.AddItem (ckDescrizione(2).Caption)
    End If
    If ckDescrizione(3).value = 1 Then
        'cmbTaglia.AddItem (ckDescrizione(3).Caption)
    End If
    If ckDescrizione(4).value = 1 Then
        'cmbTaglia.AddItem (ckDescrizione(4).Caption)
    End If
    If ckDescrizione(5).value = 1 Then
        'cmbTaglia.AddItem (ckDescrizione(4).Caption)
    End If

    DoEvents
    
    ' Seleziono il primo elemento
    If cmbTaglia.ListCount > 0 Then cmbTaglia.ListIndex = 0
       
    DoEvents
    
    Call Visualizza(RigaCorrente)
    
End Sub

Private Sub cmdAvviaOperazione_Click()
    Dim TestoCerca As String
    Dim TestoSostituisci As String
    Dim Colonna As Long
    Dim ret As Integer
    
    cmdAvviaOperazione.BackColor = vbBlack
    TestoSostituisci = txtSostituisci.Text
    TestoCerca = txtCerca.Text
    
    Select Case LCase(Trim(cmbOperazione.Text))
        Case Is = LCase("Cerca e sostituisci")
            ' Controllo che sia stata inserita la parola di ricerca
            If TestoCerca <> "" Then
                If LCase(Trim(cmbDove.Text)) = LCase("Tutte le colonne") Then
                    ' Tutte le colonne
                    Colonna = -1
                Else
                    ' Solo la colonna scelta
                    Colonna = GetNumColDaIntestazione(frmWeb.ListView1, Trim(cmbDove.Text))
                End If
                ret = CercaSostituisci(frmWeb.ListView1, TestoCerca, TestoSostituisci, , Colonna, , , txtSostCount.Text)
            Else
                MsgBox "Inserisci il testo da cercare!"
                txtCerca.SetFocus
            End If
            Call Visualizza(lblRecord.Caption)
        Case Else
    End Select

End Sub

Private Sub ButOpDes()
    Dim BOD As String
    Dim ODE(20) As String
    Dim Colonna As Integer
    Dim ODES(8) As String
    Dim OIA() As String

'If Trim(Cells(2, 13).Value) <> "" And Trim(Cells(LastRow, 13).Value) <> "" Then 'And Trim(Cells(2, 14).Value) <> "" Trim(Cells(LastRow, 14).Value) <> "" Then
'    If OpDesEli.Value = True Then
    
        ODE(1) = "(SNC)"
        ODE(2) = "SPA"
        
        ODE(3) = "srl"
        ODE(4) = "RISTORANTE "
        ODE(5) = "PIZZERIA "
        
        ODE(6) = "PIAZZA "
        ODE(7) = "HOTEL "
        ODE(8) = "ALBERGO "
        
        ODE(9) = "PIAZZALE "
        ODE(10) = "& c."
        ODE(11) = "PENSIONE "
        
        ODE(12) = "CAMPEGGIO "
        ODE(13) = "OSTELLO "
        ODE(14) = "AGRITURISMO "
        
        ODE(15) = "GELATERIA "
        ODE(16) = "& C."
        ODE(17) = "(S.N.C.)"
        
        ODE(18) = "S.N.C."
        ODE(19) = "S.P.A."
        ODE(20) = "& C"
    
        'CurrRow = 2
'        Carattere = ""
        'While Trim(Cells(CurrRow, 1).Value) <> ""
        '    BOD = Trim(Cells(CurrRow, 1).Value)
        '    For i = 1 To 20
        '        BOD = Replace(BOD, ODE(i), Carattere, , , vbTextCompare)
        '    Next
        '    Cells(CurrRow, 1).Value = BOD
        '    CurrRow = CurrRow + 1
        '
        'Wend
        
    'ElseIf OpDesAbb.Value = True Then
      
      'While Trim(Cells(CurrRow, 1).Value) <> ""
      '  BOD = Trim(Cells(CurrRow, 1).Value)
      '
            BOD = Replace(BOD, "PIZZERIA", "Pizz ", , , vbTextCompare)
            BOD = Replace(BOD, "RISTORANTE", "Rist ", , , vbTextCompare)
            BOD = Replace(BOD, "HOTEL", "Htl ", , , vbTextCompare)
            BOD = Replace(BOD, "ALBERGO", "Alb ", , , vbTextCompare)
            BOD = Replace(BOD, "DISTRIBUTORE", "Dstrb ", , , vbTextCompare)
            BOD = Replace(BOD, "AGRITURISMO", "Agrit ", , , vbTextCompare)
            BOD = Replace(BOD, "PENSIONE", "Pens ", , , vbTextCompare)
            BOD = Replace(BOD, "SNC", "", , , vbTextCompare)
            BOD = Replace(BOD, "SPA", "", , , vbTextCompare)
            BOD = Replace(BOD, "SRL", "", , , vbTextCompare)
            
       '     Cells(CurrRow, 1) = BOD
       '     CurrRow = CurrRow + 1
       ' Wend
'    End If
'Else: TextBoxControllo.Text = " Effettuare prima il Backup...."
'End If
'OpDesDef.Value = True

End Sub

Private Sub cmdMuovi_Click(index As Integer)
    
    Select Case index
        Case Is = 0
            RigaCorrente = 1
        Case Is = 1
            If RigaCorrente <= 1 Then
                RigaCorrente = 1
            Else
                RigaCorrente = RigaCorrente - 1
            End If
        Case Is = 2
            If RigaCorrente >= GetNumeroRighe(frmWeb.ListView1) Then
                RigaCorrente = GetNumeroRighe(frmWeb.ListView1)
            Else
                RigaCorrente = RigaCorrente + 1
            End If
        Case Is = 3
            RigaCorrente = GetNumeroRighe(frmWeb.ListView1)
    End Select
    
    Call Visualizza(RigaCorrente, False, True)
    
End Sub

Private Sub cmdConcatena_Click(index As Integer)
    Static Val(1) As Long
    
    If Val(index) = 0 Then
        Val(index) = 1
    ElseIf Val(index) = 1 Then
        Val(index) = 2
    ElseIf Val(index) = 2 Then
        Val(index) = 3
    ElseIf Val(index) = 3 Then
        Val(index) = 1
    End If
    Concatena(index) = Val(index)
    Call Visualizza(RigaCorrente)
    
End Sub

Private Sub ControllaConcatena(Optional AggiornaLabel As Boolean = True)
    Dim cnt As Long
    Dim cnt1 As Long
    
    If AggiornaLabel = True Then
        ' Pressione tasto 0
        Select Case Concatena(0)
            Case 0 ' (NumeroCivico <> Indirizzo) Condizione all'avvio
                txtAggiunta(11).Visible = True
                txtAggiunta(2).Visible = True
                txtAggiunta(10).Visible = True
                txtAggiunta(3).Visible = True
                Call ListBoxOrdine(1, 0)
            Case 1 '(NumeroCivico <> Indirizzo)
                txtAggiunta(11).Visible = True
                txtAggiunta(2).Visible = True
                txtAggiunta(10).Visible = True
                txtAggiunta(3).Visible = True
                Call ListBoxOrdine(1, 0)
            Case 2 '(NumeroCivico & Indirizzo)
                txtAggiunta(11).Visible = False
                txtAggiunta(2).Visible = False
                txtAggiunta(10).Visible = True
                txtAggiunta(3).Visible = True
                Call ListBoxOrdine(2, 0)
            Case 3 '(Indirizzo & NumeroCivico)
                txtAggiunta(11).Visible = True
                txtAggiunta(2).Visible = True
                txtAggiunta(10).Visible = False
                txtAggiunta(3).Visible = False
                Call ListBoxOrdine(3, 0)
        End Select
    
        ' Pressione tasto 1
        Select Case Concatena(1)
            Case 0 ' (Cap & Città) Condizione all'avvio
                txtAggiunta(13).Visible = False
                txtAggiunta(4).Visible = False
                txtAggiunta(12).Visible = True
                txtAggiunta(5).Visible = True
                Call ListBoxOrdine(2, 1)
            Case 1 ' (Cap <> Città)
                txtAggiunta(13).Visible = True
                txtAggiunta(4).Visible = True
                txtAggiunta(12).Visible = True
                txtAggiunta(5).Visible = True
                Call ListBoxOrdine(1, 1)
            Case 2 ' (Cap & Città)
                txtAggiunta(13).Visible = False
                txtAggiunta(4).Visible = False
                txtAggiunta(12).Visible = True
                txtAggiunta(5).Visible = True
                Call ListBoxOrdine(2, 1)
            Case 3 ' (Città & Cap)
                txtAggiunta(13).Visible = True
                txtAggiunta(4).Visible = True
                txtAggiunta(12).Visible = False
                txtAggiunta(5).Visible = False
                Call ListBoxOrdine(3, 1)
        End Select
    End If
    
    cnt1 = 0
    ReDim arrAggiunta(lblDescrizione.UBound)
    For cnt = 0 To (lblDescrizione.UBound)
        arrAggiunta(cnt).iPrima = txtAggiunta(cnt1).Text
        arrAggiunta(cnt).iTxtPrimaIndex = txtAggiunta(cnt1).index
        
        arrAggiunta(cnt).iDopo = txtAggiunta(cnt1 + 1).Text
        arrAggiunta(cnt).iTxtDopoIndex = txtAggiunta(cnt1 + 1).index
        
        arrAggiunta(cnt).iLabelIndex = lblDescrizione.Item(cnt).index
        cnt1 = cnt1 + 2
    Next

End Sub

Private Sub ListBoxOrdine(Azione As Integer, Optional Tasto As Integer = -1)
    Dim cnt As Long
    Dim ret As Long
    '                                           Il numero è l'indice della label lblDescrizione
    Const sDes As String = " Descrizione                         |0|0"
    Const sNum As String = " NumeroCivico                        |5|5"
    Const sInd As String = " Indirizzo                           |1|1"
    Const sCap As String = " Cap                                 |6|6"
    Const sCit As String = " Città                               |2|2"
    Const sPro As String = " Provincia                           |3|3"
    Const sTel As String = " Telefono                            |4|4"
    Const sCat As String = " Categoria                           |7|7"
    
 Const mCapCit As String = " Cap & Città                         |6|2"
 Const mCitCap As String = " Città & Cap                         |2|6"
 Const mNumInd As String = " NumeroCivico & Indirizzo            |5|1"
 Const mIndNum As String = " Indirizzo & NumeroCivico            |1|5"
    
    Set LBHS = New clsListBox
    LBHS.Attach lstOrdine
    
    If Azione = -1 Then  ' Prepara la ListBoxVuota
        With lstOrdine
            .Clear
            .AddItem (sDes)
            .AddItem (sNum)
            .AddItem (sInd)
            .AddItem (mCapCit)
            .AddItem (sPro)
            .AddItem (sCat)
            .AddItem (sTel)
        End With
    End If
    
    If Tasto = 0 Then '-------------------------------------------------------
        Select Case Azione
            Case 1 ' (NumeroCivico <> Indirizzo)---------------
                ret = LBHS.CercaStringaCompleta(mNumInd)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem sInd, ret
                    lstOrdine.AddItem sNum, ret
                End If
                ret = LBHS.CercaStringaCompleta(mIndNum)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem sInd, ret
                    lstOrdine.AddItem sNum, ret
                End If
            Case 2 ' (NumeroCivico & Indirizzo)---------------
                ret = LBHS.CercaStringaCompleta(mIndNum)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mNumInd, ret
                End If
                ret = LBHS.CercaStringaCompleta(sNum)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mNumInd, ret
                End If
                ret = LBHS.CercaStringaCompleta(sInd)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mNumInd, ret
                End If
            Case 3 '(Indirizzo & NumeroCivico)-----------------
                ret = LBHS.CercaStringaCompleta(mNumInd)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mIndNum, ret
                End If
                ret = LBHS.CercaStringaCompleta(sNum)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mIndNum, ret
                End If
                ret = LBHS.CercaStringaCompleta(sInd)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mIndNum, ret
                End If
        End Select
        
    ElseIf Tasto = 1 Then '-------------------------------------------------------
        Select Case Azione
            Case 1 ' (Cap <> Città)-------------------------------------
                ret = LBHS.CercaStringaCompleta(mCapCit)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem sCit, ret
                    lstOrdine.AddItem sCap, ret
                End If
                ret = LBHS.CercaStringaCompleta(mCitCap)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem sCit, ret
                    lstOrdine.AddItem sCap, ret
                End If
            Case 2 ' (Cap & Città)--------------------------------------
                ret = LBHS.CercaStringaCompleta(mCitCap)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mCapCit, ret
                End If
                ret = LBHS.CercaStringaCompleta(sCap)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mCapCit, ret
                End If
                ret = LBHS.CercaStringaCompleta(sCit)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mCapCit, ret
                End If
            Case 3 ' (Città & Cap)--------------------------------------
                ret = LBHS.CercaStringaCompleta(mCapCit)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mCitCap, ret
                End If
                ret = LBHS.CercaStringaCompleta(sCap)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mCitCap, ret
                End If
                ret = LBHS.CercaStringaCompleta(sCit)
                If ret <> -1 Then
                    LBHS.RemoveItem CInt(ret)
                    lstOrdine.AddItem mCitCap, ret
                End If
        End Select
    End If

    If Azione <> -1 Then LBHS.KillDuplicati

End Sub

Private Sub ElaboraControlli()
    Dim cnt As Long
    
    Select Case RigaCorrente
        Case Is = 1
            cmdMuovi(0).Enabled = False
            cmdMuovi(1).Enabled = False
            cmdMuovi(2).Enabled = True
            cmdMuovi(3).Enabled = True
        Case Is = GetNumeroRighe(frmWeb.ListView1)
            cmdMuovi(0).Enabled = True
            cmdMuovi(1).Enabled = True
            cmdMuovi(2).Enabled = False
            cmdMuovi(3).Enabled = False
         Case Else
            cmdMuovi(0).Enabled = True
            cmdMuovi(1).Enabled = True
            cmdMuovi(2).Enabled = True
            cmdMuovi(3).Enabled = True
    End Select
    
    ' Elimino i ToolTipText
    For cnt = 0 To lblDescrizione.UBound
        lblDescrizione(cnt).ToolTipText = ""
    Next
    
    For cnt = 0 To txtAggiunta.UBound
        If txtAggiunta(cnt).Text = "" Then
            txtAggiunta(cnt).BackColor = vbWhite
        Else
            txtAggiunta(cnt).BackColor = vbYellow
        End If
    Next

    For cnt = 0 To txtLen.UBound
        If txtLen(cnt).Text = "" Then
            txtLen(cnt).BackColor = vbWhite
        Else
            txtLen(cnt).BackColor = vbYellow
        End If
    Next
    
    If cmbTaglia.ListCount = 0 Then
        cmbTaglia.BackColor = vbWhite
    Else
        cmbTaglia.BackColor = vbYellow
    End If
    
End Sub

Private Function getArrCampi()
    Dim arrCampi(7) As String
    
    arrCampi(0) = Des
    arrCampi(1) = Indi
    arrCampi(2) = Citta
    arrCampi(3) = Prov
    arrCampi(4) = Tel
    arrCampi(5) = NumCiv
    arrCampi(6) = Cap
    arrCampi(7) = Cat
    getArrCampi = arrCampi
    
End Function

Private Function Campo(TipoCampo As Integer, ListIndex1 As Long, Optional ListIndex2 As Long = -1, Optional TxtDopo As Boolean = False) As String
    Dim cnt As Long
    Dim arrCampi() As String
    Dim Campo1vuoto As Boolean
    Dim DelimCampo1 As String

    arrCampi = getArrCampi
    
    Select Case TipoCampo
        Case 1
            For cnt = 0 To UBound(arrCampi)
            ' Carico i campi (indicati in aListIndex(1))
                If ListIndex1 = cnt Then
                    If ckIncludiDelim = 0 And arrCampi(cnt) = "" Then
                        Campo = " "
                    Else
                        Campo = arrAggiunta(cnt).iPrima & arrCampi(cnt)
                        If TxtDopo = True Or ckIncludiDelim = 1 Then
                            Campo = Campo & arrAggiunta(cnt).iDopo
                        Else
                            Campo = Campo & " "
                        End If
                    End If
                    Exit For
                End If
            Next
        
        Case 2
            For cnt = 0 To UBound(arrCampi)
            ' Carico i campi (indicati in aListIndex(1))
                If ListIndex1 = cnt Then
                    DelimCampo1 = arrAggiunta(cnt).iPrima
                    If ckIncludiDelim = 0 And arrCampi(cnt) = "" Then
                        Campo1vuoto = True
                        Campo = ""
                    Else
                        Campo1vuoto = False
                        Campo = arrAggiunta(cnt).iPrima & arrCampi(cnt)
                        Campo = Campo
                    End If
                    Exit For
                End If
            Next

            For cnt = 0 To UBound(arrCampi)
            ' Carico i campi concatenati (indicati in aListIndex(2))
                If ListIndex2 = cnt Then
                    If ckIncludiDelim = 0 And arrCampi(cnt) = "" And Campo1vuoto = True Then
                        Campo = " "
                    Else
                        If Campo1vuoto = True Then
                            Campo = DelimCampo1
                        Else
                            If Len(Campo) <> Len(DelimCampo1) And arrCampi(cnt) <> "" Then Campo = Campo & " "
                        End If
                        Campo = Campo & arrCampi(cnt)
                        If ckIncludiDelim = 1 Or Campo1vuoto = False Or arrCampi(cnt) <> "" Then
                            Campo = Campo & arrAggiunta(cnt).iDopo
                        Else
                            Campo = Campo & " "
                        End If
                    End If
                    Exit For
                End If
            Next
    End Select
       
End Function

Private Function getOv2desc() As String
    Dim cnt As Long
    Dim cnt1 As Long
    Dim strTmp As String
    
    ' List e Label |       Text         |   Array
    ' Des =    0   | Aggiunta =  0 & 1  | arrAggiunta = 0
    ' Indi =   1   | Aggiunta =  2 & 3  | arrAggiunta = 1
    ' Citta =  2   | Aggiunta =  4 & 5  | arrAggiunta = 2
    ' Prov =   3   | Aggiunta =  6 & 7  | arrAggiunta = 3
    ' Tel =    4   | Aggiunta =  8 & 9  | arrAggiunta = 4
    ' NumCiv = 5   | Aggiunta = 10 & 11 | arrAggiunta = 5
    ' Cap =    6   | Aggiunta = 12 & 13 | arrAggiunta = 6
    ' Cat =    7   | Aggiunta = 14 & 15 | arrAggiunta = 7
    
    For cnt = 0 To lstOrdine.ListCount - 1
        ' aListIndex contiene:
        ' (0)Descrizione
        ' (1)Campo1
        ' (2)Campo2
        aListIndex = Split(lstOrdine.List(cnt), "|", , vbTextCompare)
        '
        For cnt1 = 0 To UBound(arrAggiunta)
            ' Carico i campi (indicati nell'array aListIndex(1))
            If CLng(aListIndex(1)) = cnt1 Then
                ' Se il campo 2 è uguale al campo 1
                If aListIndex(2) = cnt1 Then
                    arrAggiunta(cnt1).iTesto = Campo(1, cnt1, , True)
                Else
                    arrAggiunta(cnt1).iTesto = Campo(2, cnt1, CLng(aListIndex(2)))
                End If
                arrAggiunta(cnt1).iOrdine = cnt
                Exit For
            End If
        Next
    Next
    
    cnt1 = 0 ' Il numero di ordinamento dell'array
    cnt = 0  ' L'indice dell'array
    Do Until cnt1 = UBound(arrAggiunta) + 1
        ' Scorro l'array dall'inizio alla fine
        For cnt = 0 To UBound(arrAggiunta)
            ' Se trovo il campo cercato.....
            If arrAggiunta(cnt).iOrdine = cnt1 Then
                strTmp = strTmp & arrAggiunta(cnt).iTesto
                
                'Exit For
                
                
            End If
        Next
        cnt1 = cnt1 + 1
    Loop
        
    getOv2desc = Trim$(TrimDUP(strTmp))
    
End Function

Private Sub Visualizza(Riga As Long, Optional ScriviRisultato As Boolean = False, Optional SelezionaRiga As Boolean = False, Optional AggiornaLabel As Boolean = True)
    Dim tDes As String, tIndi As String, tNumCiv As String
    Dim tCap As String, tCitta As String, tProv As String
    Dim tTel As String, tCat As String

    Dim Numciv2 As Long, Indi2 As Long, Indi1 As Long, lCampo As Long
    Dim ov2 As String
    Dim ScartoLen As Long
    Dim arrTmp() As String ' per cercare il numero civico
    Dim CampiTagliati As Long
    Dim cnt As Long
    Dim cnt1 As Long
    Dim TagliaIndex As Long
    Dim Esempio As String
    Dim CarMax As Long
  
    If AggiornaLabel = True Then
        Call ElaboraControlli
        Esempio = "Esempio finale di Descrizione OV2:"
    End If
   
    ControllaConcatena (AggiornaLabel)
 
    CarMax = CLng(txtCarMax.Text)
    CampiTagliati = 0
    
    Des = ""
    Indi = ""
    Citta = ""
    Prov = ""
    Tel = 0
    NumCiv = ""
    Cap = ""
    Cat = ""
    Numciv2 = 0
    
    Des = Trim$(GetValoreCella(frmWeb.ListView1, Riga, 1))
    Indi = Trim$(GetValoreCella(frmWeb.ListView1, Riga, 2))
    
    If Indi <> "" Then
        ' Divido la stringa Indi per cercare il numero civico
        arrTmp = Split(Indi, " ", , vbTextCompare)
        ' L'ultimo valore dell'array è il numero civico
        NumCiv = Trim$(arrTmp(UBound(arrTmp)))
        If ContieneNumero(NumCiv) = True Then
            Numciv2 = Len(NumCiv)
            Indi1 = Len(Indi) - Numciv2
            Indi = Trim$(Left(Indi, Indi1))
        Else
            NumCiv = ""
        End If
    End If
    
    Cap = Trim$(GetValoreCella(frmWeb.ListView1, Riga, 3))
    Cat = Trim$(GetValoreCella(frmWeb.ListView1, Riga, 7))
    Citta = Trim$(GetValoreCella(frmWeb.ListView1, Riga, 4))
    Prov = Trim$(GetValoreCella(frmWeb.ListView1, Riga, 5))
    Tel = Trim$(GetValoreCella(frmWeb.ListView1, Riga, 6))
    
    If ckTelInt.value = 1 And Tel <> "" Then
        'If Left(Tel, 1) = "0" Then
            'Tel =Mid(Tel, 2)
        'End If
        If Left(Tel, Len(txtTelInt)) <> txtTelInt.Text Then
            Tel = txtTelInt.Text & Tel
        End If
    End If

    If txtLen(0).Text <> "" Then
        lCampo = txtLen(0).Text
        Des = Left(Des, lCampo)
    End If
    If txtLen(1).Text <> "" Then
        lCampo = txtLen(1).Text
        Indi = Left(Indi, lCampo)
    End If
    If txtLen(2).Text <> "" Then
        lCampo = txtLen(2).Text
        Citta = Left(Citta, lCampo)
    End If
    
    If ckDescrizione(0).value = 0 Then
        Des = ""
        If AggiornaLabel = True Then lblDescrizione(0).Caption = ""
    End If
    If ckDescrizione(1).value = 0 Then
        Indi = ""
        NumCiv = ""
        If AggiornaLabel = True Then lblDescrizione(1).Caption = ""
    End If
    If ckDescrizione(2).value = 0 Then
        Citta = ""
        If AggiornaLabel = True Then lblDescrizione(2).Caption = ""
    End If
    If ckDescrizione(3).value = 0 Then
        Prov = ""
        If AggiornaLabel = True Then lblDescrizione(3).Caption = ""
    End If
    If ckDescrizione(4).value = 0 Then
        Tel = ""
        If AggiornaLabel = True Then lblDescrizione(4).Caption = ""
    End If
    If ckDescrizione(5).value = 0 Then
        Cap = ""
        If AggiornaLabel = True Then lblDescrizione(2).Caption = ""
    End If
    If ckDescrizione(6).value = 0 Then
        Cat = ""
        If AggiornaLabel = True Then lblDescrizione(7).Caption = ""
    End If
    
    If AggiornaLabel = True Then
        lblEsempio(0).ForeColor = vbBlack
        lblEsempio(1).ForeColor = vbBlack
        For cnt = 0 To lblDescrizione.UBound
            lblDescrizione(cnt).ForeColor = vbBlack
        Next
    End If

    If Des <> "" Then tDes = arrAggiunta(0).iPrima & Des & arrAggiunta(0).iDopo
    If Indi <> "" Then tIndi = arrAggiunta(1).iPrima & Indi & arrAggiunta(1).iDopo
    If Citta <> "" Then tCitta = arrAggiunta(2).iPrima & Citta & arrAggiunta(2).iDopo
    If Prov <> "" Then tProv = arrAggiunta(3).iPrima & Prov & arrAggiunta(3).iDopo
    If Tel <> "" Then tTel = arrAggiunta(4).iPrima & Tel & arrAggiunta(4).iDopo
    If NumCiv <> "" Then tNumCiv = arrAggiunta(5).iPrima & NumCiv & arrAggiunta(5).iDopo
    If Cap <> "" Then tCap = arrAggiunta(6).iPrima & Cap & arrAggiunta(6).iDopo
    If Cat <> "" Then tCat = arrAggiunta(7).iPrima & Cat & arrAggiunta(7).iDopo
        
    If AggiornaLabel = True Then
        lblDescrizione(0).Caption = "(" & Format(Len(tDes), "00") & ") " & tDes
        lblDescrizione(1).Caption = "(" & Format(Len(tIndi), "00") & ") " & tIndi
        lblDescrizione(2).Caption = "(" & Format(Len(tCitta), "00") & ") " & tCitta
        lblDescrizione(3).Caption = "(" & Format(Len(tProv), "00") & ") " & tProv
        lblDescrizione(4).Caption = "(" & Format(Len(tTel), "00") & ") " & tTel
        lblDescrizione(5).Caption = "(" & Format(Len(tNumCiv), "00") & ") " & tNumCiv
        lblDescrizione(6).Caption = "(" & Format(Len(tCap), "00") & ") " & tCap
        lblDescrizione(7).Caption = "(" & Format(Len(tCat), "00") & ") " & tCat
    End If

    ' La stringa con la descrizione per il poi .ov2
    ov2 = getOv2desc
    
    ' Se la lunghezza di ov2 supera il numero di caratteri massimi.....
    If CarMax < Len(ov2) Then
        cnt = 1
        If AggiornaLabel = True Then lblEsempio(0).ForeColor = vbRed
        If AggiornaLabel = True Then lblEsempio(1).ForeColor = vbRed

        Select Case cmbTaglia.Text
            Case Is = "Descrizione"
                TagliaIndex = 0
            Case Is = "Indirizzo"
                TagliaIndex = 1
            Case Is = "Città"
                TagliaIndex = 2
            Case Is = ""
                TagliaIndex = 0
            Case Else
                TagliaIndex = 0
        End Select
        
TagliaAncora:
        Do
            Select Case TagliaIndex
                Case Is = 0
                    If AggiornaLabel = True Then lblDescrizione(0).ForeColor = vbRed
                    If Des <> "" Then Des = RTrim(Left$(Des, Len(Des) - 1))
                    ov2 = getOv2desc
                    If CarMax >= Len(ov2) Or Des = "" Then
                        If AggiornaLabel = True Then lblDescrizione(0).ToolTipText = " Nella Descrizione Ov2 verrà cambiato in: " & Des & " "
                        Exit Do
                    End If
                Case Is = 1
                    If AggiornaLabel = True Then lblDescrizione(1).ForeColor = vbRed
                    If Indi <> "" Then Indi = RTrim(Left$(Indi, Len(Indi) - 1))
                    ov2 = getOv2desc
                    If CarMax >= Len(ov2) Or Indi = "" Then
                        If AggiornaLabel = True Then lblDescrizione(1).ToolTipText = " Nella Descrizione Ov2 verrà cambiato in: " & Indi & " "
                        Exit Do
                    End If
                Case Is = 2
                    If AggiornaLabel = True Then lblDescrizione(2).ForeColor = vbRed
                    If Citta <> "" Then Citta = RTrim(Left$(Citta, Len(Citta) - 1))
                    ov2 = getOv2desc
                    If CarMax >= Len(ov2) Or Citta = "" Then
                        If AggiornaLabel = True Then lblDescrizione(2).ToolTipText = " Nella Descrizione Ov2 verrà cambiato in: " & Citta & " "
                        Exit Do
                    End If
                Case Else
                    ov2 = ov2
                    Exit Do
            End Select
            cnt = cnt + 1
        Loop
        
        ' Se la lunghezza di ov2 supera il numero di caratteri massimi.....
        ' Cambio TagliaIndex e ripeto il ciclo Do
        If CarMax < Len(ov2) Then
            Select Case CampiTagliati
                Case Is = 0
                    CampiTagliati = CampiTagliati + 1
                    If TagliaIndex = 0 Then
                        TagliaIndex = 1
                    Else
                        TagliaIndex = 0
                    End If
                    GoTo TagliaAncora:
                Case Is = 1
                    CampiTagliati = CampiTagliati + 1
                    TagliaIndex = TagliaIndex + 1
                    GoTo TagliaAncora:
                Case Is = 2
                    CampiTagliati = CampiTagliati + 1
                    TagliaIndex = TagliaIndex + 1
                    GoTo TagliaAncora:
            End Select
        End If
        CampiTagliati = CampiTagliati + 1
    End If

    If AggiornaLabel = True Then
        If CampiTagliati <> 0 Then Esempio = "(Ho tagliato i valori in " & CampiTagliati & " campo/i)" & " " & Esempio
        lblRecord.Caption = Riga
        lblEsempio(0).Caption = Esempio
        lblEsempio(1).Caption = "(" & Format(Len(ov2), "00") & ") " & ov2
    End If
    
    If SelezionaRiga = True Then
        Call SelezionaRigaListView(frmWeb.ListView1, RigaCorrente, AggiornaLabel)
    End If

    ' Se serve scrivo il risultato nell'apposita colonna della ListView
    If ScriviRisultato = True Then
        Call ScriviCella(frmWeb.ListView1, Riga, GetNumColDaIntestazione(frmWeb.ListView1, "Desc.ov2"), ov2)
    End If
    
    If FormCaricata = True Then PreparaSetupDescrizione (0)
    
End Sub

Public Function PreparaSetupDescrizione(ByVal Azione As Integer, Optional ByVal strSetupDescrizione As String = "") As String
    ' Prepara la stringa di testo che contiene le impostazioni della form
    ' Restituisce la stringa preaparata
    Dim cnt As Integer
    Dim cnt1 As Integer
    Dim arrTmpSetDesc
    Dim cntArr As Integer

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    Select Case Azione
        Case Is = 0 ' 0 Salvo i dati nella variabile SetupDescrizione
            SetupDescrizione = ""
            
            ' Salvo le impostazione del testo di separazione
            For cnt = 0 To (lblDescrizione.UBound)
                SetupDescrizione = SetupDescrizione & txtAggiunta(cnt1).Text & ","
                SetupDescrizione = SetupDescrizione & txtAggiunta(cnt1 + 1).Text & ","
                cnt1 = cnt1 + 2
            Next
            
            ' Salvo le impostazione dei CheckBox
            For cnt = 0 To ckDescrizione.UBound
                SetupDescrizione = SetupDescrizione & ckDescrizione(cnt).value & ","
            Next
            SetupDescrizione = SetupDescrizione & ckTelInt.value & ","
        
            ' Salvo le impostazione dei txtLen
            For cnt = 0 To txtLen.UBound
                SetupDescrizione = SetupDescrizione & txtLen(cnt).Text & ","
            Next
            
            ' Salvo le impostazione di ckIncludiDelim
            SetupDescrizione = SetupDescrizione & ckIncludiDelim.value & ","
            
            ' Salvo le impostazione di cmbTaglia
            SetupDescrizione = SetupDescrizione & cmbTaglia.ListIndex
            
        Case Is = 1 'Carico i dati dalla variabile SetupDescrizione
            If strSetupDescrizione = "" Then
                SetupDescrizione = arrCampiRmkFile(UBound(arrCampiRmkFile))
            Else
                SetupDescrizione = strSetupDescrizione
            End If
            
            If SetupDescrizione = "" Or InStr(1, SetupDescrizione, ",", vbTextCompare) = 0 Then
                Exit Function
            End If

            arrTmpSetDesc = Split(SetupDescrizione, ",", , vbTextCompare)
            
            cntArr = 0
            ' Carico le impostazione del testo di separazione
            For cnt = 0 To (lblDescrizione.UBound)
                txtAggiunta(cntArr).Text = arrTmpSetDesc(cntArr)
                txtAggiunta(cntArr + 1).Text = arrTmpSetDesc(cntArr + 1)
                cntArr = cntArr + 2
            Next
            
            ' Carico le impostazione dei CheckBox
            For cnt = 0 To ckDescrizione.UBound
                ckDescrizione(cnt).value = arrTmpSetDesc(cntArr)
                cntArr = cntArr + 1
            Next
            ckTelInt.value = arrTmpSetDesc(cntArr)
            cntArr = cntArr + 1
            
            ' Carico le impostazione dei txtLen
            For cnt = 0 To txtLen.UBound
                txtLen(cnt).Text = arrTmpSetDesc(cntArr)
                cntArr = cntArr + 1
            Next

            ' Carico le impostazione di ckIncludiDelim
            ckIncludiDelim.value = arrTmpSetDesc(cntArr)
            cntArr = cntArr + 1

            ' Carico le impostazione di cmbTaglia
            cmbTaglia.ListIndex = arrTmpSetDesc(cntArr)
            cntArr = cntArr + 1
        
    End Select

    txtSetupDescrizione.Text = SetupDescrizione
    txtSetupDescrizione.ToolTipText = SetupDescrizione
    PreparaSetupDescrizione = SetupDescrizione
    
    Exit Function
    
Errore:
    PreparaSetupDescrizione = ""
    GestErr Err, "Errore nella funzione PreparaSetupDescrizione."
    
End Function

Private Sub VScroll1_Change()
    txtSostCount.Text = VScroll1.value
End Sub
