VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmActProg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmActProg"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   Icon            =   "frmActProg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmActProg.frx":23D2
   ScaleHeight     =   7275
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin Remakeov2.XpBs XpBs1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      Caption         =   "Continua ad utilizzare la versione &Freeware del programma"
      ButtonStyle     =   4
      Picture         =   "frmActProg.frx":8C4D
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   14
      OriginalPicSizeH=   14
      PictureHover    =   "frmActProg.frx":8F07
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      XPColor_Pressed =   13461299
      XPColor_Hover   =   13461299
      BackColor       =   13461299
      ForeColor       =   16777215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00CD6733&
      Caption         =   "Controllo Versione  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   1320
      TabIndex        =   10
      Top             =   3600
      Width           =   6735
      Begin VB.Frame Frame3 
         BackColor       =   &H00CD6733&
         Caption         =   "Sul Web "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   2160
         TabIndex        =   14
         Top             =   600
         Width           =   1695
         Begin VB.Label lblWebVer 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "?.?.?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   0
            TabIndex        =   17
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00CD6733&
         Caption         =   "Attuale "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1695
         Begin VB.Label lblCurrVers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0.0.0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   0
            TabIndex        =   16
            Top             =   360
            Width           =   1695
         End
      End
      Begin Remakeov2.XpBs cmdControlla 
         Height          =   495
         Left            =   4440
         TabIndex        =   7
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "&Controlla versione"
         ButtonStyle     =   3
         PictureWidth    =   16
         PictureHeight   =   16
         PictureSize     =   0
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
      End
      Begin VB.Label lblRet 
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   4935
      End
      Begin VB.Label lblConnesso 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota: devi essere connesso ad internet per utilizzare questo servizio."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   700
      Left            =   0
      Picture         =   "frmActProg.frx":9259
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   6000
      Width           =   700
   End
   Begin Remakeov2.XpBs XpBs2 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      Caption         =   "&Attiva il programma tramite internet"
      ButtonStyle     =   4
      Picture         =   "frmActProg.frx":BCFA
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   14
      OriginalPicSizeH=   14
      PictureHover    =   "frmActProg.frx":BFB4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      XPColor_Pressed =   13461299
      XPColor_Hover   =   13461299
      BackColor       =   13461299
      ForeColor       =   16777215
   End
   Begin Remakeov2.XpBs XpBs3 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      Caption         =   "&Visita la pagina web del programma"
      ButtonStyle     =   4
      Picture         =   "frmActProg.frx":C306
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   14
      OriginalPicSizeH=   14
      PictureHover    =   "frmActProg.frx":C5C0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      XPColor_Pressed =   13461299
      XPColor_Hover   =   13461299
      BackColor       =   13461299
      ForeColor       =   16777215
   End
   Begin Remakeov2.XpBs cmbEsci 
      Height          =   495
      Left            =   8400
      TabIndex        =   8
      Top             =   6720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "&Esci  [Esc]"
      ButtonStyle     =   3
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.CommandButton cmdEsciDue 
      Cancel          =   -1  'True
      Caption         =   "&Esci  [Esc]"
      Height          =   855
      Left            =   6360
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   " Esce dal programma "
      Top             =   3960
      Width           =   1455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7680
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblInfo 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   960
      Width           =   8295
   End
   Begin VB.Label lblAiutoDesc 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmActProg.frx":C912
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   6480
      Width           =   7455
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "Aiuto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label lblMaster 
      BackStyle       =   0  'Transparent
      Caption         =   "lblMaster"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   840
      TabIndex        =   4
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmActProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim FileApplicationVer As String
Dim FileWebProgramma As String
Dim FileNews As String

Private Sub cmdEsciDue_Click()
    cmbEsci_Click
End Sub

Private Sub cmbEsci_Click()
    Unload Me
End Sub

Private Sub cmdControlla_Click()
    'This function assume files "NomeProgramma.ver", "news.txt" and "application.zip"
    'on server http://server.com/user
    'Inspect contain of files "news.txt" and "application.ver" at examples
    Dim version As String
    Dim News As String
    Dim strTmp As String
    
    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    Me.MousePointer = vbHourglass
    
    lblRet.Caption = "---"
    lblWebVer.Caption = "?.?.?"
    DoEvents
    
    ' Now assign content of file application.ver to variable Version
    'version = Inet1.OpenURL("File://" + App.Path & "/" & "files" & "/" & "application.ver")
    version = SplitOne(Inet1.OpenURL(FileApplicationVer), vbCrLf, 0)
    version = Replace(Replace(version, vbCr, ""), vbLf, "")
    
    ' If file not found or file is empty then exit
    If (version = "") Or (IsNumeric(Replace(Replace(version, "b", ""), ".", "")) = False) Then
        lblWebVer.Caption = "!.!.!"
        lblRet.Caption = "Non è stato possibile controllare se sono disponibili nuove versioni"
        GoTo Skip
    End If
    
    lblWebVer.Caption = version
    
    If version = GetProgVers Then
        lblRet.Caption = "Questa versione è gia aggiornata"
        GoTo Skip
        
    ElseIf version < GetProgVers Then
        lblRet.Caption = "Questa versione è più aggiornata di quella disponibile sul web"
        GoTo Skip
    End If
    
    strTmp = Inet1.OpenURL(FileNews)
    News = Trim$(Right$(strTmp, Len(strTmp) - InStr(1, strTmp, "News:", vbTextCompare) - 4))
    
    ' Now display MessageBox with news in newer version of application and two buttons Yes(update), No(end)
    If MsgBox("E' disponibile una nuova versione del programma." & vbNewLine & vbNewLine & vbNewLine & News, vbInformation + vbYesNo, Me.Caption) = vbYes Then
        MsgBox "Verrà aperta la pagina web del programma.", vbInformation, App.ProductName
        HyperJump "http://remakeov2.poigps.com/index.php?option=com_docman&task=cat_view&gid=13&Itemid=27"
        'HyperJump "file://" + App.Path & "/" & "files" & "/" & "application.zip" 'this will run default download manager (probable also open default browser)
        'HyperJump FileWebProgramma
    Else
        lblRet.Caption = "Download della nuova versione annullato dall'utente"
    End If
    
Skip:
    Me.MousePointer = vbDefault
    Exit Sub
    
Errore:
    Me.MousePointer = vbDefault
    strTmp = "Controllo versione fallito." & vbNewLine & "Devi controllare la nuova versione manualmente dal sito web."
    GestErr Err, "Errore nella funzione cmdControlla_Click." & vbNewLine & strTmp
    
End Sub

Private Sub Form_Activate()
    If FormIsLoad("frmMain") = True Then frmMain.Visible = False
End Sub

Private Sub Form_Load()
    Dim DirHttp As String
    
    Me.Caption = Replace(Versione, vbNewLine, " - ")
    
    DirHttp = HomePage & "/dmdocuments/"
    FileWebProgramma = DirHttp & "RemakeOv2.zip"
    FileApplicationVer = DirHttp & "RemakeOv2.ver"
    FileNews = DirHttp & "RemakeOv2.ver"
    
    If VerificaStatoInternet = True Or CStr(VerificaStatoInternet) = "Vero" Or CStr(VerificaStatoInternet) = "True" Then
        lblConnesso.Caption = "Connessione ad internet attiva. " & "Premi il pulante """ & cmdControlla.Caption & """ per avviare il controllo."
    End If
    
    lblMaster.Caption = "Aggiorna " & App.ProductName
    lblCurrVers.Caption = GetProgVers
    lblInfo.Caption = "Da questa finestra puoi controllare se sono disponibili nuove versioni del programma."
    lblAiutoDesc.Caption = "Se vuoi maggiori informazioni sul programma visita il sito web" & vbNewLine & HomePage
    
    'to verify the registration file if registered ,verifies _check.ini and compares Reg ID and Product ID
    '_check.ini file will be creted when activated
    'if _check.ini is not available then Trial is not diabled
    Close #1
    Dim regname
    Dim productid
    On Error GoTo errors
    
    Open App.path & "\" & "_check.ini" For Input As #1
    Dim Code1 As Single
    Dim i
    Dim zip
    Dim final
    Line Input #1, regname
    Line Input #1, productid
    For i = 1 To Len(regname) - 1
        Code1 = Format(ASC(Right(regname, Len(regname) - i)) * 2 + (79 / i) + (i + 3 / 71), "#.#")
        zip = zip & Code1
    Next i
    zip = Right(zip, 8)
    
    For i = 1 To Len(zip) - 1
        Code1 = Format(ASC(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 9), "#00")
        final = final & Code1
    Next i
    final = Right(final, Len(final) - 4)
    final = final & ASC(regname)
    If final = productid Then
    'Form2.Label1.Caption = 0
    'Form2.Label.Caption = "Registered"
    'Form2.Xp_ProgressBar1.Visible = False
    'Form2.Label4.Visible = True
    'Form2.Label.Visible = False
    'Form2.Label2.Visible = False
    'Form2.XpBs1.Visible = False
    'Form2.Label4.Caption = "Now you are in registered mode. Delete _check.ini in apps folder to setback trial"
    XpBs1.Caption = "Enter Registered Software        "
    
    Close #1
    End If
    
errors:     'Form1.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAbout.Show
    frmAbout.SetFocus
End Sub

Private Sub XpBs1_Click()
'Form2.Show
Unload Me

End Sub

Private Sub XpBs2_Click()
'Form3.Show
Unload Me

End Sub

Private Sub XpBs3_Click()
    HyperJump HomePage
End Sub

Private Function HyperJump(ByVal url As String) As Long
    
    HyperJump = ShellExecute(0&, vbNullString, url, vbNullString, vbNullString, vbNormalFocus)

End Function



