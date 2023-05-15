VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRegEstensioni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registra Estensioni"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
   Icon            =   "frmRegEstensioni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      Picture         =   "frmRegEstensioni.frx":030A
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ImageList ImageListFile 
      Left            =   1200
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRegistra 
      Caption         =   "&Registra estensioni file"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4560
      Width           =   2175
   End
   Begin MSComctlLib.ImageList imglRunSearch 
      Left            =   2880
      Top             =   1440
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
            Picture         =   "frmRegEstensioni.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":092E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":0C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":33FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":5104
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":541E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":5578
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":56D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":582C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":5B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":6998
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":77EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":8694
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":94E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":BC98
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":BFB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":CE04
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":D11E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":E778
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":EA92
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":EDAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":F0C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":F3E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":FCBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":10594
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":113E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":12238
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":1308A
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":13964
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":147B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":164C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":167DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":1762C
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":19DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":1C590
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":1E29A
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":20A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":22756
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":23030
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":23E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":254DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":25DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":26690
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":274E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":28B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2998E
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":29CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2A582
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2A89C
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2ABB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2AED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2B02A
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2B344
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2B65E
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2B978
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2BC92
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2BFAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2C2C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2C420
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2C57A
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEstensioni.frx":2C6D4
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
      Left            =   2160
      ScaleHeight     =   990
      ScaleWidth      =   2295
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   2325
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3075
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   7500
      _ExtentX        =   13229
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
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "&Esci  [Esc]"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblInfo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   3840
      Width           =   7575
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblHelp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmRegEstensioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute
Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1

Private Type SHFILEINFO
    hIcon As Long                      '  out: icon
    iIcon As Long                      '  out: icon index
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80          '  out: type name
End Type
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type
Private Type CLSID
    id(16) As Byte
End Type

Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private AppNameFile As String
Private AppPatchFile As String
Private LenAppPatchFile As Long
Private hKey As ERegistryClassConstants
Private Estensioni

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If FormIsLoad("frmMain") = True Then frmMain.Visible = False
End Sub

Private Sub Form_Load()
    Dim LarghezzaColonne As Variant
    Dim ColonneListView1 As Variant
    Dim TagColonne As Variant
    Dim strTmp As String
    Dim cnt As Long
    
    ' Creo gli array con i dati delle colonne
    LarghezzaColonne = Array(1000, 1000, 3000)
    ColonneListView1 = Array("", "Estensione", "Programma")
        ' Serve per la funzione di ordinamento dei dati nella colonna
          TagColonne = Array("NUMBER", "STRING", "STRING")
    
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

    With ImageListFile
        .ListImages.Clear
        .ImageWidth = 16
        .ImageHeight = 16
    End With
    ImageListFile.ListImages.Add Picture:=picFile.Picture
    ListView1.SmallIcons = ImageListFile

    Call SetListViewColor(ListView1, Picture1, 1, vbWhite, vbGreenLemon)
    Call AutoSizeUltimaColonna(ListView1)

    strTmp = "Da questa finestra si possono registrare i file per essere aperti in modo automatico facendovi doppio click con il mouse"
    lblHelp.Caption = strTmp

    ' Definisco la path completa del file .exe
    AppPatchFile = App.Path
    If Right$(AppPatchFile, 1) <> "\" Then AppPatchFile = AppPatchFile & "\"
    AppPatchFile = AppPatchFile & App.EXEName
    If LCase$(Right$(AppPatchFile, 4)) <> ".exe" Then AppPatchFile = AppPatchFile & ".exe"

    AppNameFile = "RemakeOv2 File"
    LenAppPatchFile = Len(AppPatchFile)
    hKey = HKEY_CLASSES_ROOT
    
    lblInfo.Caption = ""
    
    Estensioni = Array(".rmk", ".ov2", ".asc", ".kml", ".gpx", ".csv")
    
    Call LeggiEstensioni
    
End Sub

Private Function LeggiEstensioni()
    Dim cnt As Integer
    Dim strTmp As String
    Dim itmX As Variant
    
    ListView1.ListItems.Clear
    
    For cnt = 0 To UBound(Estensioni)
    
        ' Carico l'icona del file
        picFile.Picture = GetTypeIcon(Estensioni(cnt), SHGFI_SMALLICON)
        ImageListFile.ListImages.Add Picture:=picFile.Picture
        
        Set itmX = ListView1.ListItems.Add(, , Format(cnt + 1, "000"), cnt + 2, cnt + 2)
        itmX.SubItems(GetNumColDaIntestazione(ListView1, "Estensione") - 1) = Estensioni(cnt)
        
        strTmp = GetAssociatedApp(CStr(Estensioni(cnt)))
        If strTmp <> "" Then
            strTmp = Left$(strTmp, Len(strTmp) - 3)
            If LCase$(strTmp) = LCase$(AppPatchFile) Then ListView1.ListItems.Item(cnt + 1).Checked = True
            itmX.SubItems(GetNumColDaIntestazione(ListView1, "Programma") - 1) = strTmp
        End If
    Next
    
    ListView1.Refresh
    Call ControllaCheck(ListView1, False)
    
End Function

Private Sub cmdRegistra_Click()
        
    lblInfo.Caption = ""
    Call RegistraEstensioni
    Call LeggiEstensioni

End Sub

Private Function RegistraEstensioni()
    Dim cntReg As Integer
    Dim cnt As Integer
    Dim Est As String
    
    cntReg = 0
    
    If IsRegistryEditable = True Then
        
        For cnt = 1 To ListView1.ListItems.Count
            Est = GetValoreCella(ListView1, cnt, 1)
            If ListView1.ListItems.Item(cnt).Checked = True Then
                If Left$(GetAssociatedApp(Est), LenAppPatchFile) <> AppPatchFile Then
                    RegistraFileExtension Est, AppPatchFile, AppNameFile
                    cntReg = cntReg + 1
                End If
            Else
                ' Tolgo l'associazione al programma
                UnRegistraFileExtension Est, AppNameFile, False
                cntReg = cntReg + 1
            End If
       
        Next
    
    Else
    
        lblInfo.Caption = "ATTENZIONE! Non puoi effettuare modifiche sul registro. Operazione annullata."
    End If
    
    If cntReg = 0 Then
        lblInfo.Caption = "Nessuna operazione eseguita!"
    Else
        lblInfo.Caption = "Operazione eseguita! Registrazioni fatte: " & cntReg
    End If

End Function

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
    frmMain.WindowState = vbNormal
    frmMain.ZOrder
    frmMain.SetFocus
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ControllaRigaChecked(ListView1, Item, False)
End Sub

Private Function GetTypeIcon(ByVal Filename As String, icon_size As Long) As IPictureDisp
    Dim index As Integer
    Dim hIcon As Long
    Dim item_num As Long
    Dim icon_pic As IPictureDisp
    Dim sh_info As SHFILEINFO

    SHGetFileInfo Filename, FILE_ATTRIBUTE_NORMAL, sh_info, Len(sh_info), SHGFI_USEFILEATTRIBUTES Or (SHGFI_ICON + icon_size)

    hIcon = sh_info.hIcon
    Set icon_pic = IconToPicture(hIcon)
    Set GetTypeIcon = icon_pic
    
End Function

Private Function IconToPicture(hIcon As Long) As IPictureDisp
    ' Convert an icon handle into an IPictureDisp
    Dim cls_id As CLSID
    Dim hRes As Long
    Dim new_icon As TypeIcon
    Dim lpUnk As IUnknown

    With new_icon
        .cbSize = Len(new_icon)
        .picType = vbPicTypeIcon
        .hIcon = hIcon
    End With
    
    With cls_id
        .id(8) = &HC0
        .id(15) = &H46
    End With
    
    hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
    
    If hRes = 0 Then Set IconToPicture = lpUnk
    
End Function


