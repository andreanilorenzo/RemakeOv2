VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCheckOV2 
   Caption         =   "Verifica File"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12945
   Icon            =   "frmCheckOv2 .frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   12945
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBmp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   120
      Width           =   375
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
      Height          =   2220
      Left            =   1440
      ScaleHeight     =   2190
      ScaleWidth      =   3975
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.CommandButton cmdImportaFile 
      Caption         =   "&Apri file"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   11895
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6825
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   12039
      View            =   2
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
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
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "frmCheckOV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Obbliga a dichiarare le costanti
Option Explicit

Dim LarghezzaColonne As Variant
Dim TotLarghezzaCol As Long

Dim FileDaAprire As String
Dim FormInApertura As Boolean

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If FormIsLoad("frmMain") = True Then frmMain.Visible = False
End Sub

Private Sub Form_Load()
    Dim I As Long
        
    FormInApertura = True
    
    LarghezzaColonne = Array(0, 800, 980, 980, 980, 980, 1900, 980, 980, 5500)
    
    For I = 0 To UBound(LarghezzaColonne)
        TotLarghezzaCol = TotLarghezzaCol + LarghezzaColonne(I)
    Next
    
    With ListView1
        .View = lvwReport
        .MultiSelect = False
        .FullRowSelect = True
        .ColumnHeaders.Add , "x1", "    "
        .ColumnHeaders.Item(1).Width = LarghezzaColonne(1)
        .ColumnHeaders.Add , "x2", " A "
        .ColumnHeaders.Item(2).Width = LarghezzaColonne(2)
        .ColumnHeaders.Add , "x3", " B "
        .ColumnHeaders.Item(3).Width = LarghezzaColonne(3)
        .ColumnHeaders.Add , "x4", " C "
        .ColumnHeaders.Item(4).Width = LarghezzaColonne(4)
        .ColumnHeaders.Add , "x5", " D "
        .ColumnHeaders.Item(5).Width = LarghezzaColonne(5)
        .ColumnHeaders.Add , "x6", " E (Distanza)"
        .ColumnHeaders.Item(6).Width = LarghezzaColonne(6)
        .ColumnHeaders.Add , "x7", " F (Lat.) "
        .ColumnHeaders.Item(7).Width = LarghezzaColonne(7)
        .ColumnHeaders.Add , "x8", " G (Long.) "
        .ColumnHeaders.Item(8).Width = LarghezzaColonne(8)
        .ColumnHeaders.Add , "x9", " H (Descrizione) "
        .ColumnHeaders.Item(9).Width = LarghezzaColonne(9)
    End With

    Call SetListViewColor(ListView1, Picture1, 1, vbWhite, vbGreenLemon)

    FileDaAprire = PercorsoFileList
    PercorsoFileList = ""
    cmdImportaFile_Click
    FormInApertura = False
    
End Sub

Private Sub Form_Resize()
    Dim Sin As Integer 'Il borbo da lasciare a sinistra
    Dim Inf As Integer 'Il bordo da lasciare sotto
    
    On Error Resume Next
    
    Sin = 0
    Inf = 0
    ListView1.Move ListView1.Left, ListView1.Top, Me.ScaleWidth - ListView1.Left - Sin, Me.ScaleHeight - ListView1.Top - Inf

    With cmdImportaFile
        ' Centro il controllo nella form
        .Move (Me.ScaleWidth - .Width) / 2
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.Visible = True
    frmMain.WindowState = vbNormal
    frmMain.ZOrder
    frmMain.SetFocus

End Sub

Private Sub cmdImportaFile_Click()
    Dim ret
    
    Screen.MousePointer = vbHourglass
    DoEvents
   
    LockWindowUpdate ListView1.hWnd

    If ImportaDati(Me.hWnd, , FileDaAprire, "ov2|asc", False, True, , "*.ov2", , , True) = True Then
        If Right$(FileDaAprire, 4) <> ".rmk" Then
            Me.Caption = "Verifica File:  " & FileDaAprire
            Call CaricaImmagine(picBmp, FileDaAprire, "bmp")
            Call InserisciListView
        Else
            Me.Caption = "Verifica File"
            ret = MsgBox("Da questa finestra non puoi aprire questo formato di file (.rmk)!   ", vbInformation)
        End If
    Else
        If FormInApertura = True Then
            ' Chiudo la form
            Unload Me
            GoTo Esci
        End If
        
        Me.Caption = "Verifica File"
        If FileDaAprire <> "" Then
            ret = MsgBox("Da questa finestra non puoi aprire questo formato di file (.rmk)!   ", vbInformation)
            FileDaAprire = ""
        End If
    End If
    
Esci:
    LockWindowUpdate 0&
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub InserisciListView()
    Dim itmX As Variant
    Dim I
    
    On Error Resume Next
    
    ListView1.ListItems.Clear
    ListView1.Sorted = False
    'Così si può cancellare anche l'intestazione
    'ListView1.ColumnHeaders.Clear
    DoEvents
    For I = 0 To UBound(ArrayOv2PoiRec) - 1 ' Scorro tutte le righe della ListView che servono
        Set itmX = ListView1.ListItems.Add(, , Format(I + 1, "00000"))
        If ArrayOv2PoiRec(I).aTy1PoiLatitude <> 0 Then itmX.SubItems(1) = ArrayOv2PoiRec(I).aTy1PoiLatitude
        If ArrayOv2PoiRec(I).bTy1PoiLongitude <> 0 Then itmX.SubItems(2) = ArrayOv2PoiRec(I).bTy1PoiLongitude
        If ArrayOv2PoiRec(I).cTy1Poi3Latitude <> 0 Then itmX.SubItems(3) = ArrayOv2PoiRec(I).cTy1Poi3Latitude
        If ArrayOv2PoiRec(I).dTy1Poi3Longitude <> 0 Then itmX.SubItems(4) = ArrayOv2PoiRec(I).dTy1Poi3Longitude
        If ArrayOv2PoiRec(I).eTy1Distanza <> 0 Then itmX.SubItems(5) = ArrayOv2PoiRec(I).eTy1Distanza
        If ArrayOv2PoiRec(I).fTy2PoiLatitude <> 0 Then itmX.SubItems(6) = ArrayOv2PoiRec(I).fTy2PoiLatitude
        If ArrayOv2PoiRec(I).gTy2PoiLongitude <> 0 Then itmX.SubItems(7) = ArrayOv2PoiRec(I).gTy2PoiLongitude
        itmX.SubItems(8) = ArrayOv2PoiRec(I).hTy2descrizione
    Next I
    FileDaAprire = ""

End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    
    If Data.GetFormat(vbCFFiles) Then
        For I = 1 To 1 'Data.Files.count
            FileDaAprire = (Data.Files(I))
            Call ImportaDati(Me.hWnd, , FileDaAprire)
            Call InserisciListView
        Next
    End If
    
End Sub
