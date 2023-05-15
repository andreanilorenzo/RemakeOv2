VERSION 5.00
Begin VB.Form frmRinomina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rinomina File"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   Icon            =   "frmRinomina.frx":0000
   LinkTopic       =   "frmRinomina"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox cekSovrascrivi 
      Caption         =   "&Sovrascrivi i file esistenti con quelli rinominati"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   5160
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.Timer tmrCaricaFile 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3840
      Top             =   4920
   End
   Begin VB.CommandButton cmdRinomina 
      Caption         =   "&Rinomina i file"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtEst2 
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
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtEst 
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
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstFile 
      Height          =   4350
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   7575
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "&Esci  [Esc]"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblEst2 
      Caption         =   "Altra estensione associata:"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblEst 
      Caption         =   "Estensioni dei file:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmRinomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LBHS As clsListBox
Dim arrSostituzioni() As String
Dim sDirLavoro As String            ' La directory che contiene i file

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim arrFile() As String
    Dim cnt As Integer
    Dim sEstensioni As String
    Dim sFolder As String

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    ' Inizializzo la classe per la ListBox
    Set LBHS = New clsListBox
    LBHS.Attach lstFile
    
    txtEst.Text = ".ov2"
    txtEst2.Text = ".bmp"
    
    sDirLavoro = Var(PoiScaricati).Valore
    
    Call CaricaElencoFile(txtEst.Text & "|" & txtEst2.Text)
    
    frmMain.Visible = False
    
    Exit Sub

Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione frmRinomina.Form_Load."

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Visible = True
    frmMain.WindowState = vbNormal
    frmMain.ZOrder
    frmMain.SetFocus

End Sub

Private Sub CaricaElencoFile(ByVal sEstensioni As String)
    Dim arrFile() As String
    Dim cnt As Integer

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore
    
    Screen.MousePointer = vbHourglass
    
    ' Pulisco la listbox
    LBHS.Clear
    
    If GetFileInFolder(arrFile, sDirLavoro, sEstensioni) = True Then
        For cnt = 0 To UBound(arrFile)
            LBHS.AddItem sDirLavoro & "\" & arrFile(cnt)
        Next
        LBHS.ListIndex = 0
        
    Else
        MsgBox "Non sono stati trovati file nella cartella" & vbNewLine & Var(PoiScaricati).Valore, vbInformation, App.ProductName
    End If

    Screen.MousePointer = vbDefault
    
    Exit Sub

Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione frmRinomina.CaricaElencoFile."

End Sub

Private Sub cmdRinomina_Click()
    Dim cnt As Integer
    Dim cnt1 As Integer
    Dim sNomeTmp, oldNomeTmp As String
    Dim FullPatchNuovo As String
    
    LoadFile Var(Sostituisci).Valore, arrSostituzioni
    
    For cnt = 0 To LBHS.ListCount - 1
        
        sNomeTmp = FileNameFromPath(LBHS.List(cnt))
        oldNomeTmp = sNomeTmp
        
        For cnt1 = 0 To UBound(arrSostituzioni, 2)
            sNomeTmp = Replace(sNomeTmp, arrSostituzioni(0, cnt1), arrSostituzioni(1, cnt1))
        Next
        
        If sNomeTmp <> oldNomeTmp Then
            FullPatchNuovo = DirectoryFromFile(LBHS.List(cnt), False) & sNomeTmp
            If (FileExists(FullPatchNuovo) = True And cekSovrascrivi.value = 1) Then
                 ' Se il file esiste lo cancello
                fso.DeleteFile FullPatchNuovo
                DoEvents
                
            ElseIf (FileExists(FullPatchNuovo) = True And cekSovrascrivi.value = 0) Then
                sNomeTmp = ""
                
            End If
            
            ' Rinomino il file
            If sNomeTmp <> "" Then fso.Rename LBHS.List(cnt), sNomeTmp
        End If
        
    Next
    
    Call CaricaElencoFile(txtEst.Text & "|" & txtEst2.Text)
    
End Sub

Private Function LoadFile(ByVal FullNameFile As String, ByRef arrSost() As String, Optional ByVal Separatore As String) As Boolean
    ' Legge i dati da file di testo.........
    Dim MyString As String
    Dim File1 As Integer
    Dim strTmp() As String
    Dim cnt As Integer
    
    On Error Resume Next
    
    If FileExists(FullNameFile) = False Then
        LoadFile = False
        Exit Function
    End If
    
    ReDim arrSost(2, 1000)
    
    If Separatore = "" Then Separatore = "|"
    
    File1 = FreeFile
    
    Open FullNameFile For Input Access Read As #File1
    Do While EOF(File1) = False
        Line Input #File1, MyString
        strTmp = Split(MyString, Separatore, , vbTextCompare)
        arrSost(0, cnt) = strTmp(0)
        arrSost(1, cnt) = strTmp(1)
        cnt = cnt + 1
    Loop
    Close #File1
    
    ReDim Preserve arrSost(2, cnt)
    
    LoadFile = True

End Function

Private Sub tmrCaricaFile_Timer()

    tmrCaricaFile.Enabled = False
    Call CaricaElencoFile(txtEst.Text & "|" & txtEst2.Text)

End Sub

Private Sub txtEst_KeyPress(KeyAscii As Integer)
    tmrCaricaFile.Enabled = True
End Sub
Private Sub txtEst2_KeyPress(KeyAscii As Integer)
    tmrCaricaFile.Enabled = True
End Sub

