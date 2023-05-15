VERSION 5.00
Begin VB.Form frmMessageCombo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmMessageCombo"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   Icon            =   "frmMessageCombo.frx":0000
   LinkTopic       =   "frmMessageCombo"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalva 
      Cancel          =   -1  'True
      Caption         =   "&Salva"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   380
      Width           =   1095
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "&Annulla"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblMessage"
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
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label lblTimer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblTimer"
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
      Left            =   4560
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleziona e premi invio"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   5415
   End
End
Attribute VB_Name = "frmMessageCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Risultato As String
Dim tim As Integer
Dim FormCaricata As Boolean
Dim ValoreIniziale As String
Const totTim As Integer = 10

Public Function ApriForm(ByVal arrCombo, Optional ByVal Predefinito As String = "", Optional ByVal frmCaption As String = "", Optional ByVal MsgCaption As String = "") As String
    Dim cnt As Integer
    
    FormCaricata = False
    
    Load Me
    
    If frmCaption = "" Then frmCaption = App.ProductName
    Me.Caption = frmCaption
    lblMessage.Caption = MsgCaption
    lblTimer.Caption = totTim
    
    For cnt = 0 To UBound(arrCombo)
        Combo1.AddItem arrCombo(cnt)
    Next
    
    If Combo1.ListCount >= 1 Then
        Combo1.ListIndex = 0
    End If
    If Predefinito <> "" Then
        CercaInComboBox Combo1, Predefinito
    End If
    
    ValoreIniziale = Predefinito
    
    Combo1_Click
    
    Me.Show vbModal
    
    ApriForm = Risultato
    
End Function

Private Sub cmdEsci_Click()
    Risultato = ValoreIniziale
    Unload Me
End Sub

Private Sub cmdSalva_Click()
    Unload Me
End Sub

Private Sub Combo1_Click()
    If FormCaricata = True Then FermaTimer
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Activate()
    Timer1.Enabled = True
    FormCaricata = True
End Sub

Private Sub Form_Click()
    FermaTimer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    tim = 0
End Sub

Private Sub FermaTimer()

    Risultato = Combo1.List(Combo1.ListIndex)
    Timer1.Enabled = False
    lblTimer.Caption = ""

End Sub

Private Sub lblInfo_Click()
    FermaTimer
End Sub

Private Sub lblMessage_Click()
    FermaTimer
End Sub

Private Sub lblTimer_Click()
    FermaTimer
End Sub

Private Sub Timer1_Timer()
    
    tim = tim + 1
    lblTimer.Caption = totTim - tim
    DoEvents
    DoEvents
    
    If tim >= totTim Then
        tim = 0
        Combo1_KeyPress 13
        Timer1.Enabled = False
    End If
    
End Sub
