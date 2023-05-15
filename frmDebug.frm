VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDebug 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    
    Dim bordo As Integer
    
    bordo = 50
    
    txtDebug.Move Me.ScaleLeft + bordo, Me.ScaleTop + bordo, Me.ScaleWidth + bordo, Me.ScaleHeight + bordo

End Sub
