VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPhoneFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mOutputs As String                  ' Variable for the ReadOnly Outputs property
Private mPhoneFile As String                ' Il nome o la posizione del programma
Private mNomeTelefono As String             ' Il nome del telefono
Private WithEvents objDOS As DOSOutputs
Attribute objDOS.VB_VarHelpID = -1

' Eventi:
Public Event RispostaPhoneFile(CommandOutputs As String)

Public Function LeggiDir(ByVal pDirectoryCell As String) As String
    Dim sComando As String
    Dim sComandLine As String
    
    sComando = " -dir "
    sComandLine = mPhoneFile & " " & mNomeTelefono & sComando & pDirectoryCell
    ' Invio il camndo all'eseguibile
    mOutputs = objDOS.ExecuteCommand(, sComandLine, "PhoneFile")

    ' Invio i dati all'evento RispostaPhoneFile
    RaiseEvent RispostaPhoneFile(mOutputs)
    
    LeggiDir = mOutputs
    
End Function

Public Function ScriviFile(ByVal sFileTel As String, ByVal sFilePc As String) As String
    Dim sComando As String
    Dim sComandLine As String
    
    sComando = " -write "
    sComandLine = mPhoneFile & " " & mNomeTelefono & sComando & " " & sFileTel & " " & sFilePc
    ' Invio il camndo all'eseguibile
    mOutputs = objDOS.ExecuteCommand(, sComandLine, "PhoneFile")

    ' Invio i dati all'evento RispostaPhoneFile
    RaiseEvent RispostaPhoneFile(mOutputs)
    
    ScriviFile = mOutputs
    
End Function

Public Function GetPhoneFileInfo(Optional bTutto As Boolean = True) As String
    
    mOutputs = objDOS.ExecuteCommand(, mPhoneFile, "PhoneFile")
    
    If bTutto = False Then
        ' Tolgo i dati di troppo dalla stringa di ritorno
        mOutputs = mOutputs
    End If
    
    ' Invio i dati all'evento RispostaPhoneFile
    RaiseEvent RispostaPhoneFile(mOutputs)
    
    GetPhoneFileInfo = mOutputs
    
End Function

Public Function InviaCmdLine(ByVal sComando As String) As String
    ' Invia una riga di comando al pprogramma
    
    mOutputs = objDOS.ExecuteCommand(, sComando, "PhoneFile")

    ' Invio i dati all'evento RispostaPhoneFile
    RaiseEvent RispostaPhoneFile(mOutputs)
    
    InviaCmdLine = mOutputs
    
End Function

Private Function VerificaNomeTelefono(ByVal sNomeTelefono As String) As String
    ' Se il nome del telefono contiene degli spazi aggiungo le virgolette
    Dim tmpNome As String
    
    If sNomeTelefono = "" Then
        VerificaNomeTelefono = ""
        Exit Function
    End If
    
    tmpNome = sNomeTelefono
    
    If InStr(1, tmpNome, " ", vbTextCompare) > 0 Then
        If Left$(tmpNome, 1) <> Chr(34) Then
            tmpNome = Chr(34) & tmpNome
        End If
        If Right$(tmpNome, 1) <> Chr(34) Then
            tmpNome = tmpNome & Chr(34)
        End If
    End If
    
    VerificaNomeTelefono = tmpNome
    
End Function


Private Sub Class_Initialize()
    Set objDOS = New DOSOutputs
    mPhoneFile = "PhoneFile"
End Sub

Private Sub Class_Terminate()
    Set objDOS = Nothing
End Sub


Public Property Let PhoneFile(NomePhoneFileApp As String)
    ' Imposta il nome oppure il percorso completo del programma PhoneFile
    mPhoneFile = NomePhoneFileApp
End Property

Public Property Get PhoneFile() As String
    PhoneFile = mPhoneFile
End Property

Public Property Let NomeTelefono(sNomeTelefono As String)
    ' Imposta il nome oppure il percorso completo del programma PhoneFile
    mNomeTelefono = VerificaNomeTelefono(sNomeTelefono)
End Property

Public Property Get NomeTelefono() As String
    NomeTelefono = mNomeTelefono
End Property


