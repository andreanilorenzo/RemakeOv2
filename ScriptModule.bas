Attribute VB_Name = "ScriptModule"
Option Explicit

Public Type tyScript ' Il tipo dati utilizzato per la lettura dei valori dello script
    oOK As Boolean       ' Se false i dati non sono aggiornati
    oNAME As String
    oWEB As String
    oURL As String
    oREPLACE As String
    iURL As String      ' La pagina web con i valori modificati
    iIndirizzo As String
    iCitta As String
    iCap As String
    iLat As String
    iLong As String
    iWebBrowser As WebBrowser
    iInStrRetURL As String
End Type

Public ValScript As tyScript

Global CACHE

Public Function LeggiScript(PathScript As String)
    Dim lines
    Dim i As Integer
    Dim testoFile As String

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    CACHE = Empty ' Line of script had an error in syntax
    
    testoFile = fso.ReadFile(PathScript)
    If testoFile = "" Then GoTo Esci
    lines = Split(testoFile, vbCrLf)
    
    For i = 0 To UBound(lines)
        If lines(i) <> Empty And (Left$(lines(i), 1) <> "#") Then
            ParseCommandLine lines(i)
        End If
    Next
    
Esci:
    Exit Function
    
Errore:
    Screen.MousePointer = vbDefault
    GestErr Err, "Errore nella funzione LeggiScript alla riga " & i & " del file " & PathScript

End Function

Private Sub ParseCommandLine(commandLine)
    Dim cmd()
    Dim strTmp As String
    
    If commandLine = Empty Or commandLine = "" Then Exit Sub
    
    cmd() = GetArgs(commandLine)
    
    Select Case cmd(0)
        Case "NAME"
            ValScript.oNAME = cmd(1)
            
        Case "WEB"
            ValScript.oWEB = cmd(1)
            
        Case "REPLACE"
            ValScript.oREPLACE = cmd(1)
            Call scrReplace(cmd)
        
        Case "URL"
            strTmp = cmd(1)
            ValScript.oURL = strTmp
            Call SostituisciCampi(strTmp, "DAT")
            ValScript.iURL = strTmp
            
        Case "INSTR"
            If cmd(1) = "RETURL" Then
                ValScript.iInStrRetURL = cmd(2)
            End If
            
        
        'Case "GET": Call ParseGet(cmd)
        'Case "DOWNLOAD": Call ParseDownload(cmd)
        'Case "LOAD": Call LoadFile(cmd)
        'Case "EXECUTE": Call Execute(cmd)
        'Case "PAUSE": Call Sleep(CLng(cmd(1)))
        'Case "SEARCH": Call search(cmd)
        'Case "STORE": Call Store(cmd)
        'Case "APPEND": Call append(cmd)
        'Case "SHOW": Call Showit(cmd)
        'Case "CLEAR": CACHE = Empty: SCRIPTCACHE = Empty: ReDim VARIABLES(0)
        'Case "FILTER": Call filter(cmd)
        'Case "SAVE": Call Save(cmd)
        
        Case "NAVIGATE"
            ValScript.oOK = True
            
    End Select
        
End Sub

Function GetArgs(cmd) As Variant()
    ' this function is used for grabbing the command line info passed in from a text string. it will recgonize
    ' double and single quoted arguments with spaces as one argument it will also recgonize " -ac" as two switchs
    If cmd = Empty Then Exit Function
    
    Dim args()
    Dim inquotes As Boolean
    Dim inminus As Boolean
    Dim isword As Boolean
    Dim tmp, lastLetter, letter, nextlet
    Dim i As Long
    
    tmp = ""
    lastLetter = ""
    cmd = REPLACE(cmd, """", "'")
    
    For i = 1 To Len(cmd)
      letter = Mid(cmd, i, 1)
      nextlet = Mid(cmd, i + 1, 1)
      
      Select Case letter
        Case "-":
                  If lastLetter = " " And Not inquotes Then
                        inminus = True: isword = False
                  End If
        Case " ":
                    inminus = False
                    If isword Then
                      isword = False
                      push args, tmp
                      tmp = ""
                    End If
        Case "'":
                   isword = False
                   If inquotes = True Then
                    inquotes = False
                    push args, tmp
                    tmp = ""
                   Else
                     inquotes = True
                   End If
      End Select
      
      If inminus And Not inquotes And letter <> "-" Then
         push args, letter
      ElseIf inquotes And letter <> "'" Then
         tmp = tmp & letter
      ElseIf Not inminus And Not inquotes And letter <> "'" Then
         isword = True
         tmp = tmp & letter
         If i = Len(cmd) Then push args, tmp
      End If
      lastLetter = letter
    Next
    
    If AryIsEmpty(args) Then Exit Function
    
    For i = 0 To UBound(args)
        args(i) = Trim(LTrim(args(i)))
    Next
    
    If args(UBound(args)) = "'" Or args(UBound(args)) = Empty Then pop args
    
    GetArgs = args()
End Function

Private Sub scrReplace(ByVal ary)
    'REPLACE (<INDIRIZZO> | <CITTA> | <CAP>) A WHITH B
    '
    Dim searchStr, nextStr
    Dim strTmp As String

    If Var(GestioneErrori).Valore = 0 Then On Error GoTo Errore

    Call SostituisciCampi(ary(2))
    Call SostituisciCampi(ary(4))

    Select Case ary(1)
        Case "<INDIRIZZO>"
            ValScript.iIndirizzo = REPLACE(ValScript.iIndirizzo, ary(2), ary(4))
            
        Case "<CITTA>"
            ValScript.iCitta = REPLACE(ValScript.iCitta, ary(2), ary(4))
            
        Case "<CAP>"
            ValScript.iCap = REPLACE(ValScript.iCap, ary(2), ary(4))
            
    End Select
            
    Exit Sub

Errore:

End Sub

Public Function SostituisciCampi(ByRef Testo, Optional Tipo As String = "CHR")
    Dim strTmp As String
    
    strTmp = Testo
    
    If Tipo = "CHR" Then
        strTmp = REPLACE(strTmp, "<SPACE>", " ")
        
    ElseIf Tipo = "DAT" Then
        strTmp = REPLACE(strTmp, "<INDIRIZZO>", ValScript.iIndirizzo)
        strTmp = REPLACE(strTmp, "<CITTA>", ValScript.iCitta)
        strTmp = REPLACE(strTmp, "<CAP>", ValScript.iCap)
    End If

    Testo = strTmp
    
End Function
