VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Cmdlg Class Module by Wpsjr1@syix.com   updated June 09, 1999

' initialize the UDT

Public Enum Flag
   OFN_ALLOWMULTISELECT = &H200
   OFN_CREATEPROMPT = &H2000
   OFN_ENABLEHOOK = &H20
   OFN_ENABLETEMPLATE = &H40
   OFN_ENABLETEMPLATEHANDLE = &H80
   OFN_EXPLORER = &H80000
   OFN_EXTENSIONDIFFERENT = &H400
   OFN_FILEMUSTEXIST = &H1000
   OFN_HIDEREADONLY = &H4
   OFN_LONGNAMES = &H200000
   OFN_NOCHANGEDIR = &H8
   OFN_NODEREFERENCELINKS = &H100000
   OFN_NOLONGNAMES = &H40000
   OFN_NONETWORKBUTTON = &H20000
   OFN_NOREADONLYRETURN = &H8000
   OFN_NOTESTFILECREATE = &H10000
   OFN_NOVALIDATE = &H100
   OFN_OVERWRITEPROMPT = &H2
   OFN_PATHMUSTEXIST = &H800
   OFN_READONLY = &H1
   OFN_SHAREAWARE = &H4000
   OFN_SHAREFALLTHROUGH = 2
   OFN_SHAREWARN = 0
   OFN_SHARENOWARN = 1
   OFN_SHOWHELP = &H10
   OFS_MAXPATHNAME = 128
End Enum

Public Enum StandardFlag
  OpenFile = Flag.OFN_EXPLORER Or Flag.OFN_LONGNAMES Or Flag.OFN_CREATEPROMPT Or Flag.OFN_NODEREFERENCELINKS
  OpenMulti = StandardFlag.OpenFile Or Flag.OFN_ALLOWMULTISELECT
  SaveFile = Flag.OFN_EXPLORER Or Flag.OFN_LONGNAMES Or Flag.OFN_OVERWRITEPROMPT Or Flag.OFN_HIDEREADONLY
End Enum
  

Private Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
 End Type
 
Private Enum CmdlgErrors
  CDERR_GENERALCODES = &H0
  CDERR_STRUCTSIZE = &H1
  CDERR_INITIALIZATION = &H2
  CDERR_NOTEMPLATE = &H3
  CDERR_NOHINSTANCE = &H4
  CDERR_LOADSTRFAILURE = &H5
  CDERR_FINDRESFAILURE = &H6
  CDERR_LOADRESFAILURE = &H7
  CDERR_LOCKRESFAILURE = &H8
  CDERR_MEMALLOCFAILURE = &H9
  CDERR_MEMLOCKFAILURE = &HA
  CDERR_NOHOOK = &HB
  CDERR_REGISTERMSGFAIL = &HC

  FNERR_FILENAMECODES = &H3000
  FNERR_SUBCLASSFAILURE = &H3001
  FNERR_INVALIDFILENAME = &H3002
  FNERR_BUFFERTOOSMALL = &H3003
  FRERR_FINDREPLACECODES = &H4000
  FRERR_BUFFERLENGTHZERO = &H4001
  CDERR_DIALOGFAILURE = &HFFFF&
End Enum
 
Private OFN As OPENFILENAME

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
   
' the next two declares are unused but could be used to add other functions
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' make private variables used within the Class
Private m_bShowSave As Boolean
Private m_bCancelError As Boolean
Private m_sDefaultExtension As String
Private m_sDialogTitle As String
Private m_sFilename As String
Private m_sFiletitle As String
Private m_iFilterIndex As Integer
Private m_sInitDir As String
Private m_sFilter As String
Private m_bCancelPressed As Boolean
Private m_sFileExtension As String
Private m_lFlags As Long
Private m_bMultiSelect As Boolean
Private m_iMaxFileSize As Integer
Private m_lhWnd As Long
'

Public Sub ShowOpen()
  Dim ret&

  Call Initialize               ' fill-in the OFN structure
  ret = GetOpenFileName(OFN)    ' open the dialog and see what we get back
  Call HandleReturn(ret)        ' process the return value

End Sub

Public Sub ShowSave()
  Dim ret&

  m_bShowSave = True
  Call Initialize
  ret = GetSaveFileName(OFN)
  Call HandleReturn(ret)

End Sub

Private Sub Initialize()
  OFN.nStructSize = Len(OFN)
  OFN.hwndOwner = m_lhWnd ' will not be modal if hWnd is not specified

  OFN.sFilter = m_sFilter
  OFN.nFilterIndex = m_iFilterIndex

  If m_sFilename = "" Then       ' make a buffer depending on input
    OFN.sFile = vbNullChar & String$(m_iMaxFileSize, 0)
  Else
    OFN.sFile = m_sFilename & String$(m_iMaxFileSize, 0)
  End If

  OFN.nFileSize = Len(OFN.sFile)

  ' only set it if a default ext is provided
  If m_sDefaultExtension <> vbNullString Then OFN.sDefFileExt = m_sDefaultExtension

  OFN.sFileTitle = String$(m_iMaxFileSize / 2, 0)
  OFN.nTitleSize = Len(OFN.sFileTitle)
  OFN.sInitDir = m_sInitDir
  OFN.sDlgTitle = m_sDialogTitle
  OFN.Flags = m_lFlags
End Sub

Private Sub HandleReturn(lReturn As Long)
  Dim lError&

  If lReturn > 0 Then  ' 0 is returned when cancel is pressed
    
    If (OFN_ALLOWMULTISELECT And OFN.Flags) = OFN_ALLOWMULTISELECT Then
      Let FileName = OFN.sFile
    Else
      Let FileName = TrimNull(OFN.sFile)
    End If
    
    Let Filetitle = TrimNull(OFN.sFileTitle)

    If m_bShowSave = True Then m_bShowSave = False
    
    ' the FileExtension property is read-only so the variable must be set
    If Asc(Mid$(OFN.sFile, (InStr(OFN.sFile, vbNullChar)) + 1, 1)) = 0 And (OFN_ALLOWMULTISELECT = OFN_ALLOWMULTISELECT And OFN.Flags) Then   ' test for MultiSelect or if ShowSave was chosen
      m_sFileExtension = Mid$(OFN.sFile, OFN.nFileExt + 1, Len(OFN.sFile) - OFN.nFileExt)
      m_bMultiSelect = False
    Else
      m_bMultiSelect = True
      m_sFileExtension = ""
    End If
    If Not OFN_ALLOWMULTISELECT = (OFN_ALLOWMULTISELECT And OFN.Flags) Then m_bMultiSelect = False
    
    m_bCancelPressed = False         ' read-only also
  Else
    lError = CommDlgExtendedError()
    
    If lError > 0 Then HandleExtendedError (lError)
    
    Let FileName = ""
    Let Filetitle = ""
    m_sFileExtension = ""
    m_bCancelPressed = True
    m_bMultiSelect = False
  End If
  
End Sub

Private Sub HandleExtendedError(cdlgError As Long)
   Select Case cdlgError
     'Case CDERR_GENERALCODES: Err.Raise , , "CDERR_GENERALCODES"
     'Case CDERR_STRUCTSIZE: Err.Raise , , "CDERR_STRUCTSIZE"
     'Case CDERR_INITIALIZATION: Err.Raise , , "CDERR_INITIALIZATION"
     'Case CDERR_NOTEMPLATE: Err.Raise , , "CDERR_NOTEMPLATE"
     'Case CDERR_NOHINSTANCE: Err.Raise , , "CDERR_NOHINSTANCE"
     'Case CDERR_LOADSTRFAILURE: Err.Raise , , "CDERR_LOADSTRFAILURE"
     'Case CDERR_FINDRESFAILURE: Err.Raise , , "CDERR_FINDRESFAILURE"
     'Case CDERR_LOADRESFAILURE: Err.Raise , , "CDERR_LOADRESFAILURE"
     'Case CDERR_LOCKRESFAILURE: Err.Raise , , "CDERR_LOCKRESFAILURE"
     'Case CDERR_MEMALLOCFAILURE: Err.Raise , , "CDERR_MEMALLOCFAILURE"
     'Case CDERR_MEMLOCKFAILURE: Err.Raise , , "CDERR_MEMLOCKFAILURE"
     'Case CDERR_NOHOOK: Err.Raise , , "CDERR_NOHOOK"
     'Case CDERR_REGISTERMSGFAIL: Err.Raise , , "CDERR_REGISTERMSGFAIL"
     'Case FNERR_FILENAMECODES: Err.Raise , , "FNERR_FILENAMECODES"
     'Case FNERR_SUBCLASSFAILURE: Err.Raise , , "FNERR_SUBCLASSFAILURE"
     'Case FNERR_INVALIDFILENAME: Err.Raise , , "FNERR_INVALIDFILENAME"
     ' the first two bytes of OFN.sFile could also be passed along
     ' some sort of retry routine could be fashioned
     'Case FNERR_BUFFERTOOSMALL: Err.Raise , , "FNERR_BUFFERTOOSMALL - Increase MaxFileSize to: " & CStr(Asc(Left$(OFN.sFile, 1))) & CStr(Asc(Mid$(OFN.sFile, 2, 1)))
     'Case FRERR_FINDREPLACECODES: Err.Raise , , "FRERR_FINDREPLACECODES"
     'Case FRERR_BUFFERLENGTHZERO: Err.Raise , , "FRERR_BUFFERLENGTHZERO"
     'Case CDERR_DIALOGFAILURE: Err.Raise , , "CDERR_DIALOGFAILURE"
     'Case Else: Err.Raise 666, , "Unknown CommonDialog Error"
   End Select
End Sub

Function TrimNull(s As String) As String
  Dim lWhere&
  lWhere = InStr(1, s, vbNullChar)
  
  If lWhere > 0 Then
    TrimNull = Left$(s, lWhere - 1)
  Else
    TrimNull = s
    'Debug.Print s & " " & "no null present"
  End If
End Function


Public Property Let InitDir(ByVal sInitDir As String)

  m_sInitDir = sInitDir
    
End Property

Public Property Get InitDir() As String

  InitDir = m_sInitDir
  
End Property

Public Property Let FilterIndex(ByVal iFilterIndex As Integer)
 
 If iFilterIndex = 0 Then iFilterIndex = 1
 m_iFilterIndex = iFilterIndex
 
End Property

Public Property Get FilterIndex() As Integer

  FilterIndex = m_iFilterIndex
  
End Property

Public Property Let Filetitle(ByVal sFileTitle As String)

  m_sFiletitle = sFileTitle
    
End Property

Public Property Get Filetitle() As String
    
  Filetitle = m_sFiletitle
  
End Property

Public Property Let FileName(ByVal sFilename As String)
    
  m_sFilename = sFilename
  
End Property

Public Property Get FileName() As String

  FileName = m_sFilename
    
End Property

Public Property Let DialogTitle(ByVal sDialogTitle As String)

  m_sDialogTitle = sDialogTitle
  
End Property

Public Property Get DialogTitle() As String

  DialogTitle = m_sDialogTitle
  
End Property

Public Property Let DefaultExtension(ByVal sDefaultExtension As String)

  m_sDefaultExtension = sDefaultExtension
    
End Property

Public Property Get DefaultExtension() As String

  DefaultExtension = m_sDefaultExtension
  
End Property

Public Property Get CancelPressed() As Boolean      ' read-only

  CancelPressed = m_bCancelPressed
  
End Property

Public Property Get FileExtension() As String       ' read-only

  FileExtension = m_sFileExtension
  
End Property

Public Property Get MultiSelect() As Boolean ' read only

  MultiSelect = m_bMultiSelect
 
End Property

Public Property Let Filter(ByVal sFilter As String) ' write-only

  m_sFilter = sFilter
  
End Property

Public Property Let Flags(ByVal lFlags As Long)     ' write only

  m_lFlags = lFlags
  
End Property

Public Property Let hWnd(ByVal lhWnd As Long)
  
  m_lhWnd = lhWnd
  
End Property

Public Property Let MaxFileSize(ByVal iMaxFileSize As Integer)

  If iMaxFileSize < 1024 Then iMaxFileSize = 1024   ' set the minimum
  m_iMaxFileSize = iMaxFileSize

End Property

' prepare the class with the minimun settings (in case nothing was set)
Private Sub Class_Initialize()
  m_bMultiSelect = False
  m_bShowSave = False
  MaxFileSize = 2048
  InitDir = ""
  FilterIndex = 1
  Filetitle = ""
  FileName = ""
  DialogTitle = App.Path
  DefaultExtension = ""
  Filter = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
End Sub
