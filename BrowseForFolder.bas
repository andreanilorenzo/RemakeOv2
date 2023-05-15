Attribute VB_Name = "BrowseForFolder"
Option Explicit
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const MAX_PATH = 260
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Function BrosweForFolder(NameForm As Form, Title As String) As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim path As String
    Dim pos As Integer
    bi.hOwner = NameForm.hwnd
    bi.pidlRoot = 0&
    bi.lpszTitle = Title
    bi.ulFlags = BIF_RETURNONLYFSDIRS
    pidl = SHBrowseForFolder(bi)
    path = Space$(MAX_PATH)
    If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
    pos = InStr(path, Chr$(0))
    path = Left(path, pos - 1)
    If Right(path, 1) = "\" Then
    BrosweForFolder = Left(path, pos - 1)
    Else
    BrosweForFolder = Left(path, pos - 1) & "\"
    End If
    End If
    Call CoTaskMemFree(pidl)
End Function
