Attribute VB_Name = "fso"
'Revision 3 <- Incompatiable with all previous..simplified & streamlined
'
'Info:     These are basically macros for VB's built in file processes
'            this should streamline your code quite a bit and hopefully
'            remove alot of redundant coding.
'
'License:  you are free to use this library in your personal projects, so
'               long as this header remains inplace. This code cannot be
'               used in any project that is to be sold. This source code
'               can be freely distributed so long as this header reamins
'               intact.
'
'Author:   dzzie@yahoo.com
'Sight:    http://www.geocities.com/dzzie
    
Function GetFolderFiles(folder, Optional filter = ".*", Optional retFullPath As Boolean = True) As String()
   Dim fnames() As String
   
   If Not FolderExists(folder) Then
        'returns empty array if fails
        GetFolderFiles = fnames()
        Exit Function
   End If
   
   folder = IIf(Right(folder, 1) = "\", folder, folder & "\")
   If Left(filter, 1) = "*" Then extension = Mid(filter, 2, Len(filter))
   If Left(filter, 1) <> "." Then filter = "." & filter
   
   fs = Dir(folder & "*" & filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   While fs <> ""
     If fs <> "" Then push fnames(), IIf(retFullPath = True, folder & fs, fs)
     fs = Dir()
   Wend
   
   GetFolderFiles = fnames()
End Function

Function GetSubFolders(folder, Optional retFullPath As Boolean = True) As String()
    Dim fnames() As String
    
    If Not FolderExists(folder) Then
        'returns empty array if fails
        GetSubFolders = fnames()
        Exit Function
    End If
    
   If Right(folder, 1) <> "\" Then folder = folder & "\"

   fd = Dir(folder, vbDirectory)
   While fd <> ""
     If Left(fd, 1) <> "." Then
        If (GetAttr(folder & fd) And vbDirectory) = vbDirectory Then
           push fnames(), IIf(retFullPath = True, folder & fd, fd)
        End If
     End If
     fd = Dir()
   Wend
   
   GetSubFolders = fnames()
End Function

Function FolderExists(path) As Boolean
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Function FileExistsA(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExistsA = True _
  Else FileExistsA = False
End Function

Function GetParentFolder(path) As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function

Sub CreateFolderA(path)
   If FolderExists(path) Then Exit Sub
   MkDir path
End Sub

Function FileNameFromPathA(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPathA = CStr(tmp(UBound(tmp)))
    End If
End Function

Function WebFileNameFromPath(fullpath)
    If InStr(fullpath, "/") > 0 Then
        tmp = Split(fullpath, "/")
        WebFileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Sub Move(fpath, toFolder)
    Copy fpath, toFolder
    Kill fpath
End Sub

Sub DeleteFile(fpath)
    If FileExistsA(fpath) Then
        SetAttr fpath, vbNormal
        Kill fpath
    End If
End Sub

Sub Rename(fullpath, newName)
    On Error Resume Next
  pf = fso.GetParentFolder(fullpath)
  Name fullpath As pf & "\" & newName
End Sub

Function RimuoviExt(path) As String
    RimuoviExt = Left$(path, InStrRev(path, ".") - 1)
End Function

Sub SetAttribute(fpath, it As VbFileAttribute)
   SetAttr fpath, it
End Sub

Function GetExtension(path) As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetExtension = Mid(ub, InStrRev(ub, "."), Len(ub))
    Else
       GetExtension = ""
    End If
End Function

Function GetBaseName(path) As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(1, ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

Function ChangeExt(path, ext)
    ext = IIf(Left(ext, 1) = ".", ext, "." & ext)
    If fso.FileExistsA(path) Then
        fso.Rename path, fso.GetBaseName(path) & ext
    Else
        'hack to just accept a file name might not be worth supporting
        bn = Mid(path, 1, InStr(1, path, ".") - 1)
        ChangeExt = bn & ext
    End If
End Function

Function SafeFileName(proposed) As String
  badChars = ">,<,&,/,\,:,|,?,*,"""
  bad = Split(badChars, ",")
  For i = 0 To UBound(bad)
    proposed = Replace(proposed, bad(i), "")
  Next
  If proposed = Empty Then proposed = RandomNum()
  SafeFileName = CStr(proposed)
End Function

Function RandomNum()
    Randomize
    tmp = Round(Timer * Now * Rnd(), 0)
    RandomNum = tmp
End Function

Function GetFreeFileName(folder, Optional extension = ".txt") As String
    
    If Not fso.FolderExists(folder) Then Exit Function
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Left(extension, 1) <> "." Then extension = "." & extension
    
    Dim tmp As String
    Do
      tmp = folder & RandomNum() & extension
    Loop Until Not fso.FileExistsA(tmp)
    
    GetFreeFileName = tmp
End Function

Function buildPath(folderpath) As Boolean
    On Error GoTo oops
    
    If FolderExists(folderpath) Then buildPath = True: Exit Function
    
    tmp = Split(folderpath, "\")
    build = tmp(0)
    For i = 1 To UBound(tmp)
        build = build & "\" & tmp(i)
        If InStr(tmp(i), ".") < 1 Then
            If Not FolderExists(build) Then CreateFolderA (build)
        End If
    Next
    buildPath = True
    Exit Function
oops: buildPath = False
End Function


Function ReadFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Sub WriteFile(path, it)
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Sub AppendFile(path, it)
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub


Sub Copy(fpath, toFolder)
   If FolderExists(toFolder) Then
       baseName = fso.FileNameFromPathA(fpath)
       toFolder = IIf(Right(toFolder, 1) = "\", toFolder, toFolder & "\")
       FileCopy fpath, toFolder & baseName
   Else 'assume tofolder is actually new desired file path
       FileCopy fpath, toFolder
   End If
End Sub

Sub CreateFile(fpath)
    f = FreeFile
    If fso.FileExistsA(fpath) Then Exit Sub
    Open fpath For Binary As f
    Close f
End Sub


Function DeleteFolder(folderpath, Optional force As Boolean = True) As Boolean
 On Error GoTo failed
   Call delTree(folderpath, force)
   Call RmDir(folderpath)
   DeleteFolder = True
 Exit Function
failed:  DeleteFolder = False
End Function

Private Sub delTree(folderpath, Optional force As Boolean = True)
   Dim sfi() As String, sfo() As String
   sfi() = fso.GetFolderFiles(folderpath)
   sfo() = fso.GetSubFolders(folderpath)
   If Not AryIsEmpty(sfi) And force = True Then
        For i = 0 To UBound(sfi)
            Kill sfi(i)
        Next
   End If
   
   If Not AryIsEmpty(sfo) And force = True Then
        For i = 0 To UBound(sfo)
            Call DeleteFolder(sfo(i), True)
        Next
   End If
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    X = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function WebParentFolderFromURL(url) As String
    If url = Empty Or InStr(url, "/") < 1 Then Exit Function
    tmp = Split(url, "/")
    If InStr(tmp(UBound(tmp)), ".") > 0 Then tmp(UBound(tmp)) = Empty
    tmp = Join(tmp, "/")
    If Right(tmp, 2) = "//" Then tmp = Mid(tmp, 1, Len(tmp) - 1)
    WebParentFolderFromURL = CStr(tmp)
End Function
