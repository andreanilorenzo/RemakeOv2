Attribute VB_Name = "mDirToHtml"
Option Explicit

Public Sub ExtractAll(ByVal SiteRoot As String, ByVal FilePath As String)
    Dim i, t As Integer
    Dim a As String
    Dim x As Variant
    Dim strFileExtension
    Dim rGetIcon As Long
    Dim MyName As String
    Dim SubDir(500) As String

    If Mid$(FilePath, Len(FilePath), 1) <> "\" Then
        FilePath = FilePath + "\"
    End If
    
    i = 0
    
    On Error GoTo nodir
    
    MyName = Dir(FilePath, vbDirectory)   ' Retrieve the first entry.
    Do While MyName <> ""   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If MyName <> "." And MyName <> ".." And MyName <> "Directories" Then
          ' Use bitwise comparison to make sure MyName is a directory.
          If (GetAttr(FilePath & MyName) And vbDirectory) = vbDirectory Then
             SubDir(i) = MyName ' Display entry only if it
             i = i + 1
          End If   ' it represents a directory.
       End If
       MyName = Dir   ' Get next entry.
    Loop
    Open FilePath & "index.html" For Output As #1
        Print #1, "<html><head><Title>" + Replace(Mid$(FilePath, Len(SiteRoot) + 1), "\", "/") + "</Title>"
        Print #1, "<link REL='STYLESHEET' TYPE='text/css' HREF='Directories.css'>"
        Print #1, "</head><body>"
        Print #1, "<h1 class='FolderName'>" + Mid$(FilePath, Len(SiteRoot) + 2) + "</h1>"
        For t = 0 To i - 1
            Print #1, "<p class='SubForlder'><img border='0' src='folder.bmp'><a href='" + Replace(Mid$(FilePath & SubDir(t) + "\", Len(SiteRoot) + 2), "\", ".") + "index.html'>  " + Mid$(FilePath & SubDir(t), Len(SiteRoot) + 2) + "</a></p>"
        Next t
        MyName = Dir(FilePath, vbNormal)   ' Retrieve the first entry.
        DoEvents
        Do While MyName <> ""
            'addIcon MyName, SiteRoot + "\Directories\"
            
            a = MyName
            x = Split(a, ".")
            strFileExtension = x(UBound(x))
            
            'rGetIcon = ExtractAssociatedIcon(0, FilePath + MyName, 1)
            'Set picTemp.Picture = Nothing
            'DrawIcon picTemp.hdc, 0, 0, rGetIcon
            'picTemp.Picture = picTemp.Image
            If strFileExtension = "exe" Then
                'SavePicture picTemp.Picture, SiteRoot + "\Directories\" + LCase(MyName) + ".bmp"
                Print #1, "<p class='FileExe'><img border='0' src='" + LCase(MyName) + ".bmp" + "'><a href='../" + Replace(Mid$(FilePath + MyName, Len(SiteRoot) + 2), "\", "/") + "'>  " + MyName + "</a></p>"
            Else
                'SavePicture picTemp.Picture, SiteRoot + "\Directories\" + LCase(strFileExtension) + ".bmp"
                Print #1, "<p class='File'><img border='0' src='" + LCase(strFileExtension) + ".bmp" + "'><a href='../" + Replace(Mid$(FilePath + MyName, Len(SiteRoot) + 2), "\", "/") + "'>  " + MyName + "</a></p>"
            End If
            
            DoEvents
            MyName = Dir   ' Get next entry.
        Loop
        Print #1, "</body></html>"
    Close #1
nodir:
    For t = 0 To i - 1
        ExtractAll SiteRoot, FilePath & SubDir(t) + "\"
    Next t
    
    Exit Sub
End Sub

Private Function ClearFileName(ByVal FileName As String) As String
    FileName = Replace(FileName, "\", ".")
    FileName = Replace(FileName, "/", ".")
    FileName = Replace(FileName, "#", "")
    FileName = Replace(FileName, "*", "")
    FileName = Replace(FileName, "?", "")
    FileName = Replace(FileName, ":", "")
    ClearFileName = FileName
End Function



