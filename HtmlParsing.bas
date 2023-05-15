Attribute VB_Name = "html"
'Info:   this bas file contains code for parsing html documents.
'        parsehtml removes html tags and tries to preserve linebreaks
'        parsescript removes the script blocks and breaks any event code
'        parselinks returns array of link href= values
'        parseImages returns array of img src= values
'
'        some of this coding is rather old so milage may vary :-\

'License:  you are free to use this library in your personal projects, so
'               long as this header remains inplace. This code cannot be
'               used in any project that is to be sold. This source code
'               can be freely distributed so long as this header reamins
'               intact.
'
'Author:   dzzie@yahoo.com
'Sight:    http://www.geocities.com/dzzie


Public Function parseScript(info) As String
  Dim trimpage
  info = filt(info, "javascript,vbscript,mocha,createobject,activex")
  Script = Split(info, "<script")
  
  If AryIsEmpty(Script) Then parseScript = info: Exit Function _
  Else: trimpage = Script(0)
  
  For i = 1 To UBound(Script)
    EndOfScript = InStr(1, Script(i), "</script>")
    trimpage = trimpage & Mid(Script(i), EndOfScript + 10, Len(Script(i)))
  Next
  
  parseScript = CStr(trimpage)
End Function

'bugs if html tag contains quoted < or >
Public Function parseHtml(info) As String
     Dim temp As String, EndOfTag As Integer
     fmat = Replace(info, "&nbsp;", " ")
     cut = Split(fmat, "<")

   For i = 0 To UBound(cut)  'cut at all html start tags
     EndOfTag = InStr(1, cut(i), ">")
        If EndOfTag > 0 Then
          EndOfText = Len(cut(i))
          NL = False
          If Left(cut(i), 2) = "br" Then NL = True
          cut(i) = Mid(cut(i), EndOfTag + 1, EndOfText)
          If NL Then cut(i) = vbCrLf & cut(i)
          If cut(i) = vbCrLf Then cut(i) = ""
        End If
     temp = temp & cut(i)
    Next
    
    parseHtml = temp
End Function

'trims out &amp; type html for text
Public Function parseAnds(info)
  Dim temp As String
  cut = Split(info, "&")
  If UBound(cut) > 0 Then
    For i = 0 To UBound(cut)            'cut at all start tags (&)
      EndOfTag = InStr(1, cut(i), ";")
        If EndOfTag > 0 Then
           EndOfText = Len(cut(i))
           cut(i) = Mid(cut(i), EndOfTag + 1, EndOfText)
        End If
      temp = temp & cut(i)
    Next
   parseAnds = temp
  Else: parseAnds = info
  End If
End Function

Function ParseLinkHTML(dat, ByVal pageUrl, Optional mask = "*", Optional likeit = True) As Variant()
    'return entire link + link text so we can filter by link text
    'also have to make sure all href= values are absolute urls
    
    Dim links()
    tmp = Split(dat, "<a")
    
    For i = 1 To UBound(tmp)
        e = InStr(1, tmp(i), "</a>", 1) + 3
        push links(), "<a" & Mid(tmp(i), 1, e)
    Next
    
    For i = 0 To UBound(links)
        Dim t(), url(), absUrl()
        url() = Extract(links(i), "<a", "href=")
        absUrl() = resolve(url(), pageUrl)
        If url(0) <> absUrl(0) Then
            links(i) = Replace(links(i), url(0), absUrl(0))
        End If
    Next
    
    links() = filter_(links, mask, likeit)
    
    ParseLinkHTML = links()
End Function

Function ParseLinkURLS(dat, ByVal pageUrl, Optional mask = "*", Optional likeit As Boolean = True) As Variant()
    Dim link()
    link() = Extract(dat, "<a", "href=")
    link() = resolve(link(), pageUrl)
    link() = filter_(link(), mask, likeit)
    ParseLinkURLS = link()
End Function

Function ParseImageURLS(dat, ByVal pageUrl, Optional mask = "*", Optional likeit As Boolean = True) As Variant()
    Dim image()
    image() = Extract(dat, "<img", "src=")
    image() = resolve(image(), pageUrl)
    image() = filter_(image(), mask, likeit)
    ParseImageURLS = image()
End Function

Function Extract(data, splitAt, key) As Variant()
    data = Replace(data, splitAt, LCase(splitAt), , , 1)
    splitAt = LCase(splitAt)
    If InStr(1, data, splitAt, 1) < 1 Then Exit Function
    tmp = Split(data, splitAt)
    Dim ret()
    For i = 0 To UBound(tmp)
        st.Strng = tmp(i)
        t = st.IndexOf(key, , True)
        If t > 0 Then
            quotechar = st.GetChar
            Select Case quotechar
                Case """": push ret, st.SubstringToNext("""")
                Case "'": push ret, st.SubstringToNext("'")
                Case Else:
                    g = InStr(t, tmp(i), ">")
                    s = InStr(t, tmp(i), " ")
                    If s < g And s > 0 Then push ret, quotechar & st.SubstringToNext(" ") _
                    Else push ret, quotechar & st.SubstringToNext(">")
            End Select
        End If
    Next
    Extract = ret()
End Function

'make sure they are absolute URL's
Function resolve(ary, pageUrl) As Variant()
    
    If Not AryIndexExists(ary, 0) Then Exit Function
    
    st.Strng = fso.WebParentFolderFromURL(Replace(pageUrl, "http://", ""))
    sf = st.IndexOf("/")
    
    If sf > 2 And sf <> Len(st.Strng) Then
        svr = Mid(st.Strng, 1, sf - 1)
        sf = st.ToEndOfStr
    Else
        svr = Empty
        sf = st.Strng
    End If
    
    For i = 0 To UBound(ary)
        If InStr(ary(i), sf) < 1 And InStr(ary(i), ":") < 1 Then
            ary(i) = "http://" & svr & "/" & sf & ary(i)
            ary(i) = Replace(ary(i), "/./", "/")
        End If
    Next
    
    resolve = ary
End Function

Function filter_(ary, sFilter, Optional likeit = True) As Variant()
    If Not AryIndexExists(ary, 0) Then Exit Function

    Dim ret()
    For i = 0 To UBound(ary)
        If likeit Then
            If LCase(ary(i)) Like LCase(sFilter) Then push ret, ary(i)
        Else
            If LCase(ary(i)) Like LCase(sFilter) Then: Else push ret, ary(i)
        End If
    Next
    
    filter_ = ret()
End Function

Private Function filt(txt, remove As String)
  If Right(txt, 1) = "," Then txt = Mid(txt, 1, Len(txt) - 1)
  tmp = Split(remove, ",")
  For i = 0 To UBound(tmp)
     txt = Replace(txt, tmp(i), "", , , vbTextCompare)
  Next
  filt = txt
End Function
