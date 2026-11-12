' Parity script: the expected output (ComParity_Msxml2ParseError.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' DOMDocument.parseError (issue #173): errorCode, reason, line, linepos, filepos, srcText and url of the most recent loadXML - for unclosed tags, no
' or several top level elements, a mismatched end tag, a duplicate attribute, an undefined entity, an unquoted value - and after a successful load
Dim x : Set x = CreateObject("MSXML2.DOMDocument") : x.async = False
Dim c, e
For Each c In Array("<form><parameters>", "", "<a></b>", "<a b='1' b='2'/>", "<a>" & vbCrLf & "  <b>x</c>" & vbCrLf & "</a>", "text", "<a/><b/>", "<a>&nope;</a>", "<a x=1/>", "<?xml version=""1.0""?>")
    Dim ok : ok = x.loadXML(c)
    Set e = x.parseError
    Out ok & "|" & e.errorCode & "|" & Hex(e.errorCode) & "|" & Replace(Replace(e.reason, vbCr, "\r"), vbLf, "\n") & "|" & e.line & "|" & e.linepos & "|" & e.filepos & "|" & e.srcText & "|" & e.url & "|" & TypeName(e)
Next
x.loadXML "<a/>"
Set e = x.parseError
Out e.errorCode & "|" & e.reason & "|" & e.line & "|" & e.linepos & "|" & e.filepos & "|" & e.srcText & "|" & e.url & "|" & TypeName(e.reason)
Dim x2 : Set x2 = CreateObject("MSXML2.DOMDocument")
x2.async = False
Dim ok3 : ok3 = x2.loadXML("<form><parameters>")
Dim e2 : Set e2 = x2.parseError
Out ok3 & "|" & (e2.errorCode <> 0) & "|" & (Len(e2.reason) > 0) & "|" & e2.line & "|" & (e2.linepos > 0) & "|" & e2.srcText & "|" & x2.parseError.errorCode
x2.loadXML "<a/>"
Out x2.parseError.errorCode & "|" & x2.parseError.reason & "|"
