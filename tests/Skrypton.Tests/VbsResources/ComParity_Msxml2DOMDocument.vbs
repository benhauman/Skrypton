' Parity script: the expected output (ComParity_Msxml2DOMDocument.expected.txt) was captured with Windows cscript.exe and the real COM components.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
Dim doc : Set doc = CreateObject("Msxml2.DOMDocument")
Out TypeName(doc) & "|" & VarType(doc)
doc.async = False
doc.setProperty "SelectionLanguage", "XPath"
Out doc.loadXML("<root><item id=""1"">one</item><item id=""2"">two</item></root>")
Out doc.selectSingleNode("//item[@id='2']").text
Out doc.selectNodes("//item").length
Out doc.documentElement.nodeName
Out doc.xml
Out doc.loadXML("<root><bad></root>")
Out TypeName(doc.selectSingleNode("//nothing"))
Out doc.loadXML("<root a=""x""><item id=""1"">one</item></root>")
Out TypeName(doc.documentElement) & "|" & doc.documentElement.tagName & "|" & doc.documentElement.getAttribute("a") & "|" & IsNull(doc.documentElement.getAttribute("zz"))
Out TypeName(doc.selectNodes("//item")) & "|" & TypeName(doc.selectSingleNode("//item"))
Out doc.selectSingleNode("//item").xml
' Building a document (issue #91): createProcessingInstruction, createElement, insertBefore, firstChild, appendChild and save
Dim b : Set b = CreateObject("Msxml2.DOMDocument")
Out "[" & b.xml & "]|" & TypeName(b.firstChild)
Dim pi : Set pi = b.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
Out TypeName(pi) & "|" & pi.nodeName & "|" & pi.xml & "|" & pi.nodeValue
Dim added : Set added = b.insertBefore(pi, b.firstChild)
Out TypeName(added) & "|" & (added Is pi) & "|" & (b.firstChild Is pi)
Dim root : Set root = b.createElement("ASAPBatch")
Out TypeName(root) & "|" & root.xml
Out (b.appendChild(root) Is root) & "|" & (b.documentElement Is root)
root.setAttribute "xmlns", "urn:batch"
root.setAttribute "xmlns:dt", "urn:datatypes"
root.setAttribute "q", "a<b>c&d""e'f"
Dim session : Set session = b.createElement("Session")
root.AppendChild (session)
session.setAttribute "id", "s1"
Dim request : Set request = b.createElement("Request")
root.appendChild request
Dim company : Set company = b.createElement("dt:CompanyName")
request.AppendChild(company)
company.text = "t<x>&y""z'"
Dim first : Set first = b.createElement("First")
root.insertBefore first, session
Out root.firstChild.nodeName & "|" & TypeName(session.firstChild) & "|" & company.nodeName & "|" & company.tagName
Out b.xml
Out root.xml
Out TypeName(b.selectSingleNode("//Session")) & "|" & b.selectSingleNode("/ASAPBatch/Session").getAttribute("id") & "|" & b.selectNodes("/ASAPBatch/*").length
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim savedPath : savedPath = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName())
b.save savedPath
Out fso.GetFile(savedPath).Size
Out fso.OpenTextFile(savedPath).ReadAll
Dim plain : Set plain = CreateObject("Msxml2.DOMDocument")
plain.appendChild plain.createElement("x")
plain.documentElement.appendChild plain.createElement("y")
plain.save savedPath
Out fso.GetFile(savedPath).Size & "|" & fso.OpenTextFile(savedPath).ReadAll
fso.DeleteFile savedPath
Out plain.loadXML("<?xml version=""1.0"" encoding=""UTF-8""?><a x=""1""><b/><c></c><d>e &amp; &lt; &gt;</d></a>")
Out plain.xml
Out TypeName(plain.firstChild) & "|" & plain.firstChild.nodeName
plain.documentElement.appendChild plain.createElement("n")
Out plain.documentElement.xml
