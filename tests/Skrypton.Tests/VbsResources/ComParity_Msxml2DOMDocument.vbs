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
' Selection languages (issue #94): XSLPattern (the default of Msxml2.DOMDocument) matches names as they are written, XPath resolves
' prefixes through SelectionNamespaces
Function NodeText(n)
    If n Is Nothing Then
        NodeText = "Nothing"
    Else
        NodeText = n.nodeName & "=" & n.text
    End If
End Function
Dim sel : Set sel = CreateObject("Msxml2.DOMDocument")
Out sel.getProperty("SelectionLanguage") & "|" & sel.getProperty("SelectionNamespaces") & "|"
Dim selRoot : Set selRoot = sel.createElement("root")
sel.appendChild selRoot
selRoot.setAttribute "xmlns", "urn:x"
selRoot.setAttribute "xmlns:dt", "urn:dt"
Dim selChild : Set selChild = sel.createElement("dt:c")
selRoot.appendChild selChild
selChild.text = "built"
Set selChild = sel.createElement("plain")
selRoot.appendChild selChild
selChild.text = "p"
Out NodeText(sel.selectSingleNode("//dt:c")) & " " & NodeText(sel.selectSingleNode("//plain")) & " " & sel.selectNodes("//dt:*").length & " " & NodeText(sel.selectSingleNode("/root/dt:c")) & " " & NodeText(selRoot.selectSingleNode("dt:c"))
Dim lsel : Set lsel = CreateObject("Msxml2.DOMDocument")
Out lsel.loadXML("<root xmlns='urn:x' xmlns:dt='urn:dt'><dt:c n='1'>loaded</dt:c><plain id='7'>p</plain><item id='1'>one</item><item id='2'>two</item><x:y xmlns:x='urn:dt'>xy</x:y></root>")
Out NodeText(lsel.selectSingleNode("//dt:c")) & " " & NodeText(lsel.selectSingleNode("//plain")) & " " & NodeText(lsel.selectSingleNode("//c")) & " " & NodeText(lsel.selectSingleNode("//dt:y")) & " " & NodeText(lsel.selectSingleNode("//x:y"))
Out NodeText(lsel.selectSingleNode("//item[@id='2']")) & " " & NodeText(lsel.selectSingleNode("//item[@id = ""2""]")) & " " & NodeText(lsel.selectSingleNode("//item[@id='2' and @id!='1']")) & " " & NodeText(lsel.selectSingleNode("//item[@id > 1]"))
Out NodeText(lsel.selectSingleNode("//item[. = 'two']")) & " " & NodeText(lsel.selectSingleNode("/root/item[0]")) & " " & NodeText(lsel.selectSingleNode("/root/item[ 1 ]")) & " " & NodeText(lsel.selectSingleNode("/root/item[2]")) & " " & lsel.selectNodes("//item | //plain").length & " " & lsel.selectNodes("//@id").length & " " & lsel.selectNodes("/root/*").length
Out NodeText(lsel.selectSingleNode("//dt:c/@n")) & " " & lsel.selectNodes("//dt:*/@*").length & " " & NodeText(lsel.selectSingleNode("//item/text()")) & " " & NodeText(lsel.documentElement.selectSingleNode("item[1]"))
lsel.setProperty "SelectionLanguage", "XPath"
On Error Resume Next
Set selChild = lsel.selectSingleNode("//dt:c")
Out Err.Number & " " & Err.Source & " " & Err.Description
Err.Clear
lsel.setProperty "SelectionLanguage", "xpath"
Out Err.Number & " " & Err.Description & lsel.getProperty("SelectionLanguage")
Err.Clear
lsel.setProperty "NoSuchProperty", 1
Out Err.Number & " " & Err.Description
Err.Clear
Out lsel.getProperty("selectionlanguage")
Out Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0
Out NodeText(lsel.selectSingleNode("//plain")) & " " & NodeText(lsel.selectSingleNode("//*[local-name()='plain']"))
lsel.setProperty "SelectionNamespaces", "xmlns:q='urn:dt' xmlns:d=""urn:x"""
Out NodeText(lsel.selectSingleNode("//q:c")) & " " & NodeText(lsel.selectSingleNode("//q:y")) & " " & NodeText(lsel.selectSingleNode("//d:plain")) & " " & lsel.selectNodes("//d:item").length
Out lsel.getProperty("SelectionNamespaces") & "|" & lsel.getProperty("SelectionLanguage")
sel.setProperty "SelectionLanguage", "XPath"
sel.setProperty "SelectionNamespaces", "xmlns:dt='urn:dt'"
Out NodeText(sel.selectSingleNode("//dt:c")) & " " & NodeText(sel.selectSingleNode("//plain"))
' Attributes, CDATA sections and text nodes (issue #121): createAttribute, attributes (setNamedItem, getNamedItem, item, removeNamedItem, For Each),
' createCDATASection ("]]>" is an error with MSXML 3), createTextNode, childNodes, save
Sub AttributesAndCData()
    Dim xmldoc : Set xmldoc = CreateObject("msxml2.DomDocument")
    Dim nodeData : Set nodeData = xmldoc.appendChild(xmldoc.createElement("Data"))
    Dim nodeObject : Set nodeObject = nodeData.appendChild(xmldoc.createElement("Asset"))
    Dim attAliasName : Set attAliasName = xmldoc.createAttribute("aliasname")
    attAliasName.Text = "NewCIAsset"
    nodeObject.Attributes.setNamedItem attAliasName
    Dim valueNode : Set valueNode = xmldoc.createElement("AssetGeneral.AssetName")
    Dim cdata : Set cdata = xmldoc.createCDATASection("PC <4711> & Co")
    valueNode.appendChild(cdata)
    nodeObject.appendChild(valueNode)
    Dim attIsKey : Set attIsKey = xmldoc.createAttribute("iskey")
    attIsKey.Text = "false"
    valueNode.Attributes.setNamedItem attIsKey
    Out xmldoc.xml
    Out nodeObject.Attributes.length & " " & nodeObject.Attributes.getNamedItem("aliasname").Text & " " & TypeName(attAliasName) & " " & TypeName(cdata)
    Out TypeName(nodeObject.Attributes) & " " & TypeName(nodeObject.Attributes.getNamedItem("nope"))
    Out attAliasName.nodeName & "|" & attAliasName.nodeValue & "|" & attAliasName.xml & "|" & cdata.nodeName & "|" & cdata.nodeValue & "|" & cdata.text & "|" & cdata.xml
    Dim att2 : Set att2 = xmldoc.createAttribute("aliasname")
    att2.Text = "Replaced"
    Dim old : Set old = nodeObject.Attributes.setNamedItem(att2)
    Out TypeName(old) & " " & nodeObject.Attributes.length & " " & nodeObject.getAttribute("aliasname")
    nodeObject.setAttribute "z", "last"
    Out nodeObject.Attributes.item(1).nodeName & " " & nodeObject.Attributes(0).nodeName & " " & nodeObject.Attributes.length
    Dim at, s
    For Each at In nodeObject.Attributes
        s = s & at.nodeName & "=" & at.nodeValue & ";"
    Next
    Out s
    Dim e2 : Set e2 = xmldoc.createElement("e")
    e2.appendChild xmldoc.createCDATASection("")
    On Error Resume Next
    e2.appendChild xmldoc.createCDATASection("a]]>b")
    Out "cdata ]]>: " & Err.Number & " " & Err.Source & " " & Replace(Err.Description, vbCrLf, "")
    Err.Clear
    On Error GoTo 0
    e2.appendChild xmldoc.createTextNode("t<&>")
    nodeData.appendChild e2
    Out e2.xml & " " & e2.text & " " & e2.childNodes.length
    Out TypeName(xmldoc.createTextNode("x"))
    Out TypeName(xmldoc.selectSingleNode("//Asset/@aliasname")) & " " & xmldoc.selectSingleNode("//Asset/@aliasname").text
    Out xmldoc.selectSingleNode("//Asset").attributes.getNamedItem("z").value
    nodeObject.Attributes.removeNamedItem "z"
    Out nodeObject.Attributes.length
    Dim attrFso : Set attrFso = CreateObject("Scripting.FileSystemObject")
    Dim attrPath : attrPath = attrFso.BuildPath(attrFso.GetSpecialFolder(2), attrFso.GetTempName())
    xmldoc.save attrPath
    Out attrFso.GetFile(attrPath).Size & "|" & attrFso.OpenTextFile(attrPath, 1).ReadAll
    attrFso.DeleteFile attrPath
End Sub
AttributesAndCData
