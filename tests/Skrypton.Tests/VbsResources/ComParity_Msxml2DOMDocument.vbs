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
