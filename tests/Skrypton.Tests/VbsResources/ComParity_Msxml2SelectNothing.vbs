' Parity script: the expected output (ComParity_Msxml2SelectNothing.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' MSXML: selectSingleNode of an element (or any node) and the item of a node list out of range are Nothing (issue #179), as for the document
Dim doc : Set doc = CreateObject("Msxml2.DOMDocument.6.0")
doc.loadXML "<r><a>v</a></r>"
Dim root : Set root = doc.documentElement
Dim n : Set n = root.selectSingleNode("b")
Out TypeName(n) & "," & (n Is Nothing) & "," & (doc.selectSingleNode("/r/b") Is Nothing)
If Not (root.selectSingleNode("b") Is Nothing) Then
    Out root.selectSingleNode("b").text
End If
Out TypeName(root.selectSingleNode("a")) & " " & root.selectSingleNode("a").text
Dim l : Set l = doc.selectNodes("//a")
Out TypeName(l.item(5)) & "," & (l.item(5) Is Nothing) & "," & TypeName(l.item(0))
Dim c : Set c = root.firstChild.selectSingleNode("x")
Out TypeName(c) & "," & (c Is Nothing)
Dim d2 : Set d2 = CreateObject("Msxml2.DOMDocument")
d2.loadXML "<r><a>v</a></r>"
Out TypeName(d2.documentElement.selectSingleNode("b")) & "," & (d2.documentElement.selectSingleNode("b") Is Nothing)
Out "done"
Dim tags : Set tags = CreateObject("Scripting.Dictionary")
tags.Add "a", "" : tags.Add "b", ""
Dim tag
For Each tag In tags.Keys
    If Not (root.selectSingleNode(tag) Is Nothing) Then tags.Item(tag) = root.selectSingleNode(tag).text
Next
Out "a=[" & tags("a") & "] b=[" & tags("b") & "]"
