' Parity script: the expected output (ComParity_Msxml2DocumentProperties.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' MSXML DOMDocument: validateOnParse and resolveExternals (issue #164) - defaults of MSXML 3 and 6, set to False before loadXML, set from a number / text
Dim p : For Each p In Array("MSXML2.DOMDocument", "Microsoft.XMLDOM", "Msxml2.DOMDocument.6.0")
Dim x : Set x = CreateObject(p)
Out p & ": " & x.validateOnParse & "|" & x.resolveExternals & "|" & x.preserveWhiteSpace
x.async = False : x.validateOnParse = False : x.resolveExternals = False
Out x.validateOnParse & "|" & x.resolveExternals & "|" & x.loadXML("<a b=""1""/>") & "|" & TypeName(x.validateOnParse)
x.validateOnParse = 1 : x.resolveExternals = "True"
Out x.validateOnParse & "|" & x.resolveExternals
Next
