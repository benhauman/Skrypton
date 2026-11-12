' Parity script: the expected output (ComParity_Msxml2BinaryData.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' MSXML element dataType and nodeTypedValue (issue #170): bin.hex and bin.base64 as Byte(), the dt:dt attribute, the text MSXML writes
Dim d : Set d = CreateObject("Msxml2.DOMDocument")
Dim x : Set x = d.createElement("b")
Out "unset: " & IsNull(x.dataType) & " " & TypeName(x.nodeTypedValue) & " [" & x.nodeTypedValue & "]"
x.dataType = "bin.hex" : x.text = "4772FCdf65"
Dim b : b = x.nodeTypedValue
Out TypeName(b) & " " & LenB(b) & " " & VarType(b) & " " & AscB(MidB(b, 3, 1)) & " " & x.dataType & " " & x.xml
Dim y : Set y = d.createElement("c")
y.dataType = "bin.base64" : y.nodeTypedValue = b
Out y.text & " " & y.dataType & " " & y.xml
Dim z : Set z = d.createElement("h") : z.dataType = "bin.hex" : z.nodeTypedValue = b
Out z.text
Dim s : Set s = CreateObject("ADODB.Stream") : s.Type = 2 : s.Charset = "us-ascii" : s.Open : s.WriteText String(100, "A") : s.Position = 0 : s.Type = 1
Dim big : big = s.Read : s.Close
y.nodeTypedValue = big
Out Replace(Replace(y.text, vbCr, "\r"), vbLf, "\n")
Dim w : Set w = d.createElement("w") : w.dataType = "bin.base64" : w.text = "R3L832U="
Out LenB(w.nodeTypedValue)
d.appendChild y
Out d.documentElement.dataType & " " & TypeName(d.documentElement.nodeTypedValue)
