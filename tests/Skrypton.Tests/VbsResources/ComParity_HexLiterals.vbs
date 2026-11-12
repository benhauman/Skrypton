' Parity script: the expected output (ComParity_HexLiterals.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' Hex and octal literals (issue #169): &H8000 to &HFFFF are negative 16-bit values, &H80000000 to &HFFFFFFFF negative 32-bit values, a trailing &
' makes a Long; the type is the one of VBScript (-32768 and &H8000 are Long)
Dim v
v = &H7FFF : Out "&H7FFF " & v & " " & TypeName(v)
v = &H8000 : Out "&H8000 " & v & " " & TypeName(v)
v = &H8001 : Out "&H8001 " & v & " " & TypeName(v)
v = &HDFFC : Out "&HDFFC " & v & " " & TypeName(v)
v = &HFFFF : Out "&HFFFF " & v & " " & TypeName(v)
v = &H10000 : Out "&H10000 " & v & " " & TypeName(v)
v = &H7FFFFFFF : Out "&H7FFFFFFF " & v & " " & TypeName(v)
v = &H80000000 : Out "&H80000000 " & v & " " & TypeName(v)
v = &HFFFFFFFF : Out "&HFFFFFFFF " & v & " " & TypeName(v)
v = &H8000& : Out "&H8000& " & v & " " & TypeName(v)
v = &HFFFF& : Out "&HFFFF& " & v & " " & TypeName(v)
v = &O177777 : Out "&O177777 " & v & " " & TypeName(v)
v = -&H8000 : Out "-&H8000 " & v & " " & TypeName(v)
Out ChrW(&HDFFC) = ChrW(-8196)
Out (&HFFFF) + 1
Out TypeName(-32768) & " " & TypeName(-32767) & " " & TypeName(32767) & " " & TypeName(&H8000) & " " & TypeName(&H8001)
Out "&H41 " & "|" & &H41 & "|" & &h1f & "|" & &O17
