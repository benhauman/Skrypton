' Parity script: the expected output (ComParity_ByteFunctions.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' Byte strings (issue #93): LenB, LeftB, RightB, MidB, InStrB, AscB, ChrB - a string is its UTF-16LE bytes, it may have an odd number of them
Dim s : s = "AbC"
Out LenB(s) & " " & AscB(s) & " " & AscB(MidB(s, 2, 1)) & " " & LenB(MidB(s, 3)) & " " & AscB(MidB(s, 3))
Out LenB(LeftB(s, 3)) & " " & AscB(RightB(s, 2)) & " " & InStrB(s, "C") & " " & InStrB(2, s, ChrB(0)) & " " & InStrB(s, "x")
Out LenB(ChrB(65)) & " " & AscB(ChrB(200)) & " " & (ChrB(65) & ChrB(0) = "A") & " " & VarType(ChrB(65)) & " " & VarType(MidB(s, 1, 2))
Out LeftB(s, 2) & "|" & MidB(s, 3, 2) & "|" & RightB(s, 4) & "|" & InStrB(1, s, "b", 1) & "|" & InStrB(1, s, "B", 1) & "|" & InStrB(1, s, "B") & "|" & InStrB(s, "")
Out TypeName(AscB("A")) & " " & TypeName(MidB(s, 1)) & " " & TypeName(InStrB(s, "C")) & " " & TypeName(LenB("A")) & " " & TypeName(ChrB(1))
Out LenB(ChrB(65) & ChrB(66) & ChrB(67)) & " " & Len(ChrB(65) & ChrB(66)) & " " & (ChrB(65) & ChrB(0) & ChrB(66) & ChrB(0) = "AB")
Out LenB(MidB(s, 2)) & " " & AscB(MidB(s, 2)) & " " & LenB(MidB(s, 7)) & " " & LenB(MidB(s, 5, 10)) & " " & LenB(LeftB(s, 0)) & " " & LenB(RightB(s, 100))
Out InStrB(7, s, "C") & " " & InStrB(6, s, "") & " " & InStrB(7, s, "") & " " & InStrB(8, s, "") & " " & InStrB(s, "bC") & " " & InStrB(4, s, "b")
Out LenB("") & " " & IsNull(LenB(Null)) & " " & IsNull(MidB(Null, 1)) & " " & IsNull(LeftB(Null, 1)) & " " & IsNull(RightB(Null, 1)) & " " & IsNull(InStrB(Null, "a")) & " " & IsNull(InStrB("a", Null))
Out LenB(12) & " " & AscB(12) & " " & LenB(True) & " " & LenB(Empty)
Dim hx, i, odd : odd = ChrB(1) & ChrB(2) & ChrB(255)
For i = 1 To LenB(odd) : hx = hx & Right("0" & Hex(AscB(MidB(odd, i, 1))), 2) : Next
Out hx

' a byte array (ADODB.Stream.Read) is its bytes
Dim st : Set st = CreateObject("ADODB.Stream")
st.Type = 2 : st.Charset = "us-ascii" : st.Open : st.WriteText "Hello"
st.Position = 0 : st.Type = 1
Dim b : b = st.Read : st.Close
Out TypeName(b) & " " & LenB(b) & " " & AscB(b) & " " & AscB(MidB(b, 2, 1)) & " " & InStrB(b, ChrB(108)) & " " & AscB(RightB(b, 1)) & " " & LenB(LeftB(b, 3))
hx = "" : For i = 1 To LenB(b) : hx = hx & Right("0" & Hex(AscB(MidB(b, i, 1))), 2) : Next
Out hx

' errors
On Error Resume Next
Dim x
x = AscB("") : Out "AscB("""") " & Err.Number & " " & Err.Description : Err.Clear
x = AscB(Null) : Out "AscB(Null) " & Err.Number & " " & Err.Description : Err.Clear
x = MidB(s, 0) : Out "MidB(s, 0) " & Err.Number & " " & Err.Description : Err.Clear
x = MidB(s, 1, -1) : Out "MidB(s, 1, -1) " & Err.Number & " " & Err.Description : Err.Clear
x = LeftB(s, -1) : Out "LeftB(s, -1) " & Err.Number & " " & Err.Description : Err.Clear
x = RightB(s, -1) : Out "RightB(s, -1) " & Err.Number & " " & Err.Description : Err.Clear
x = ChrB(256) : Out "ChrB(256) " & Err.Number & " " & Err.Description : Err.Clear
x = ChrB(-1) : Out "ChrB(-1) " & Err.Number & " " & Err.Description : Err.Clear
x = ChrB(Null) : Out "ChrB(Null) " & Err.Number & " " & Err.Description : Err.Clear
x = InStrB(0, s, "b") : Out "InStrB(0, s, ""b"") " & Err.Number & " " & Err.Description : Err.Clear
x = InStrB(1, s, "b", 2) : Out "InStrB(1, s, ""b"", 2) " & Err.Number & " " & Err.Description : Err.Clear
On Error GoTo 0
