' Parity script: the expected output (ComParity_UndeclaredVariables.expected.txt) was captured with Windows cscript.exe.
' Without Option Explicit, variables that are not declared are implicitly declared Variants (issue #86).
total = 0
For i = 1 To 5
    total = total + i
    text = text & i
Next
Out total & "|" & text & "|" & TypeName(total) & "|" & i
name = "first"
name = name & " second"
Out name
Set dict = CreateObject("Scripting.Dictionary")
dict.Add "a", 1
Set dict = Nothing
Out TypeName(dict)
Sub Accumulate(value)
    localSum = value * 2
    counter = counter + 1
    Out "in sub: " & localSum & "|" & counter
End Sub
counter = 10
Accumulate 21
Out "after sub: " & counter
Function Twice(value)
    doubled = value + value
    Twice = doubled
End Function
Out Twice(4) & "|" & TypeName(neverAssigned) & "|" & IsEmpty(neverAssigned)
arr = Array(1, 2, 3)
arr(1) = "x"
Out Join(arr, ",") & "|" & UBound(arr)
For Each item In arr
    joined = joined & "[" & item & "]"
Next
Out joined
