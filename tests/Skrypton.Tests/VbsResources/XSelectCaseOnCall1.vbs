Dim hlObj, CaseAttributes, backColor
Select Case hlObj.GetValue("CaseClassificationAttribute.Priority",0,0,0,0)
Case "Priority1"
CaseAttributes.BackColor = "RGB(107,105,248)"
Case "Priority2" 
CaseAttributes.BackColor = "RGB(119,170,251)"
Case Else
CaseAttributes.BackColor = "RGB(248,245,240)"
End Select