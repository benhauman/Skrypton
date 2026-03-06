Dim hlObj, CaseAttributes, backColor
SELECT CASE hlObj.GetValue("CaseClassificationAttribute.Priority", 0, 0, 0, 0)
  CASE "Priority1"
    CaseAttributes.BackColor = "RGB(107,105,248)"
  CASE "Priority2"
    CaseAttributes.BackColor = "RGB(119,170,251)"
  CASE ELSE
    CaseAttributes.BackColor = "RGB(248,245,240)"
END SELECT
