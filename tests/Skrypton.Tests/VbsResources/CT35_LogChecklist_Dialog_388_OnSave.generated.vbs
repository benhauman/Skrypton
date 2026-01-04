'Check if invalid characters are in any of the url textboxes
Set dict = CreateObject("Scripting.Dictionary")
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = true
objRegEx.Pattern = "[^A-Z0-9][^\:][^\/][^\.][^\S][^\?][^\€][^\@]"


dict.Add "Checkliste 1 URL", TextBoxChecklist1URL.Text
dict.Add "Checkliste 2 URL", TextBoxChecklist2URL.Text
dict.Add "Checkliste 3 URL", TextBoxChecklist3URL.Text
dict.Add "Checkliste 4 URL", TextBoxChecklist4URL.Text
dict.Add "Checkliste 5 URL", TextBoxChecklist5URL.Text
dict.Add "Checkliste 6 URL", TextBoxChecklist6URL.Text
dict.Add "Checkliste 7 URL", TextBoxChecklist7URL.Text
dict.Add "Checkliste 8 URL", TextBoxChecklist8URL.Text
dict.Add "Checkliste 9 URL", TextBoxChecklist9URL.Text
dict.Add "Checkliste 10 URL", TextBoxChecklist10URL.Text

Dim element
For Each element In dict
  IF dict(element) <> "" THEN
    Set match = objRegEx.execute(dict(element))
    IF match.Count > 0 THEN
      Dim errMsg
      errMsg = model.Translate("#ERR_Checklists_InvalidChars")
      errMsg = Replace(errMsg, "{0}", element)
      model.MsgBox errMsg
      model.CurrentCommand.Abort "OnSave"
    END IF
  END IF
Next
