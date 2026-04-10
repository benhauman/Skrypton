Dim doc, hlOrgUnit
doc.Bookmarks("Firma").Range.Text = CStr(hlOrgUnit.GetValue("OrganizationGeneral.Name",0,0,0,0))