SUB ButtonCreateWord_Click()
	Dim group, sum, ust, usum, prozent, waehrung, name , ort, str, tel, mail, inet, fax, anrede, titel, timestamp, objWord, objDoc, objTable, contCount,objRange, gender
	
	timestamp = Now
	
			prozent = 19
			waehrung = " €"
			name = "helpLine GmbH"
			ort = "D-65520 Bad Camberg"	
			str = "Carl-Zeiss-Straße 16"
			tel = "+49 (0) 6434 930 76 0"	
			mail = "kontakt@helpline.de"
			inet = "www.helpline.de"		
			fax = "+49 (0) 6434 930 76 300"
		
	Set objWord = CreateObject("Word.Application")
	objWord.Caption = "Angebot"
	objWord.Visible=true
	
	
	'Pfad zum Vorlagedokument
	Set objDoc = objWord.Documents.Open ("C:\helpline\Dokumente\A-XXX-VL-Firma-hL_Cons.docx")
	Set objTable = objDoc.Tables(1)
	
	Dim Contid, stundensatz,counter, tag, pm, pmmail, pmtel, monat, dltyp, art, bezeichnung  
	counter =1
	stundensatz = ""
	stundensatz = CDBL(0.0)
	
	Contid = 0
	
	Contid = 60001 'For Each Contid in 
	
		dltyp = "DLTypeSRM"
		group = "SRMWorkGrouphelpLine"
		
		art = "Dienstleistung" & vbCRLF & "(Zeitstunden)"
		bezeichnung ="helpLine Consulting" & vbCrLf
		bezeichnung = bezeichnung & "(Service Request Management):" & vbCrLf & vbCrLf
			 bezeichnung = bezeichnung & "Gerne setzen wir Ihre Anforderung im Rahmen" & vbCrLf
		bezeichnung = bezeichnung & "des Service Request Management kurzfristig" & vbCrLf
		bezeichnung = bezeichnung & "um. Ihr Vorteil: Schnelle und unkomplizierte" &vbCrLf
		bezeichnung = bezeichnung & "remote Unterstützung durch erfahrene" & vbCrLf
		bezeichnung = bezeichnung & "helpLine Consultants ohne lange Wartezeiten." & vbCrLf & vbCrLf 
		bezeichnung = bezeichnung & "Die Abrechnung erfolgt nach geleisteten" & vbCrLf
		bezeichnung = bezeichnung & "Stunden."
		
		objTable.Rows(counter + 1).Cells(1).Range = "DL-" & counter
		
		objTable.Rows(counter + 1).Cells(2).Range = art
		
		objTable.Rows(counter + 1).Cells(3).Range = bezeichnung 
		
		objTable.Rows(counter + 1).Cells(4).Range = 5055
		
		objTable.Rows(counter + 1).Cells(5).Range = 7006
		Dim summe, aufwand 
		summe = 0.00
		aufwand = 0.00
		aufwand =CDBL(71) * CDBL(1.00)
		summe = aufwand * CDbl(56)
		
		objTable.Rows(counter + 1).Cells(6).Range = summe & waehrung
		
		sum =sum + summe
		
		If contCount >= counter Then 
			objTable.Rows.Add()
		End if
		counter = counter +1
	
	'Next : For Each Contid in 
	
	'Lese Bookmark aus Word anhand der Bezeichnung aus unf fülle es mit meinen gesetzten Attributen aus Vorgang
	Set objRange = objDoc.Bookmarks("firmenname").Range
	objRange.Text ="bzzzz"
	
	gender= "GenderMale"
		titel = "Herr "
		anrede = "Sehr geehrter Herr "
	
	Set objRange = objDoc.Bookmarks("Vorname").Range
	objRange.Text =titel & "gg" & " " & "hh"
	
	Set objRange = objDoc.Bookmarks("email").Range
	objRange.Text = "zz@srv1.com"
	
	Set objRange = objDoc.Bookmarks("telefonnr").Range
	objRange.Text = "+004912345678"
	
	Dim plz2
	plz2 = "PLZ50670"
	plz2 = plz2
	
	Dim street 
	street = "MainStreet"
	Set objRange = objDoc.Bookmarks("Strasse").Range
	objRange.Text = street
	
	Set objRange = objDoc.Bookmarks("PLZ").Range
	objRange.Text = plz2
	
	Set objRange = objDoc.Bookmarks("Stadt").Range
	objRange.Text = "Cologne"
	
	Set objRange = objDoc.Bookmarks("pmcsName").Range
	objRange.Text = name
	
	Set objRange = objDoc.Bookmarks("pmcsOrt").Range
	objRange.Text = ort
	
	Set objRange = objDoc.Bookmarks("pmcsStr").Range
	objRange.Text = str
	
	Set objRange = objDoc.Bookmarks("pmcsTel").Range
	objRange.Text = tel
	
	Set objRange = objDoc.Bookmarks("pmcsMail").Range
	objRange.Text = mail
	
	Set objRange = objDoc.Bookmarks("pmcsFax").Range
	objRange.Text = fax
	
	Set objRange = objDoc.Bookmarks("pmcsInet").Range
	objRange.Text = inet
	
	pm = "nadine"
	Set objRange = objDoc.Bookmarks("projectmanager").Range
	objRange.Text = pm
	
	pmmail = "nadine@test.de"
	pmtel = "+492233"
	
	Set objRange = objDoc.Bookmarks("telprojectm").Range
	objRange.Text = pmtel
	
	Set objRange = objDoc.Bookmarks("emailprojectm").Range
	objRange.Text = pmmail
	
	Set objRange = objDoc.Bookmarks("TicketNr").Range
	objRange.Text = "20260406-0001"
	
	Set objRange = objDoc.Bookmarks("nachname").Range
	objRange.Text =anrede & "jarr"
	
	Set objRange = objDoc.Bookmarks("GueltigBis").Range
	objRange.Text = DateAdd("d",14,Date)
	monat =CStr(Month(date))
	tag =CStr(Day(date))
	
	monat = "0" & Month(date)
	
	tag = "0" & Day(date)
	
	Set objRange = objDoc.Bookmarks("Angebot").Range
	objRange.Text = Year(date) &monat & tag & "-" & "OUx" & "cons"
	
	Set objRange = objDoc.Bookmarks("ExclSumm").Range
	objRange.Text = Round(sum,2)
	
	ust = (prozent / 100) * sum
	Set objRange = objDoc.Bookmarks("UST").Range
	objRange.Text =Round(ust, 2)
	
	usum = sum + ust
	Set objRange = objDoc.Bookmarks("inklSum").Range
	objRange.Text = Round(usum, 2)
	
END SUB
ButtonCreateWord_Click