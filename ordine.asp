<!--#include file="inc_strConn.asp"-->
<%
	Call Visualizzazione("",0,"ordine.asp")
	
	mode=request("mode")
	if mode="" then mode=0
	
	IdOrdine=session("ordine_shop")
	if IdOrdine="" then IdOrdine=0
	if idOrdine=0 then response.redirect("carrello1.asp")
	
	if idsession=0 then response.Redirect("iscrizione.asp?prov=1")
		
	session("ordine_shop")=""
	
	
	Set ss = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where pkid="&idOrdine
	ss.Open sql, conn, 3, 3
	
	if ss.recordcount>0 then
		TotaleCarrello=ss("TotaleCarrello")
		CostoSpedizioneTotale=ss("CostoSpedizione")
		TipoTrasporto=ss("TipoTrasporto")
		DatiSpedizione=ss("DatiSpedizione")
		NoteCliente=ss("NoteCliente")
		
		FkPagamento=ss("FkPagamento")
		TipoPagamento=ss("TipoPagamento")
		CostoPagamento=ss("CostoPagamento")
		
		Nominativo=ss("Nominativo")
		Rag_Soc=ss("Rag_Soc")
		Cod_Fisc=ss("Cod_Fisc")
		PartitaIVA=ss("PartitaIVA")
		Indirizzo=ss("Indirizzo")
		Citta=ss("Citta")
		Provincia=ss("Provincia")
		CAP=ss("CAP")
		
		TotaleGenerale=ss("TotaleGenerale")
		
		DataAggiornamento=ss("DataAggiornamento")
		
		ss("stato")=6
		ss("DataAggiornamento")=now()
		ss("IpOrdine")=Request.ServerVariables("REMOTE_ADDR")
		ss.update	
	end if
	
	ss.close
	
	if FkPagamento=1 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Clienti where pkid="&idsession
		rs.Open sql, conn, 1, 1	
		
		nominativo_email=rs("nome")&" "&rs("nominativo")
		email=rs("email")
		
		rs.close
			
			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Cristalensi</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Grazie "&nominativo_email&" per aver scelto i nostri prodotti!<br>Questa &egrave; un email di conferma per il completamento dell'ordine n&deg; "&idordine&".<br><br><b>TOTALE ORDINE: <u>"&TotaleGenerale&"&#8364;</u></b><br><br> Per completare l'ordine &egrave; necessario effettuare il bonifico con i seguenti dati:<br><u>BANCA DI CREDITO COOPERATIVO DI CAMBIANO AG. MONTELUPO FIORENTINO</u><br>IBAN: <u>IT33E0842537960000030277941</u><br>Codice BIC/SWIFT: <u>CRACIT33</u><br>Nella causale indicare: Ordine da sito internet n&deg; "&idordine&"<br><br>Beneficiario: Cristalensi Snc di Lensi Massimiliano & C. (P.Iva e C.Fiscale 05305820481)<br>Via arti e mestieri, 1 - 50056 Montelupo F.no (FI)<br><br><br>Il nostro staff avr&agrave; cura di spedirti la merce appena la banca avr&agrave; notificato il pagamento del bonifico oppure, per velocizzare la spedizione, &egrave; possibile inviarci per email la ricevuta dell'avvenuto pagamento con bonifico (in caso di bonifico fatto con home banking spesso viene fornita dalla banca una ricevuta, oppure &egrave; possibile scannerizzare la ricevuta rilasciata dalla banca).</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it con pagamento con bonifico bancario"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")
		
			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig 
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
				' Timeout 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.cristalensi.it"
				' Porta SMTP 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@cristalensi.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "m0nt3lup0"
				
				.Fields.update 
			End With 
			Set eMail_cdo.Configuration = myConfig
		
			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto
		
			eMail_cdo.HTMLBody = Testo
		
			eMail_cdo.Send()
		
			Set myConfig = Nothing
			Set eMail_cdo = Nothing
			
			'fine invio email
			
			'invio l'email all'amministratore
			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Cristalensi</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento con bonifico dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it con pagamento con bonifico bancario"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")
		
			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig 
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
				' Timeout 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.cristalensi.it"
				' Porta SMTP 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@cristalensi.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "m0nt3lup0"
				
				.Fields.update 
			End With 
			Set eMail_cdo.Configuration = myConfig
		
			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto
		
			eMail_cdo.HTMLBody = Testo
		
			eMail_cdo.Send()
		
			Set myConfig = Nothing
			Set eMail_cdo = Nothing
			
			'invio al webmaster
			
			Mittente = "info@cristalensi.it"
			Destinatario = "viadeimedici@gmail.com"
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it con pagamento con bonifico bancario"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")
		
			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig 
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
				' Timeout 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.cristalensi.it"
				' Porta SMTP 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@cristalensi.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "m0nt3lup0"
				
				.Fields.update 
			End With 
			Set eMail_cdo.Configuration = myConfig
		
			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto
		
			eMail_cdo.HTMLBody = Testo
		
			eMail_cdo.Send()
		
			Set myConfig = Nothing
			Set eMail_cdo = Nothing
			
			'fine invio email
	end if
	
	if FkPagamento=3 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Clienti where pkid="&idsession
		rs.Open sql, conn, 1, 1	
		
		nominativo_email=rs("nome")&" "&rs("nominativo")
		email=rs("email")
		
		rs.close
			
			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Cristalensi</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Grazie "&nominativo_email&" per aver scelto i nostri prodotti!<br>Questa &egrave; un email di conferma per il completamento dell'ordine n&deg; "&idordine&".<br><br><br>Il nostro staff avr&agrave; cura di spedirti la merce appena sar&agrave; disponibile nel nostro magazino.<br>Ti ricordiamo che per il pagamento in contrassegno, il corriere consegner&agrave; la merce solo se verr&agrave; pagata in contanti, non accetter&agrave; altri metodi di pagamento (anche gli assegni non saranno accettati).</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it con pagamento in contrassegno"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")
		
			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig 
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
				' Timeout 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.cristalensi.it"
				' Porta SMTP 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@cristalensi.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "m0nt3lup0"
				
				.Fields.update 
			End With 
			Set eMail_cdo.Configuration = myConfig
		
			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto
		
			eMail_cdo.HTMLBody = Testo
		
			eMail_cdo.Send()
		
			Set myConfig = Nothing
			Set eMail_cdo = Nothing
			
			'fine invio email
			
			'invio l'email all'amministratore
			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Cristalensi</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento in contrassegno dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it con pagamento in contrassegno"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")
		
			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig 
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
				' Timeout 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.cristalensi.it"
				' Porta SMTP 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@cristalensi.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "m0nt3lup0"
				
				.Fields.update 
			End With 
			Set eMail_cdo.Configuration = myConfig
		
			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto
		
			eMail_cdo.HTMLBody = Testo
		
			eMail_cdo.Send()
		
			Set myConfig = Nothing
			Set eMail_cdo = Nothing
			
			'invio al webmaster
			
			Mittente = "info@cristalensi.it"
			Destinatario = "viadeimedici@gmail.com"

			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it con pagamento in contrassegno"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")
		
			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig 
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
				' Timeout 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.cristalensi.it"
				' Porta SMTP 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@cristalensi.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "m0nt3lup0"
				
				.Fields.update 
			End With 
			Set eMail_cdo.Configuration = myConfig
		
			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto
		
			eMail_cdo.HTMLBody = Testo
		
			eMail_cdo.Send()
		
			Set myConfig = Nothing
			Set eMail_cdo = Nothing
			
			'fine invio email
	end if
	
	if FkPagamento=4 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Clienti where pkid="&idsession
		rs.Open sql, conn, 1, 1	
		
		nominativo_email=rs("nome")&" "&rs("nominativo")
		email=rs("email")
		
		rs.close
			
			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Cristalensi</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Grazie "&nominativo_email&" per aver scelto i nostri prodotti!<br>Questa è  un email di conferma per il completamento dell'ordine n° "&idordine&".<br><br><strong>TOTALE ORDINE: <u>"&TotaleGenerale&"&#8364;</u></strong><br><br> Per completare l'ordine è necessario effettuare il pagamento su Carta POSTEPAY con i seguenti dati:<br><br><strong>Beneficiario: LENSI GIULIANO - c.f. LNS GLN 42A30 D403J<br>Numero carta: 4023600915402442</strong><br><br>Nella causale indicare: <strong>Ordine da sito internet n&deg; "&idordine&"</strong><br><br><br>Il nostro staff avr&agrave; cura di spedirti la merce appena ricever&agrave; la notifica del pagamento oppure, per velocizzare la spedizione, &egrave; possibile inviarci per email la ricevuta dell'avvenuto pagamento.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it con pagamento con POSTEPAY"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")
		
			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig 
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
				' Timeout 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.cristalensi.it"
				' Porta SMTP 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@cristalensi.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "m0nt3lup0"
				
				.Fields.update 
			End With 
			Set eMail_cdo.Configuration = myConfig
		
			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto
		
			eMail_cdo.HTMLBody = Testo
		
			eMail_cdo.Send()
		
			Set myConfig = Nothing
			Set eMail_cdo = Nothing
			
			'fine invio email
			
			'invio l'email all'amministratore
			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Cristalensi</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento con POSTEPAY dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it con pagamento con POSTEPAY"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")
		
			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig 
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
				' Timeout 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.cristalensi.it"
				' Porta SMTP 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@cristalensi.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "m0nt3lup0"
				
				.Fields.update 
			End With 
			Set eMail_cdo.Configuration = myConfig
		
			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto
		
			eMail_cdo.HTMLBody = Testo
		
			eMail_cdo.Send()
		
			Set myConfig = Nothing
			Set eMail_cdo = Nothing
			
			'invio al webmaster
			
			Mittente = "info@cristalensi.it"
			Destinatario = "viadeimedici@gmail.com"
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it con pagamento con POSTEPAY"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")
		
			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig 
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
				' Timeout 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.cristalensi.it"
				' Porta SMTP 
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@cristalensi.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "m0nt3lup0"
				
				.Fields.update 
			End With 
			Set eMail_cdo.Configuration = myConfig
		
			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto
		
			eMail_cdo.HTMLBody = Testo
		
			eMail_cdo.Send()
		
			Set myConfig = Nothing
			Set eMail_cdo = Nothing
			
			'fine invio email
	end if
%>
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Cristalensi - Ordine</title>
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="/js/media-queries-ie.js"></script>
        <![endif]-->
        <link href="/css/css.css" rel="stylesheet" type="text/css">
        <link href="/css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="/css/tipTip.css" rel="stylesheet" type="text/css">
        
        <link href="/css/cookies-enabler.css" rel="stylesheet" type="text/css">
        
        <script src="http://code.jquery.com/jquery-1.11.2.min.js"></script>
        <script src="/js/jquery.blueberry-min.js"></script>
        <script src="/js/jquery.tipTip-min.js"></script>
        <style type="text/css">
            .clearfix:after {
                content: ".";
                display: block;
                height: 0;
                clear: both;
                visibility: hidden;
            }
        </style>
        <!--[if lt IE 9]>
            <style>
                #menu, #language {
                    display: block !important;
                    
                }
                #language li {
                    display: inline-block !important;
                    float: left !important; 
                    text-align: center !important;
                    padding: 6px 17px !important;
                    height: auto !important;
                    
                }
                #menu li {
                    display: inline-block !important;
                    float: left !important; 
                    text-align: center !important;
                    padding: 11px 17px !important;
                    height: auto !important;
                    
                }
                ul.slides {height: 170px !important}
                .button_link {
                    background: #999 !important;
                }
                .button_link_red {
                    background: #c00 !important;
                }
            </style>
        <![endif]-->
        <!--[if lt IE 8]>
            <link href="/css/tipTip_ie7.css" media="all" rel="stylesheet" type="text/css" />
        <![endif]-->
        <!--[if IE]>
            <style type="text/css">
                .clearfix {
                    zoom: 1;   /* triggers hasLayout */
                }   /* Only IE can see inside the conditional comment
                    and read this CSS rule. Don't ever use a normal HTML
                    comment inside the CC or it will close prematurely. */
            </style>
        <![endif]-->
        <!--Codice Statistiche Google Analytics Iury Mazzoni ## NON CANCELLARE!! ## -->

		<script type="text/plain" class="ce-script">
        
          var _gaq = _gaq || [];
          _gaq.push(['_setAccount', 'UA-320952-2']);
          _gaq.push(['_trackPageview']);
          _gaq.push(['_addTrans',
            '<%=IdOrdine%>',           // order ID - required
            'Cristalensi ITA',  // affiliation or store name
            '<%=TotaleCarrello%>',          // total - required
            '<%=CostoPagamento%>',           // tax
            '<%=CostoSpedizioneTotale%>',              // shipping
            '<%=Citta%>',       // city
            '<%=Provincia%>',     // state or province
            ''             // country
          ]);
        
           // add item might be called for every item in the shopping cart
           // where your ecommerce engine loops through each item in the cart and
           // prints out _addItem for each
           <%
           Set ars = Server.CreateObject("ADODB.Recordset")
           sql ="SELECT RigheOrdine.FkOrdine, RigheOrdine.Quantita, RigheOrdine.PrezzoProdotto, Prodotti.Titolo AS Prodotto, Prodotti.CodiceArticolo, Categorie2.Titolo AS Categoria2 "
           sql = sql + "FROM RigheOrdine INNER JOIN (Prodotti INNER JOIN Categorie2 ON Prodotti.FkCategoria2 = Categorie2.PkId) ON RigheOrdine.FkProdotto = Prodotti.PkId "
           sql = sql + "WHERE (((RigheOrdine.FkOrdine)="&IdOrdine&"))"
           ars.Open sql, conn, 1, 1
           if ars.recordcount>0 then
                Do While not ars.EOF
           %>
              _gaq.push(['_addItem',
                '<%=ars("FkOrdine")%>',           // order ID - required
                '<%=ars("CodiceArticolo")%>',           // SKU/code - required
                '<%=ars("Prodotto")%>',        // product name
                '<%=ars("Categoria2")%>',   // category or variation
                '<%=ars("PrezzoProdotto")%>',          // unit price - required
                '<%=ars("Quantita")%>'               // quantity - required
              ]);
          <%
              ars.movenext
              loop
          end if
          ars.close
          %>
          _gaq.push(['_trackTrans']); //submits transaction to the Analytics servers
        
          (function() {
            var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
            ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
            var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
          })();
        
        </script>
        <!--Codice Statistiche Google Analytics Iury Mazzoni ## NON CANCELLARE!! ## -->
    </head>
    <body>
        <div id="wrap">
            <!--#include file="inc_header.asp"-->

            <div id="main-content">
               
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <h3 style="font-size: 14px; display: inline; border: none;">Ordine n&deg; <%=idordine%> - Data <%=Left(DataAggiornamento, 10)%></h3>
                            <div class="carrello clearfix">
                                <%if FkPagamento=1 then%>
                                <p>
                                	<br><br>Grazie per aver scelto i nostri prodotti,<br>
                                    per completare l'ordine &egrave; necessario effettuare il bonifico con i seguenti dati:<br>
                                    <br>
                                    <strong>BANCA DI CREDITO COOPERATIVO DI CAMBIANO AG. MONTELUPO FIORENTINO<br>IBAN: IT33E0842537960000030277941<br>
                                    Codice BIC/SWIFT: CRACIT33</strong>
                                    <br><br>Nella causale indicare: "<strong>Ordine da sito internet n&deg; <%=idordine%></strong>"<br><br>
                                    Beneficiario:<br><strong>Cristalensi Snc di Lensi Massimiliano & C. (P.Iva e C.Fiscale 05305820481)<br>
                             Via arti e mestieri, 1 - 50056 Montelupo F.no (FI)</strong>
                                    <br><br>
                                  La merce verr&agrave; spedita al momento che la nostra banca ricever&agrave; il pagamento oppure, per velocizzare la spedizione, &egrave; possibile inviarci per email la ricevuta dell'avvenuto pagamento con bonifico (in caso di bonifico fatto con home banking spesso viene fornita dalla banca una ricevuta, oppure &egrave; possibile scannerizzare la ricevuta rilasciata dalla banca).<br>
                                  <br>
                                  Pagando, e quindi completando l'ordine, si accettano le condizioni di vendita.<br><br>
                                  Salva oppure stampa le condizioni di vendita (consultabili anche nell'apposita pagina del sito internet) da questo file (.pdf):<br><a href="/condizioni_di_vendita.pdf" target="_blank">condizion di vendita</a>
                                  <br>
                                  <br>
                                  Cordiali saluti, lo staff di Cristalensi
                                  <br>
                                  <br>
                                </p>
                                <%end if%>
                                <%if FkPagamento=4 then%>
                                <p>
                                	<br><br>Grazie per aver scelto i nostri prodotti,<br>
                                    per completare l'ordine &egrave; necessario effettuare il versamente sulla Carta di POSTEPAY con i seguenti dati:<br>
                                    <br><br>
                                    <strong>Beneficiario: LENSI GIULIANO - c.f. LNS GLN 42A30 D403J<br>
                                    Numero carta: 4023600915402442</strong>
                                    <br><br>Nella causale indicare: "<strong>Ordine da sito internet n&deg; <%=idordine%></strong>"<br><br>
                                    
                                  La merce verr&agrave; spedita al momento che riceveremo il pagamento oppure, per velocizzare la spedizione, &egrave; possibile inviarci per email la ricevuta dell'avvenuto pagamento.<br>
                                  <br>
                                  Pagando, e quindi completando l'ordine, si accettano le condizioni di vendita.<br><br>
                                  Salva oppure stampa le condizioni di vendita (consultabili anche nell'apposita pagina del sito internet) da questo file (.pdf):<br><a href="/condizioni_di_vendita.pdf" target="_blank">condizion di vendita</a>
                                  <br>
                                  <br>
                                  Cordiali saluti, lo staff di Cristalensi
                                  <br>
                                  <br>
                                </p>
                                <%end if%>
                      			<%if FkPagamento=3 then%>
                                <p>
                                	<br><br>Grazie per aver scelto i nostri prodotti,<br>
                                    la merce verr&agrave; spedita appena sar&agrave; disponibile nel nostro magazino.<br>
                                    Ti ricordiamo che per il pagamento in contrassegno, il corriere consegner&agrave; la merce solo se verr&agrave; pagata in contanti, non accetter&agrave; altri metodi di pagamento (anche gli assegni non saranno accettati).<br>
                                    <br>
                                  Pagando, e quindi completando l'ordine, si accettano le condizioni di vendita.<br>
                                  Salva oppure stampa le condizioni di vendita (consultabili nell'apposita pagina del sito internet) da questo file (.pdf):<br>
                                  <a href="/condizioni_di_vendita.pdf" target="_blank">condizion di vendita</a>
                                  <br>
                                  <br>
                                  Cordiali saluti, lo staff di Cristalensi
                                  <br>
                                  <br>
                                </p>
                                <%end if%>
                                <%if FkPagamento=2 then%>
								  <%
                                  TotaleGeneralePP=FormatNumber(TotaleGenerale,2)
                                  TotaleGeneralePP=Replace(TotaleGeneralePP, ".", "")
                                  TotaleGeneralePP=Replace(TotaleGeneralePP, ",", ".")
                                  %>
                                  <p>
                                  	
                                    <a href="https://www.paypal.com/it/webapps/mpp/paypal-popup" title="Come funziona PayPal" onClick="javascript:window.open('https://www.paypal.com/it/webapps/mpp/paypal-popup','WIPaypal','toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes, width=1060, height=700'); return false;"><img src="https://www.paypalobjects.com/webstatic/mktg/logo-center/logo_paypal_carte.jpg" border="0" style="float:right; padding-left:5px; width:319px; height:110px;" alt="Che cos'&egrave; PayPal"></a>Grazie per aver scelto i nostri prodotti,<br>
                                    per completare l'ordine &egrave; necessario effettuare il pagamento con i sistemi sicuri di Paypal che permettono di pagare con moltissime carte di credito e carte ricaribili:<br>
                                    MasterCard, Visa e Visa Electron, PostePay, Carta Aura e ricariche PayPal.<br><br>
                                    Pagando, e quindi completando l'ordine, si accettano le condizioni di vendita.<br>
                                  Salva oppure stampa le condizioni di vendita (consultabili anche nell'apposita pagina del sito internet) da questo file (.pdf):<br>
                                  <a href="/condizioni_di_vendita.pdf" target="_blank">condizion di vendita</a>							    <br>
                                    <br>
                                    </p>
    <form action="https://www.paypal.com/it/cgi-bin/webscr" method="post">
    <input type="hidden" name="cmd" value="_xclick">
    <input type="hidden" name="business" value="cristalensi@alice.it">
    <input type="hidden" name="item_name" value="Ordine n° <%=IdOrdine%>">
    <input type="hidden" name="item_number" value="<%=IdOrdine%>">
    <input type="hidden" name="currency_code" value="EUR">
    <input type="hidden" name="amount" value="<%=TotaleGeneralePP%>">
    <input type="hidden" name="return" value="http://www.cristalensi.it/pagamento_paypal_ok.asp">
    <input type="hidden" name="rm" value="2">
    <input type="hidden" name="cancel_return" value="http://www.cristalensi.it/pagamento_paypal_ko.asp">
    <input type="hidden" name="no_note" value="1">
    <input type="hidden" name="on0" value="Effettuato da">
    <input type="hidden" name="os0" value="<%=nome_log%>">
    <input type="hidden" name="cn" value="0">
    <input type="hidden" name="image_url" value="http://www.cristalensi.it/immagini/logo_cristalensi_piccolo.jpg">
    <input type="submit" name="paga_adesso" value="CLICCA QUI PER PAGARE CON IL SISTEMA SICURO DI PAYPAL" class="button" alt="Effettua i tuoi pagamenti con PayPal. E' un sistema rapido, gratuito e sicuro." style="height: 40px; font-size:14px; font-weight:bold;">
    </form>								
                                  <p>
                                  <br><br>
                                  La merce verr&agrave; spedita al momento che la nostra banca ricever&agrave; il pagamento.<br>
                                  <br>
                                  Cordiali saluti, lo staff di Cristalensi
                                  <br>
                                  <br>
                                  </p>
                                <%end if%>
                                	
                                
                                <p class="area clearfix"><span class="colonna articolo">[Codice articolo] Nome prodotto</span><span class="colonna quantita">quantit&agrave;</span><span class="colonna prezzo_unitario">prezzo unitario</span><span class="colonna prezzo_totale">prezzo totale</span></p>
                                <div class="data">
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT PkId, FkOrdine, FkProdotto, PrezzoProdotto, Quantita, TotaleRiga, Titolo, CodiceArticolo, Colore, Lampadina FROM RigheOrdine WHERE FkOrdine="&idOrdine&""
	rs.Open sql, conn, 1, 1
	num_prodotti_carrello=rs.recordcount
	
%>                                    
									<%if rs.recordcount>0 then%>
                                        
                                        <%
                                        Do while not rs.EOF
                                        %>					
    
                                        <p class="riga"><span class="colonna articolo">[<%=rs("codicearticolo")%>]&nbsp;<strong><%=rs("titolo")%></strong><%if Len(rs("colore"))>0 or Len(rs("lampadina"))>0 then%><br /><%if Len(rs("colore"))>0 then%>&nbsp;Col.:&nbsp;<%=rs("colore")%><%end if%><%if Len(rs("lampadina"))>0 then%>&nbsp;-&nbsp;Lamp.:&nbsp;<%=rs("lampadina")%><%end if%><%end if%></span>
                                        <%
                                        quantita=rs("quantita")
                                        if quantita="" then quantita=1
                                        %>
                                        <span class="colonna quantita"><%=quantita%> pezzi </span><span class="colonna prezzo_unitario"><%=FormatNumber(rs("PrezzoProdotto"),2)%>&#8364;</span><span class="colonna prezzo_totale"><%=FormatNumber(rs("TotaleRiga"),2)%>&#8364;</span></p>
                                        <%
                                        rs.movenext
                                        loop
                                        %>
<%
	rs.close
%>                                        
                                    <%end if%>
                                </div>
                                
                                <p class="area clearfix"><span class="colonna descrizione">Modalit&agrave; di spedizione</span><span class="colonna prezzo_unitario">&nbsp;</span><span class="colonna prezzo_totale">Totale</span></p>
                                <div class="data">
                                    <p class="riga">
                                    <span class="colonna descrizione"><b><%=TipoTrasporto%></b></span>
                                    <span class="colonna prezzo_unitario">&nbsp;</span>
                                    <span class="colonna prezzo_totale"><%=FormatNumber(CostoSpedizioneTotale,2)%>&#8364;</span>
                                    </p>
                                    <p>&nbsp;</p>
                                    <h3 style="font-size:12px;">Riferimenti per l'indirizzo di spedizione</h3>
                                    <p><%=DatiSpedizione%></p>
                                    <p>&nbsp;</p>
                                    <h3 style="font-size:12px;">Eventuali annotazioni</h3>
                                    <p><%=NoteCliente%></p>
                                </div>
                                
                                <p class="area clearfix"><span class="colonna descrizione">Modalit&agrave; di pagamento</span><span class="colonna prezzo_unitario">&nbsp;</span><span class="colonna prezzo_totale">Totale</span></p>
                                <div class="data">
                                    <p class="riga">
                                    <span class="colonna descrizione"><b><%=TipoPagamento%></b></span>
                                    <span class="colonna prezzo_unitario">&nbsp;</span>
                                    <span class="colonna prezzo_totale"><%=FormatNumber(CostoPagamento,2)%>&#8364;</span>
                                    </p>
                                    <p>&nbsp;</p>
                                    <h3 style="font-size:12px;">Riferimenti per i dati di fatturazione:</h3>
                                    <div class="clearfix">
                                        <div class="table">
                                            <div class="tr">
                                                <div class="td">
                                                    <%if Rag_Soc<>"" then%><%=Rag_Soc%>&nbsp;&nbsp;<%end if%><%if nominativo<>"" then%><%=nominativo%><%end if%>
                                                </div>
                                            </div>
                                            <div class="tr">
                                                <div class="td">
                                                    Codice fiscale: <%=Cod_Fisc%><%if PartitaIVA<>"" then%> - Partita IVA: <%=PartitaIVA%><%end if%>
                                                </div>
                                            </div>
                                            <div class="tr">
                                                <div class="td">
                                                    <%=indirizzo%>
                                                </div>
                                            </div>
                                            <div class="tr">
                                                <div class="td">
                                                    <%=cap%> - <%=citta%> (<%=provincia%>)
                                                </div>
                                            </div>
                                        </div>
                                	</div> 
                                </div>
                                
                                  <h4 class="cart clearfix">
                                  <span class="total_price">Totale ordine:&nbsp;
								  <%if TotaleGenerale<>0 then%>
                                  <%=FormatNumber(TotaleGenerale,2)%>
                                  <%else%>
                                  0,00
                                  <%end if%>
                                  &#8364;
                                  </span>
                                  </h4>
                                    <form method="post" name="modulo" id="modulo" action="/stampa_ordine.asp" target="_blank">
                                    <input type="hidden" name="idordine" id="idordine" value="<%=idordine%>">
                                    <%if FkPagamento=1 or FkPagamento=3 or FkPagamento=4 then%><button type="submit" name="stampa" style="float:right;" class="button_link_red">stampa l'ordine</button><%end if%>
                                    </form>
                            </div>
                        </div>
                    </div>
                </div>
                <!--#include file="inc_sx_prodotti.asp"-->
            </div>
        </div>
         <!--#include file="inc_footer.asp"-->
    </body>
</html>
<!--#include file="inc_strClose.asp"-->