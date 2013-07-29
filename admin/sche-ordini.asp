<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
<%
	pkid = request("pkid")
	if pkid = "" then pkid = 0
	
	p = request("p")
	if p = "" then p = 1
	ordine = request("ordine")
	if ordine = "" then ordine = 0
%>
<%
	mode = request("mode")
	if mode = "" then mode = 0
	if mode=1 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Ordini where pkid="&pkid
		rs.Open sql, conn, 3, 3
		
		email=request("email")
		nominativo=request("nominativo")
		nome=request("nome")
		
		stato=request("stato")
		rs("stato")=stato
		
		InfoSpedizione=request("InfoSpedizione")
		rs("InfoSpedizione")=InfoSpedizione
		
		NoteCri=request("NoteCri")
		rs("NoteCri")=NoteCri
		rs("DataAggiornamento")=now()
		
		italia=request("italia")
		if italia="No" and stato="22" then
			CostoSpedizioneTotale=request("CostoSpedizioneTotale")
			rs("CostoSpedizione")=CostoSpedizioneTotale
		end if
		
		if request("C1")<>"ON" and italia="No" and stato="22" then
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Spett.le "&nome&" "&nominativo&", come da lei richiesto è stato aggiunto il costo di spedizione per i prodotti da lei ordinati con l'Ordine da sito internet n° "&pkid&".<br><br>Costo di spedizione: "&CostoSpedizioneTotale&"&#8364;<br><br>Adesso potrà completare l'ordine semplicemente andando sulla <a href=""http://www.cristalensi.it"">Home Page</a>, inserendo Login (Email) e Password nell'Area Clienti per farsi riconoscere, successivamente cliccando su ""I tuoi ordini"": in quella pagina troverà il suo ordine e cliccandoci potrà terminarlo scegliendo la modalità di pagamento.<br><br>Per qualsiasi chiarimento o informazione ci contatti.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td><br><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dear "&nome&" "&nominativo&", as you requested, mailing costs have been added to the cost of the products which you ordered with Order from web site n° "&pkid&".<br><br>The mailing costs are: "&CostoSpedizioneTotale&"&#8364;<br><br>You may now complete the order simply by going to the <a href=""http://www.cristalensi.it"">Home Page</a>, inserting the Login (Email) and Password in the Client Area to identify yourself, and then clicking on ""Your orders"": on that page you will find your order, and clicking on it you can complete it choosing the payment method.<br><br>Please contact us with any questions you may have.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best regards, from the staff of Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Cristalensi.it: Costi di spedizione inseriti per l'ordine n. "&pkid&" - Mailing costs inserted for order n. "&pkid&""
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Spett.le "&nome&" "&nominativo&", come da lei richiesto è stato aggiunto il costo di spedizione per i prodotti da lei ordinati con l'Ordine da sito internet n° "&pkid&".<br><br>Costo di spedizione: "&CostoSpedizioneTotale&"&#8364;<br><br>Adesso potrà completare l'ordine semplicemente andando sulla <a href=""http://www.cristalensi.it"">Home Page</a>, inserendo Login (Email) e Password nell'Area Clienti per farsi riconoscere, successivamente cliccando su ""I tuoi ordini"": in quella pagina troverà il suo ordine e cliccandoci potrà terminarlo scegliendo la modalità di pagamento.<br><br>Per qualsiasi chiarimento o informazione ci contatti.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td><br><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dear "&nome&" "&nominativo&", as you requested, mailing costs have been added to the cost of the products which you ordered with Order from web site n° "&pkid&".<br><br>The mailing costs are: "&CostoSpedizioneTotale&"&#8364;<br><br>You may now complete the order simply by going to the <a href=""http://www.cristalensi.it"">Home Page</a>, inserting the Login (Email) and Password in the Client Area to identify yourself, and then clicking on ""Your orders"": on that page you will find your order, and clicking on it you can complete it choosing the payment method.<br><br>Please contact us with any questions you may have.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best regards, from the staff of Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Cristalensi.it: Costi di spedizione inseriti per l'ordine n. "&pkid&" - Mailing costs inserted for order n. "&pkid&""
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set eMail_cdo = Nothing
			
			'invio al webmaster
			
			Mittente = "info@cristalensi.it"
			Destinatario = "iurymazzoni@hotmail.com"
			Oggetto = "Cristalensi.it: Costi di spedizione inseriti per l'ordine n. "&pkid&" - Mailing costs inserted for order n. "&pkid&""
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set eMail_cdo = Nothing
		end if
		
		'ordine in lavorazione
		if request("C1")<>"ON" and stato="7" then
			
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Spett.le "&nome&" "&nominativo&", l'Ordine da sito internet n° "&pkid&" è stato preso in carico dal nostro staff.<br>Appena sarà spedito riceverà un'email con i dati di spedizione: nome del corriere e codice identificativo.<br><br><b>Al momento del ricevimento della merce, firmare e scrivere la dicitura &quot;RISERVA DI CONTROLLO&quot; sulla CEDOLINA del corriere, tutto ci&ograve; per avere copertura assicurativa nel caso in cui siano presenti prodotti danneggiati.</b><br><br>Per qualsiasi chiarimento o informazione ci contatti.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			if italia="No" then
				HTML1 = HTML1 & "<tr>"
				HTML1 = HTML1 & "<td><br><br>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"
				HTML1 = HTML1 & "<tr>"
				HTML1 = HTML1 & "<td>"
				HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dear "&nome&" "&nominativo&", the order from internet site n° "&pkid&" is being processed by our staff.<br> As soon as it has been sent you will receive an e-mail with the mailing data:  name of the courier and identification code.<br><br><b>When you receive the goods please write “SUBJECT TO CONTROL” on the COURIER'S RECIEPT SLIP, all of which will help to have better coverage should any goods have been damaged.</b></font><br>"
				HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best regards, from the staff of Cristalensi</font>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"
			end if
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@cristalensi.it"
			Destinatario = email
			if italia="No" then
				Oggetto = "Aggiornamento ordine n. "&pkid&" effettuato su Cristalensi.it - Update of order n. "&pkid&" from Cristalensi.it"
			else
				Oggetto = "Aggiornamento ordine n. "&pkid&" effettuato su Cristalensi.it"
			end if
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set eMail_cdo = Nothing
			
			'fine invio email
						
			'invio al webmaster
			
			Mittente = "info@cristalensi.it"
			Destinatario = "iurymazzoni@hotmail.com"
			Oggetto = "Aggiornamento ordine n. "&pkid&" effettuato su Cristalensi.it"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set eMail_cdo = Nothing
			
			'fine invio email	
		
		end if
		
		'prodotti spediti - dati spedizione
		if request("C1")<>"ON" and stato="8" then
			
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Spett.le "&nome&" "&nominativo&", i prodotti da lei ordinati con l'Ordine da sito internet n° "&pkid&" sono stati spediti secondo le modalità richieste.<br><br>"
			HTML1 = HTML1 & "<b>LEGGERE ATTENTAMENTE:<br>Al momento del ricevimento della merce, firmare e scrivere la dicitura &quot;RISERVA DI CONTROLLO&quot; sulla CEDOLINA del corriere, tutto ci&ograve; per avere copertura assicurativa nel caso in cui siano presenti prodotti danneggiati.</b><br><br>"
			HTML1 = HTML1 & "Note sulla spedizione:<br>"&InfoSpedizione&"<br><br>"
			if Left(NoteCri,4)="http" then
			HTML1 = HTML1 & "<b><a href="""&NoteCri&""">"&NoteCri&"</a></b><br><br>"
			end if
			HTML1 = HTML1 & "<br><br>Per qualsiasi chiarimento o informazione ci contatti.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			if italia="No" then
				HTML1 = HTML1 & "<tr>"
				HTML1 = HTML1 & "<td><br><br>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"
				HTML1 = HTML1 & "<tr>"
				HTML1 = HTML1 & "<td>"
				HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dear "&nominativo&", the products which you ordered with  the Order from internet site n° "&pkid&" 000 have been sent in the mode which you specified.<br><br>"
				HTML1 = HTML1 & "<b>TO READ CAREFULLY:<br>When you receive the goods please write “SUBJECT TO CONTROL” on the COURIER'S RECIEPT SLIP, all of which will help to have better coverage should any goods have been damaged.</b></font><br><br>"
				HTML1 = HTML1 & "Note on the consignment:<br>"&InfoSpedizione&"<br><br>"
				if Left(NoteCri,4)="http" then
					HTML1 = HTML1 & "<b><a href="""&NoteCri&""">"&NoteCri&"</a></b><br><br>"
				end if
				HTML1 = HTML1 & "Please contact us with any questions.</font><br>"
				HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best regards, from the staff of Cristalensi</font>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"
			end if
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@cristalensi.it"
			Destinatario = email
			if italia="No" then
				Oggetto = "Conferma spedizione ordine n "&pkid&" da Cristalensi.it - Confirmation of shipment order n. "&pkid&" from Cristalensi.it"
			else
				Oggetto = "Conferma spedizione ordine n "&pkid&" da Cristalensi.it"
			end if
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Spett.le "&nome&" "&nominativo&", i prodotti da lei ordinati con l'Ordine da sito internet n° "&pkid&" sono stati spediti secondo le modalità richieste.<br><br>"
			HTML1 = HTML1 & "<b>LEGGERE ATTENTAMENTE:<br>Al momento del ricevimento della merce, firmare e scrivere la dicitura &quot;RISERVA DI CONTROLLO&quot; sulla CEDOLINA del corriere, tutto ci&ograve; per avere copertura assicurativa nel caso in cui siano presenti prodotti danneggiati.</b><br><br>"
			HTML1 = HTML1 & "Note sulla spedizione:<br>"&InfoSpedizione&"<br><br>"
			if Left(NoteCri,4)="http" then
			HTML1 = HTML1 & "<b><a href="""&NoteCri&""">"&NoteCri&"</a></b><br><br>"
			end if
			HTML1 = HTML1 & "<br><br>Per qualsiasi chiarimento o informazione ci contatti.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			if italia="No" then
				HTML1 = HTML1 & "<tr>"
				HTML1 = HTML1 & "<td><br><br>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"
				HTML1 = HTML1 & "<tr>"
				HTML1 = HTML1 & "<td>"
				HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dear "&nome&" "&nominativo&", the products which you ordered with  the Order from internet site n° "&pkid&" 000 have been sent in the mode which you specified.<br><br>"
				HTML1 = HTML1 & "<b>TO READ CAREFULLY:<br>When you receive the goods please write “SUBJECT TO CONTROL” on the COURIER'S RECIEPT SLIP, all of which will help to have better coverage should any goods have been damaged.</b></font><br><br>"
				HTML1 = HTML1 & "Note on the consignment:<br>"&InfoSpedizione&"<br><br>"
				if Left(NoteCri,4)="http" then
					HTML1 = HTML1 & "<b><a href="""&NoteCri&""">"&NoteCri&"</a></b><br><br>"
				end if
				HTML1 = HTML1 & "Please contact us with any questions.</font><br>"
				HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best regards, from the staff of Cristalensi</font>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"
			end if
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			if italia="No" then
				Oggetto = "Conferma spedizione ordine n "&pkid&" da Cristalensi.it - Confirmation of shipment order n. "&pkid&" from Cristalensi.it"
			else
				Oggetto = "Conferma spedizione ordine n "&pkid&" da Cristalensi.it"
			end if
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set eMail_cdo = Nothing
			
			'invio al webmaster
			
			Mittente = "info@cristalensi.it"
			Destinatario = "iurymazzoni@hotmail.com"
			if italia="No" then
				Oggetto = "Conferma spedizione ordine n "&pkid&" da Cristalensi.it - Confirmation of shipment order n. "&pkid&" from Cristalensi.it"
			else
				Oggetto = "Conferma spedizione ordine n "&pkid&" da Cristalensi.it"
			end if
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set eMail_cdo = Nothing
			
			'fine invio email	
		
		end if
		
		if request("C1") = "ON" then
			
			'qui devono essere inserite tutte le tabelle dove compare FkOrdine per cancellare il record oppure metterlo a 0
			Set ss=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From RigheOrdine where FkOrdine="&pkid&""
			ss.Open sql, conn, 3, 3
				if ss.recordcount>0 then
					Do while not ss.EOF
						ss.update
						ss.delete
					ss.movenext
					loop
				end if
			ss.close
			
			rs.delete
		end if
		rs.update
		
		rs.close
	end if
%>
<html>
<head>
<title>Cristalensi Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0" class="TAB_centrale">
  <!--#include file="testata.asp"-->
  <tr>
    <td height="30" colspan="2" valign="middle"><table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="159" class="menu-celle">&nbsp;Menu</td>
          <td width="267" class="menu-celle">Gestione ordini</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-ordini.asp">Elenco ordini &raquo;</a></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td colspan="2" valign="top"><table width="750" height="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="150" class="admin-menu" valign="top">
		<!--#include file="sinistra.asp"-->
		 </td>
        <td align="center" valign="top">
          <!--tab centrale-->
			<% if request("C1") <> "ON" then %>
                <% if mode = 1 and pkid = 0 then %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Record Inserito ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-ordini.asp?ordine=<%=ordine%>";
			   					}
							//-->
							</script>
                <% else %>
                <% if mode = 1 then %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Record Aggiornato ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
								<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-ordini.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
								//-->
								</script>
                <% else %>
				<table cellpadding="0" cellspacing="0" border="0" width="98%" class="admin-righe">
					<tr align="left">
                    <td height="15" colspan="2">&nbsp;</td>
                  </tr>
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	if pkid<12210 then
		sql = "SELECT RigheOrdine.PkId, RigheOrdine.FkOrdine, RigheOrdine.PrezzoProdotto as PrezzoProdotto, RigheOrdine.FkProdotto, RigheOrdine.Quantita, RigheOrdine.TotaleRiga, Prodotti.Titolo, Prodotti.CodiceArticolo, RigheOrdine.Colore FROM Prodotti INNER JOIN RigheOrdine ON Prodotti.PkId = RigheOrdine.FkProdotto WHERE (((RigheOrdine.FkOrdine)="&pkid&"))"
	else
		sql = "SELECT PkId, FkOrdine, FkProdotto, PrezzoProdotto, Quantita, TotaleRiga, Titolo, CodiceArticolo, Colore FROM RigheOrdine WHERE FkOrdine="&pkid&""
	end if
	rs.Open sql, conn, 1, 1
	num_prodotti_carrello=rs.recordcount
	
	Set ss = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Ordini.*, Clienti.* FROM Clienti RIGHT JOIN Ordini ON Clienti.PkId = Ordini.FkCliente where Ordini.pkid="&pkid
	ss.Open sql, conn, 1, 1
	
	italia=ss("Italia")
	if italia="" then italia="Sì"
%>
					<form method="post" action="sche-ordini.asp?mode=1&pkid=<%=pkid%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                  <input type="hidden" name="email" value="<%=ss("Email")%>">
				  <input type="hidden" name="nominativo" value="<%=ss("nominativo")%>">
                  <input type="hidden" name="nome" value="<%=ss("nome")%>">
				  <input type="hidden" name="italia" value="<%=italia%>">
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Cliente</strong></td>
                  </tr>
<%if ss.recordcount>0 then%>				  
				  <tr align="left">
                    <td width="52%" height="15">Nominativo: 
                    <%if ss("nominativo")<>"" then%><%=ss("nome")%>&nbsp;<%=ss("nominativo")%><%else%>Non iscritto<%end if%></td>
					<td width="48%" height="15">Rag.Soc.: <%=ss("Rag_Soc")%></td>
                  </tr>
                  <tr align="left">
                    <td height="15">Cod.Fisc.: <%=ss("Cod_Fisc")%></td>
					<td height="15">Partita IVA: <%=ss("PartitaIVA")%></td>
                  </tr>
				  <tr align="left">
                    <td height="15">Indirizzo: <%=ss("Indirizzo")%></td>
					<td height="15">CAP: <%=ss("CAP")%></td>
                  </tr>
				  <tr align="left">
                    <td height="15">Citta': <%=ss("Citta")%></td>
					<td height="15">Provincia: <%=ss("Provincia")%></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">Nazione: <%if ss("Italia")="Sì" then%>Italia<%else%><%=ss("NazioneDiversa")%><%end if%></td>
                  </tr>
				  <tr align="left">
                    <td height="15">Telefono: <%=ss("Telefono")%></td>
					<td height="15">Fax:<%=ss("Fax")%> </td>
                  </tr>
				  <tr align="left">
                    <td height="15">Email: <%=ss("Email")%></td>
					<td height="15">Data: <%=ss("Data")%></td>
                  </tr>
<%end if%>				  
				  
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Ordine</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">
					
					<table width="100%"  border="0" cellpadding="0" cellspacing="0">
                
				<tr class="menu-celle">
                  <td height="20" align="left">&nbsp;[codice articolo] nome prodotto</td>
				  <td height="20" align="right">quantità</td>
				  <td height="20" align="right">prezzo unitario</td>
				  <td width="131" height="20" align="right">prezzo totale</td>
				  </tr>
				<tr>
                <td colspan="4" align="left">&nbsp;</td>
                </tr>
<%
if rs.recordcount>0 then
%>
<%
Do while not rs.EOF
%>					
				  <tr class="admin-righe">
                  <td align="left" width="341">
				  [<%=rs("codicearticolo")%>]&nbsp;<%=rs("titolo")%><%if Len(rs("colore"))>0 then%>&nbsp;(<%=rs("colore")%>)<%end if%>
				  </td>
                  <td align="right" width="89">
				  <%
				  quantita=rs("quantita")
				  if quantita="" then quantita=1
				  %>
				  <%=quantita%> pezzi </td>
                  <td align="right" width="119"><%=FormatNumber(rs("PrezzoProdotto"),2)%>€</td>
                  <td align="right"><%=FormatNumber(rs("TotaleRiga"),2)%>€</td>
                  </tr>
                  <tr>
                  <td colspan="4" align="left" class="divisione-elenco"><img src="immagini/spacer.gif" height="10"></td>
                  </tr>
				  <tr>
                  <td colspan="4" align="left"><img src="immagini/spacer.gif" height="10"></td>
                  </tr>
<%
conta=conta+1
rs.movenext
loop
%>				
<%end if%>
              </table>
<%if ss.recordcount>0 then%>
<%
	TotaleCarrello=ss("TotaleCarrello")
	CostoSpedizioneTotale=ss("CostoSpedizione")
	TipoTrasporto=ss("TipoTrasporto")
	DatiSpedizione=ss("DatiSpedizione")
	NoteCliente=ss("NoteCliente")
	
	FkPagamento=ss("FkPagamento")
	TipoPagamento=ss("TipoPagamento")
	CostoPagamento=ss("CostoPagamento")
	
	Nominativo=ss("Nominativo")
	Nome=ss("Nome")
	Rag_Soc=ss("Rag_Soc")
	Cod_Fisc=ss("Cod_Fisc")
	PartitaIVA=ss("PartitaIVA")
	Indirizzo=ss("Indirizzo")
	Citta=ss("Citta")
	Provincia=ss("Provincia")
	CAP=ss("CAP")
	
	TotaleGenerale=ss("TotaleGenerale")
	
	DataAggiornamento=ss("DataAggiornamento")
%>
			  <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="admin-righe">
				<tr>
                  <td colspan="2" align="left">&nbsp;</td>
                </tr>
				<tr class="menu-celle">
                  <td align="left">Modalità di spedizione</td>
				  <td align="right">Totale</td>
                </tr>
				<tr>
                  <td colspan="2" align="left">&nbsp;</td>
                </tr>
				<tr>
                  <td align="left"><%=TipoTrasporto%></td>
				  <td align="right"><%if italia="No" and ss("Stato")=12 then%><input type="text" name="CostoSpedizioneTotale" value="<%=CostoSpedizioneTotale%>" size="5" class="form"><%else%><%=FormatNumber(CostoSpedizioneTotale,2)%><input type="hidden" name="CostoSpedizioneTotale" value="<%=CostoSpedizioneTotale%>"><%end if%>€</td>
                </tr>
				<tr>
                  <td colspan="2" align="left">&nbsp;</td>
                </tr>
				<tr>
                  <td colspan="2" align="left"><b>Riferimenti per l'indirizzo di spedizione:</b><br><%=DatiSpedizione%></td>
                </tr>
				<tr>
                  <td colspan="2" align="left">&nbsp;</td>
                </tr>
				<tr>
                  <td colspan="2" align="left"><b>Eventuali annotazioni:</b><br><%=NoteCliente%></td>
                </tr>
				<tr>
                  <td colspan="2" align="left">&nbsp;</td>
                </tr>
			  </table>
			  
			  <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="admin-righe">
				<tr class="menu-celle">
				<td width="89%" height="20" align="left">&nbsp;Modalit&agrave; di Pagamento </td>
				<td width="11%" height="20" align="right">totale</td>
				</tr>
				<tr>
				<td colspan="2" align="left">&nbsp;</td>
				</tr>
				<tr>
			    <td align="left" valign="top">
			    <b><%=TipoPagamento%></b>				  				  </td>
			    <td align="right" valign="top"><%=FormatNumber(CostoPagamento,2)%>€</td>
			    </tr>
				<tr>
                <td colspan="2" align="left">&nbsp;</td>
                </tr>
				<tr>
                <td height="25" colspan="2" align="left"><b>Riferimenti per i dati di fatturazione: </b></td>
                </tr>
				<tr>
                <td colspan="2" align="center">
				<table width="95%" border="0" cellpadding="0" cellspacing="0" class="admin-righe">
                    <tr align="left">
                      <td height="20" colspan="2"><%if Rag_Soc<>"" then%><%=Rag_Soc%>&nbsp;&nbsp;<%end if%><%if nominativo<>"" then%><%=nome%>&nbsp;<%=nominativo%><%end if%></td>
                      </tr>
                    <tr align="left">
                      <td height="20" colspan="2">Codice fiscale: <%=Cod_Fisc%><%if PartitaIVA<>"" then%> - Partita IVA: <%=PartitaIVA%><%end if%></td>
					  </tr>
                    <tr align="left">
                      <td height="20" colspan="2"><%=indirizzo%></td>
                      </tr>
                    <tr align="left">
                      <td height="20" colspan="2"><%=cap%> - <%=citta%> (<%=provincia%>)</td>
                      </tr>
				</table>				</td>
                </tr>
				<tr>
                <td colspan="2" align="left">&nbsp;</td>
                </tr>
			  </table>
			  
			  <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="prodotto">
				<tr>
                  <td colspan="2" align="left">&nbsp;</td>
                </tr>
				<tr>
                  <td height="25" colspan="2" align="right" class="menu-celle"><strong> Totale ordine:&nbsp;
                      <%if TotaleGenerale<>0 then%>
                      <%=FormatNumber(TotaleGenerale,2)%>
                      <%else%>
                      0,00
                      <%end if%>
                      €&nbsp;</strong></td>
                </tr>
				<tr>
                  <td colspan="2" align="left">&nbsp;</td>
                </tr>
				<tr class="admin-righe" height="20">
                  <td width="48%" align="left"><strong>Data ordine:</strong> <%=ss("DataOrdine")%></td>
                  <td width="52%" align="right"><strong>Data aggiornamento:</strong> <%=ss("DataAggiornamento")%></td>
				</tr>
				<tr class="admin-righe" height="30">
                  <td colspan="2" align="left"><strong>Stato ordine:</strong><br>
				  iniziato<input type="radio" name="stato" value="0" <%if ss("Stato")=0 then%>checked="checked"<%end if%>>
				  &nbsp;&nbsp;
				  assegnato a un cliente<input type="radio" name="stato" value="1" <%if ss("Stato")=1 then%>checked="checked"<%end if%>>
				  &nbsp;&nbsp;
				  fase spedizione<input type="radio" name="stato" value="2" <%if ss("Stato")=2 then%>checked="checked"<%end if%>>
				  &nbsp;&nbsp;
				  fase pagamento<input type="radio" name="stato" value="3" <%if ss("Stato")=3 then%>checked="checked"<%end if%>>&nbsp;&nbsp;<br>
				  fase spedizione intern.<input type="radio" name="stato" value="12" <%if ss("Stato")=12 then%>checked="checked"<%end if%>>
				  &nbsp;&nbsp;
				  fase pagamento intern.<input type="radio" name="stato" value="22" <%if ss("Stato")=22 then%>checked="checked"<%end if%>>
				  <br>
				  in pagamento/annullato<input type="radio" name="stato" value="6" <%if ss("Stato")=6 then%>checked="checked"<%end if%>>
				  &nbsp;&nbsp;
				  pagato con PP<input type="radio" name="stato" value="4" <%if ss("Stato")=4 then%>checked="checked"<%end if%>>
				  &nbsp;&nbsp;
				  annullato PP<input type="radio" name="stato" value="5" <%if ss("Stato")=5 then%>checked="checked"<%end if%>>
				  <br>
				  in lavorazione<input type="radio" name="stato" value="7" <%if ss("Stato")=7 then%>checked="checked"<%end if%>>
				  &nbsp;&nbsp;
				  spedito<input type="radio" name="stato" value="8" <%if ss("Stato")=8 then%>checked="checked"<%end if%>>
				  &nbsp;&nbsp;corriere e codice:
				  <input type="text" name="InfoSpedizione" value="<%=ss("InfoSpedizione")%>" size="40" class="form" >				  </td>
				</tr>
				<tr>
                  <td colspan="2" align="left">&nbsp;</td>
                </tr>
				<tr class="admin-righe">
                  <td colspan="2" align="left">Note riservate sull'ordine: <input type="text" name="NoteCri" value="<%=ss("NoteCri")%>" size="40" class="form" ></td>
                </tr>
				<tr>
                  <td colspan="2" align="left">&nbsp;</td>
                </tr>
				<tr class="admin-righe">
                  <td align="left" height="30"><a href="../stampa_ordine.asp?IdOrdine=<%=PkId%>" target="_blank">Stampa l'ordine</a></td>
				  <td align="right" height="30"><input name="Submit" type="submit" class="form" value="Aggiorna" align="absmiddle">&nbsp; <input type="checkbox" name="C1" value="ON" > 
                          &nbsp; Per cancellare l'ordine</td>
                  
				</tr>
              </table>
			  <br>
<%end if%>
<%
ss.close
rs.close
%>         					
					</td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                </form>
				</table>
				<% end if %>
                <% end if %>
                <% else %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Record Cancellato ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-ordini.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
							//-->
						</script>
                <% end if %>
			<!--fine tab-->
		  </td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<!--#include file="strClose.asp"-->
<!--#include file="chiusura.asp"-->