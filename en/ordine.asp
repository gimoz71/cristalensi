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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Thank you "&nominativo_email&" for having chosen our products!<br>This e-mail confirms the completion of order n° "&idordine&".<br><br><b>TOTAL ORDER: <u>"&TotaleGenerale&"&#8364;</u></b><br><br> To complete the order you must make a bank transfer with the following details:<br><u>BANCA DI CREDITO COOPERATIVO DI CAMBIANO AG. MONTELUPO FIORENTINO</u><br>IBAN: <u>IT33E0842537960000030277941</u><br>Code BIC/SWIFT: <u>CRACIT33</u><br>As the cause of payment please indicate: Order from web site n° "&idordine&"<br><br>Beneficiary: Cristalensi Snc di Lensi Massimiliano & C. (P.Iva e C.Fiscale 05305820481)<br>Via arti e mestieri, 1 - 50056 Montelupo F.no (FI)<br><br><br>Our staff will send the merchandise as soon as the bank is notified of the bank transfer or, to speed-up the consignment, send an e-mail with the bank transfer reciept (in the case of home banking a bank transfer receipt is often provided by the bank, or you could scan the receipt left by the bank).</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best wishes from the staff of Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Confirmation dispatch order n "&idordine&" to Cristalensi.it with payment by bank transfer"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento con bonifico dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it (sito inglese) con pagamento con bonifico bancario"
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
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it (sito inglese) con pagamento con bonifico bancario"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Thank you "&nominativo_email&" for having chosen our products!<br>This e-mail is confirmation of the completion of order n° "&idordine&".<br><br><br>.  It will be the care of our staff to send you the merchandise as soon as it is available in our stock room.<br>We remind you that in the case of payment on receipt of goods, the courier will consign the merchandise only if paid in cash, other types of payment will not be accepted (even checks will not be accepted).</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best wishes from the staff of Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Confirmation dispatch order n "&idordine&" a Cristalensi.it with payment on receipt of goods"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento in contrassegno dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it (sito inglese) con pagamento in contrassegno"
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

			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it (sito inglese) con pagamento in contrassegno"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Thank you "&nominativo_email&" for having chosen our products!<br>This e-mail confirms the completion of order n° "&idordine&".<br><br><b>TOTAL ORDER: <u>"&TotaleGenerale&"&#8364;</u></b><br><br> To complete the order you must make a bank transfer with the following details:<br><u>BANCA DI CREDITO COOPERATIVO DI CAMBIANO AG. MONTELUPO FIORENTINO</u><br>IBAN: <u>IT33E0842537960000030277941</u><br>Code BIC/SWIFT: <u>CRACIT33</u><br>As the cause of payment please indicate: Order from web site n° "&idordine&"<br><br>Beneficiary: Cristalensi Snc di Lensi Massimiliano & C. (P.Iva e C.Fiscale 05305820481)<br>Via arti e mestieri, 1 - 50056 Montelupo F.no (FI)<br><br><br>Our staff will send the merchandise as soon as the bank is notified of the bank transfer or, to speed-up the consignment, send an e-mail with the bank transfer reciept (in the case of home banking a bank transfer receipt is often provided by the bank, or you could scan the receipt left by the bank).</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best wishes from the staff of Cristalensi</font>"
			
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Thank you "&nominativo_email&" for having chosen our products!<br>This e-mail confirms the completion of order n° "&idordine&".<br><br><b>TOTAL ORDER: <u>"&TotaleGenerale&"&#8364;</u></b><br><br> To complete the order you must make a card PostePay transfer with the following details:<br><u><strong>Beneficiary: LENSI GIULIANO - c.f. LNS GLN 42A30 D403J<br>Card Number: 4023 6005 5507 0285</strong><br><br>As the cause of payment please indicate: Order from Cristalensi web site n° "&idordine&"<br><br>Our staff will send the merchandise as soon as the bank is notified of the bank transfer or, to speed-up the consignment, send an e-mail with the bank transfer reciept.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best wishes from the staff of Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Confirmation dispatch order n "&idordine&" to Cristalensi.it with payment by PostePay card"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento con POSTEPAY dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it (sito inglese) con pagamento con POSTEPAY"
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
			Oggetto = "Conferma invio ordine n "&idordine&" a Cristalensi.it (sito inglese) con pagamento con POSTEPAY"
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
%>
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Cristalensi - Order</title>
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="/js/media-queries-ie.js"></script>
        <![endif]-->
        <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
        <script src="/js/jquery.blueberry.js"></script>
        <script src="/js/jquery.tipTip.js"></script>
        <link href="/css/css.css" rel="stylesheet" type="text/css">
        <link href="/css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="/css/tipTip.css" rel="stylesheet" type="text/css">
        <style type="text/css">
            .clearfix:after {
                content: ".";
                display: block;
                height: 0;
                clear: both;
                visibility: hidden;
            }
        </style>
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

		<script type="text/javascript">
        
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
                            <h3 style="font-size: 14px; display: inline; border: none;">Order n&deg; <%=idordine%> - Date <%=Left(DataAggiornamento, 10)%></h3>
                            <div class="carrello clearfix">
                                <%if FkPagamento=1 then%>
                                <p>
                                	<br>
                                  	<br>
                                    Thank you for having chosen our products,<br>
                                    to complete the order you must make out a bank transfer to the following address:<br>
                                    <br>
                                    <strong><u>BANCA DI CREDITO COOPERATIVO DI CAMBIANO AG. MONTELUPO FIORENTINO</u><br>IBAN: <u>IT33E0842537960000030277941</u><br>
                                    Codice BIC/SWIFT: <u>CRACIT33</u></strong>
                                    <br>As the cause of payment please indicate: "<strong>Order from web site n° <%=idordine%></strong>"<br><br>
                                    Beneficiary: <strong>Cristalensi Snc di Lensi Massimiliano & C. (P.Iva e C.Fiscale 05305820481)<br>
                             Via arti e mestieri, 1 - 50056 Montelupo F.no (Florence) Italy</strong>
                                    <br><br><br>
                                  The merchandise will be sent when our bank recieves payment, or to speed-up the consignment, you can send by e-mail the receipt of payment by bank transfer (in the case of home banking a receipt is usually provided by the bank, or you can scan the receipt left by the bank).<br>
                                  <br>
                                  By paying, and thereby completing the order, you have accepted the conditions of sale.  Save or stamp the conditions of sale (consultable on the appropriate page of the internet site) from this file (.pdf) using Adobe Acrobat Reader:<br>
                                  <a href="condizioni_di_vendita.pdf" target="_blank">sales conditions</a>
                                  <br>
                                  <br>
                                  Best regards from the staff of Cristalensi
                                  <br>
                                  <br>
                                </p>
                                <%end if%>
                                <%if FkPagamento=4 then%>
                                <p>
                                	<br><br>Thank you for having chosen our products,<br>
                                    to complete the order you must make out a POSTEPAY card transfer to the following address:<br>
                                    <br>
                                    <strong>Beneficiary: LENSI GIULIANO - c.f. LNS GLN 42A30 D403J<br>Card number: 4023 6005 5507 0285</strong>
                                    <br><br>As the cause of payment please indicate: "<strong>Order from web site n° <%=idordine%></strong>"<br><br>
                                    The merchandise will be sent when our bank recieves payment, or to speed-up the consignment, you can send by e-mail the receipt of payment by bank transfer.
                                  <br><br>
                                   By paying, and thereby completing the order, you have accepted the conditions of sale.  Save or stamp the conditions of sale (consultable on the appropriate page of the internet site) from this file (.pdf) using Adobe Acrobat Reader:<br><a href="condizioni_di_vendita.pdf" target="_blank">conditions of sale</a>
                                  <br>
                                  <br>
                                  Best regards from the staff of Cristalensi
                                  <br>
                                  <br>
                                </p>
                                <%end if%>
                      			<%if FkPagamento=3 then%>
                                <p>
                                	<br><br>Thank you for having chosen our products,<br>
                                    the merchandise will be sent to you as soon as we have it available in our warehouse.<br>
                                    We remind you that for payment on delivery, the courier will give you the merchandise only if paid in cash, no other method of payment (not even checks) will be accepted in this case<br>
                                    <br>
                                  By paying, and thereby completing the order, you have accepted the conditions of sale.  Save or stamp the conditions of sale (consultable on the appropriate page of the internet site) from this file (.pdf) using Adobe Acrobat Reader:<br>
                                  <a href="condizioni_di_vendita.pdf" target="_blank">sales conditions</a>
                                  <br>
                                  <br>
                                  Best regards from the staff of Cristalensi
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
                                  	
                                    <a href="https://www.paypal.com/it/webapps/mpp/paypal-popup" title="Come funziona PayPal" onClick="javascript:window.open('https://www.paypal.com/it/webapps/mpp/paypal-popup','WIPaypal','toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes, width=1060, height=700'); return false;"><img src="https://www.paypalobjects.com/webstatic/mktg/logo-center/logo_paypal_carte.jpg" border="0" style="float:right; padding-left:5px; width:319px; height:110px;" alt="Che cos'&egrave; PayPal"></a><br><br>Thank you for having chosen our products,<br>
							    To complete the order it is necessary to pay with the secure system Paypal which can be used with many types of credit card and prepaid cards:<br>
							    MasterCard, Visa and Visa Electron, PostePay, Carta Aura and PayPal top-ups.<br><br>
							    By paying, and thereby completing the order, you have accepted the conditions of sale.  Save or stamp the conditions of sale (consultable on the appropriate page of the internet site) from this file (.pdf) using Adobe Acrobat Reader:
							  <a href="condizioni_di_vendita.pdf" target="_blank">sales conditions</a>
                               <br>
                                    <br>
                                    </p>
    <form action="https://www.paypal.com/it/cgi-bin/webscr" method="post">
    <input type="hidden" name="cmd" value="_xclick">
    <input type="hidden" name="business" value="cristalensi@alice.it">
    <input type="hidden" name="item_name" value="Ordine n° <%=IdOrdine%>">
    <input type="hidden" name="item_number" value="<%=IdOrdine%>">
    <input type="hidden" name="currency_code" value="EUR">
    <input type="hidden" name="amount" value="<%=TotaleGeneralePP%>">
    <input type="hidden" name="return" value="http://www.cristalensi.it/en/pagamento_paypal_ok.asp">
    <input type="hidden" name="rm" value="2">
    <input type="hidden" name="cancel_return" value="http://www.cristalensi.it/en/pagamento_paypal_ko.asp">
    <input type="hidden" name="no_note" value="1">
    <input type="hidden" name="on0" value="Effettuato da">
    <input type="hidden" name="os0" value="<%=nome_log%>">
    <input type="hidden" name="cn" value="0">
    <input type="hidden" name="image_url" value="http://www.cristalensi.it/immagini/logo_cristalensi_piccolo.jpg">
    <input type="submit" name="paga_adesso" value="Click here to pay now with the sicure payment system Paypal" class="button" alt="Pay now with the sicure payment system Paypal" style="height: 40px; font-size:14px; font-weight:bold;">
    </form>								
                                  <p>
                                  <br><br>
                                  The merchandise will be sent as soon as our bank recieves payment.<br>
							  <br>
							  Best regards from the staff of Cristalensi
                                  <br>
                                  <br>
                                  </p>
                                <%end if%>
                                	
                                
                                <p class="area clearfix"><span class="colonna articolo">[article code] product name</span><span class="colonna quantita">quantity</span><span class="colonna prezzo_unitario">unit cost</span><span class="colonna prezzo_totale">total</span></p>
                                <div class="data">
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT PkId, FkOrdine, FkProdotto, PrezzoProdotto, Quantita, TotaleRiga, Titolo, CodiceArticolo, Colore FROM RigheOrdine WHERE FkOrdine="&idOrdine&""
	rs.Open sql, conn, 1, 1
	num_prodotti_carrello=rs.recordcount
	
%>                                    
									<%if rs.recordcount>0 then%>
                                        
                                        <%
                                        Do while not rs.EOF
                                        %>					
    
                                        <p class="riga"><span class="colonna articolo">[<%=rs("codicearticolo")%>]&nbsp;<%=rs("titolo")%><%if Len(rs("colore"))>0 then%>&nbsp;(<%=rs("colore")%>)<%end if%></span>
                                        <%
                                        quantita=rs("quantita")
                                        if quantita="" then quantita=1
                                        %>
                                        <span class="colonna quantita"><%=quantita%> pieces </span><span class="colonna prezzo_unitario"><%=FormatNumber(rs("PrezzoProdotto"),2)%>&#8364;</span><span class="colonna prezzo_totale"><%=FormatNumber(rs("TotaleRiga"),2)%>&#8364;</span></p>
                                        <%
                                        rs.movenext
                                        loop
                                        %>
<%
	rs.close
%>                                        
                                    <%end if%>
                                </div>
                                
                                <p class="area clearfix"><span class="colonna descrizione">Shipment method</span><span class="colonna prezzo_unitario">&nbsp;</span><span class="colonna prezzo_totale">Total</span></p>
                                <div class="data">
                                    <p class="riga">
                                    <span class="colonna descrizione"><b><%=TipoTrasporto%></b></span>
                                    <span class="colonna prezzo_unitario">&nbsp;</span>
                                    <span class="colonna prezzo_totale"><%=FormatNumber(CostoSpedizioneTotale,2)%>&#8364;</span>
                                    </p>
                                    <p>&nbsp;</p>
                                    <h4>Mailing address</h4>
                                    <p><%=DatiSpedizione%></p>
                                    <p>&nbsp;</p>
                                    <h4>Note</h4>
                                    <p><%=NoteCliente%></p>
                                </div>
                                
                                <p class="area clearfix"><span class="colonna descrizione">Method of Payment</span><span class="colonna prezzo_unitario">&nbsp;</span><span class="colonna prezzo_totale">Total</span></p>
                                <div class="data">
                                    <p class="riga">
                                    <span class="colonna descrizione"><b><%=TipoPagamento%></b></span>
                                    <span class="colonna prezzo_unitario">&nbsp;</span>
                                    <span class="colonna prezzo_totale"><%=FormatNumber(CostoPagamento,2)%>&#8364;</span>
                                    </p>
                                    <p>&nbsp;</p>
                                    <h4>Billing details:</h4>
                                    <div class="clearfix">
                                        <div class="table">
                                            <div class="tr">
                                                <div class="td">
                                                    <%if Rag_Soc<>"" then%><%=Rag_Soc%>&nbsp;&nbsp;<%end if%><%if nominativo<>"" then%><%=nominativo%><%end if%>
                                                </div>
                                            </div>
                                            <div class="tr">
                                                <div class="td">
                                                    Tax Code: <%=Cod_Fisc%><%if PartitaIVA<>"" then%> - VAT: <%=PartitaIVA%><%end if%>
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
                                  <span class="total_price">Total order:&nbsp;
								  <%if TotaleGenerale<>0 then%>
                                  <%=FormatNumber(TotaleGenerale,2)%>
                                  <%else%>
                                  0,00
                                  <%end if%>
                                  &#8364;
                                  </span>
                                  </h4>
                                    <form method="post" name="modulo" id="modulo" action="stampa_ordine.asp" target="_blank">
                                    <input type="hidden" name="idordine" id="idordine" value="<%=idordine%>">
                                    <%if FkPagamento=1 or FkPagamento=3 or FkPagamento=4 then%><button type="submit" name="stampa" style="float:right;" class="button_link_red">Print order</button><%end if%>
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