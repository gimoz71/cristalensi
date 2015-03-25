<!--#include file="inc_strConn.asp"-->
<%
	Call Visualizzazione("",0,"calcolospedizione.asp")
	
	IdOrdine=session("ordine_shop")
	session("ordine_shop")=""
	if IdOrdine="" then IdOrdine=0
	if idOrdine=0 then response.redirect("carrello1.asp")
	
	if idsession=0 then response.Redirect("iscrizione.asp?prov=1")
	
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Thank you "&nominativo_email&" for having chosen our products!<br>This email is confirmation that we have received your request for the mailing costs necessary to continue with your order n&deg; "&idordine&".<br> In the next few hours (max 24) you will recieve our reply indicating that the order can be completed. At this point it will be necessary to insert your Login and Password in the Client area of the Home Page, click on the link ""my orders"" and then click on the order in question. To complete the order you will need to  choose the method of payment (bank transfer or PayPal with credit card or prepaid card).</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best regards, from the staff of Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Confirmation of request for mailing costs of order no. "&idordine&", Cristalensi.it"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>E' stato richiesto il calcolo dei costi di spedizione per un nuovo ordine.</font><br><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Richiesta calcolo costi di spedizione ordine n "&idordine&" (sito inglese)"
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
			Oggetto = "Richiesta calcolo costi di spedizione ordine n "&idordine&" (sito inglese)"
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
%>
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
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
		<script type="text/javascript">
        
          var _gaq = _gaq || [];
          _gaq.push(['_setAccount', 'UA-320952-2']);
          _gaq.push(['_trackPageview']);
        
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
                            <h3 style="font-size: 14px; display: inline; border: none;">Your order:&nbsp; shipment  method/receipt of product</h3>
                            <div class="carrello clearfix">
                                <p>
                                <br /><br /><br />
                                Thank you for having chosen our products,<br />
							    a request has been made for the calculation of the mailing cost<br />
							    in order to continue order n&deg;: <%=idordine%><br />
							    <br>
							    In the next few hours (max 24h) you will receive an email which will  allow you to complete the order.&nbsp; To  continue you will need to insert the Login and Password in the client area of  the Home page. To complete the order you must choose a method of payment (bank  transfer or payment through PayPal via credit card or prepaid credit card).<br /><br /><br />
							    Best wishes from the staff of Cristalensi
                                </p>
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