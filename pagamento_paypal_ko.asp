<!--#include file="inc_strConn.asp"-->
<%
	Call Visualizzazione("",0,"pagamento_paypal_ko.asp")
	
	IdOrdine=request("item_number")
	if IdOrdine="" then IdOrdine=0	
	
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Clienti where pkid="&idsession
		rs.Open sql, conn, 1, 1	
		
		nominativo_email=rs("nome")&" "&rs("nominativo")
		email=rs("email")
		
		rs.close
						
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Un ordine con pagamento da Paypal dal sito internet non &egrave; andato a buon fine.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti dell'ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Pagamento con Paypal non andato a buon fine"
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
			
			Mittente = "info@cristalensi.it"
			Destinatario = "viadeimedici@gmail.com"
			Oggetto = "Pagamento con Paypal non andato a buon fine"
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
        <title>Cristalensi - Pagamento PayPal</title>
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="js/media-queries-ie.js"></script>
        <![endif]-->
        <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
        <script src="js/jquery.blueberry.js"></script>
        <script src="js/jquery.tipTip.js"></script>
        <link href="css/css.css" rel="stylesheet" type="text/css">
        <link href="css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="css/tipTip.css" rel="stylesheet" type="text/css">
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
                            <div class="carrello clearfix">
                                <p>
                                <br /><br /><br />
                                La procedura di pagamento con Paypal &egrave; stata annullata<br>
							    oppure ci sono stati alcuni errori nel sistema di pagamento.<br>
							      <br>					    
							  Eventualmente contattare Cristalensi per avere dettagli e assistenza nel pagamento, grazie.<br><br>Telefono: 0571/911163<br>Email: <a href="mailto: info@cristalensi.it">info@cristalensi.it</a>
							  <br><br>
							  Il nostro personale &eacute; a tua disposizione per qualsiasi chiarimento.<br>
							  <br>
							  Cordiali saluti, lo staff di Cristalensi
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