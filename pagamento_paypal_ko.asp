<!--#include file="inc_strConn.asp"-->
<%
	Call Visualizzazione("",0,"pagamento_paypal_ko.asp")
	
	IdOrdine=request("item_number")
	if IdOrdine="" then IdOrdine=0	
	
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Clienti where pkid="&idsession
		rs.Open sql, conn, 1, 1	
		
		nominativo_email=rs("nominativo")
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

'			Set eMail_cdo = CreateObject("CDO.Message")
'
'			eMail_cdo.From = Mittente
'			eMail_cdo.To = Destinatario
'			eMail_cdo.Subject = Oggetto
'
'			eMail_cdo.HTMLBody = Testo
'
'			eMail_cdo.Send()
'
'			'invio al webmaster
'			
			Set eMail_cdo = Nothing
			
			Mittente = "info@cristalensi.it"
			Destinatario = "iurymazzoni@hotmail.com"
			Oggetto = "Pagamento con Paypal non andato a buon fine"
			Testo = HTML1

'			Set eMail_cdo = CreateObject("CDO.Message")
'
'			eMail_cdo.From = Mittente
'			eMail_cdo.To = Destinatario
'			eMail_cdo.Subject = Oggetto
'
'			eMail_cdo.HTMLBody = Testo
'
'			eMail_cdo.Send()
'
'			Set eMail_cdo = Nothing
			
			'fine invio email
%>

<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Cristalensi</title>
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
    </head>
    <body>
        <div id="wrap">
            <!--#include file="inc_header.asp"-->

            <div id="main-content">
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <div class="carrello clearfix">
                                <p class="testo_grande">
                                <br /><br /><br />
                                La procedura di pagamento con Paypal &egrave; stata annullata<br>
							    oppure ci sono stati alcuni errori nel sistema di pagamento.<br>
							      <br>					    
							  Eventualmente contattare Cristalensi per avere dettagli e assistenza nel pagamento, grazie.<br><br>Telefono: 0571/911163<br>Email: <a href="mailto: info@cristalensi.it">info@cristalensi.it</a>
							  <br><br>
							  Il nostro personale è a tua disposizione per qualsiasi chiarimento.<br>
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
          <script src="js/init.js"></script>
    </body>
</html>
<!--#include file="inc_strClose.asp"-->