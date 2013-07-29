<!--#include file="inc_strConn.asp"-->
<%
	mode=request("mode")
	if mode="" then mode=0
		
	if idsession=0 then response.Redirect("iscrizione.asp?prov=2")
		
	Destinazione=request("Destinazione")
	
	if mode=1 then
		testo=request("testo")
		if Len(testo)=0 then mode=2
		if Instr(1, testo, "www", 1)>0 then mode=2
		if Instr(1, testo, "@", 1)>0 then mode=2
	end if
	if mode=1 then
		Set cli_rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Commenti_Clienti"
		cli_rs.Open sql, conn, 3, 3
		cli_rs.addnew
			cli_rs("Testo")=request("Testo")
			cli_rs("FkIscritto")=idsession
			cli_rs("Data")=now()
			cli_rs("Pubblicato")=False
			cli_rs("Risposta")=False
		cli_rs.update
		cli_rs.close
		
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Thank you "&nominativo_email&" to send a comment!<br>If accepted by our moderators you will receive an email notification of the publication.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best wishes from the staff of Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Confirmation for a comment on Cristalensi.it"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo commento sul sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo commento:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Conferma invio commento a Cristalensi.it (sito inglese)"
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
			
			'invio al webmaster
			
			Mittente = "info@cristalensi.it"
			Destinatario = "iurymazzoni@hotmail.com"
			Oggetto = "Conferma invio commento a Cristalensi.it (sito inglese)"
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
	end if
	
%>
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Comments Cristalensi</title>
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="../js/media-queries-ie.js"></script>
        <![endif]-->
        <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
        <script src="../js/jquery.blueberry.js"></script>
        <script src="../js/jquery.tipTip.js"></script>
        <link href="../css/css.css" rel="stylesheet" type="text/css">
        <link href="../css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="../css/tipTip.css" rel="stylesheet" type="text/css">
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
            <link href="../css/tipTip_ie7.css" media="all" rel="stylesheet" type="text/css" />
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
        
          //var _gaq = _gaq || [];
//          _gaq.push(['_setAccount', 'UA-320952-2']);
//          _gaq.push(['_trackPageview']);
//        
//          (function() {
//            var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
//            ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
//            var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
//          })();
        
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
                            <h3 style="font-size: 14px; display: inline; border: none;">Send a comment you too!</h3>
                            <div class="carrello clearfix">
                                 <%if mode=1 then%>
                                 	<p>
                                    Your comment has been inserted correctly, now our staff will evaluate and if approved, we will send you an email notification. <br /> Thank you for your cooperation with the staff of Cristalensi.<br /><br /><a href="commenti_elenco.asp" class="button_link_red" style="float:right">ALL COMMENTS AND REVIEWS</a>
                                    </p>
								 <%else%>   
                                    <form name="modulocarrello" id="modulocarrello" method="post" action="commenti_form.asp?mode=1">
                                    <p>Send a comment on the products purchased, whether you liked it or not, or a review on the website or the company and the staff. <br /> Comment will not be published immediately but will be subject to inspection by our staff to prevent them from being inserted content to be unlawful, offensive and terms not be published. <br /> Please do not insert html code, email and links to other websites: the comment will not be published. <br /> In every comments will also be published <strong> Name </ strong> and <strong> City </ strong> submitted at registration.</p>
                                    <%if mode=2 then%>
                                    <p><br><br><strong>Warning! Check the text entered by the rules, thank you.</strong><br><br></p>
                                    <%end if%>
                                    <p>
                                    <textarea name="testo" cols="105" rows="7" id="testo"></textarea>
                                    <br><br>
                                    <button type="button" name="reset" class="button_link" style="float:left;" onClick="location.href='commenti_elenco.asp'">&laquo; ALL COMMENTS AND REVIEWS</button>
                                    <button type="submit" name="continua" class="button_link_red" style="float:right">Click here to submit your comment &raquo;</button>
                                    </p>
                                    </form>
								<%end if%>
                            </div>
                        </div>
                    </div>
                </div>
                <!--#include file="inc_sx.asp"-->
            </div>
        </div>
         <!--#include file="inc_footer.asp"-->
    </body>
</html>
<!--#include file="inc_strClose.asp"-->