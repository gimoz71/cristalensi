<!--#include file="inc_strConn.asp"-->
<%

	mode = request("mode")
	if mode = "" then mode = 0

	if mode=1 then
		email=request("email")
	end if

	if mode=1 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select email,password,nominativo,nome From Clienti where email='"&email&"'"
		rs.Open sql, conn, 1, 1
		if rs.recordcount=0 then
			mode=2
		else
			nominativo=rs("nominativo")
			nome=rs("nome")
			password=rs("password")
		end if
		rs.close
	end if
	
	if mode = 1 then
		
			
			'invio l'email di recupero pw al cliente
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dear "&nome&" "&nominativo&", the password inserted at your registration with Cristalensi.it is the following:<br><br></font>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Password: <b>"&password&"</b><br>Login: <b>"&email&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Retrieve your Cristalensi.it password"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>E' stata fatta una richiesta di recupero password dal seguente cliente: "&nome&" "&nominativo&"<br> La password inserita al momento dell'iscrizione a Cristalensi.it &egrave; la seguente:<br></font>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Password: <b>"&password&"</b><br>Login: <b>"&email&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Richiesta recupero password dal sito Cristalensi.it (sito inglese)"
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
        <title>Cristalensi Retrieval password client</title>
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
        <SCRIPT language="JavaScript">

		function verifica() {
				
			email=document.newsform.email.value;
		
			if (email==""){
				alert("It has not been filled in the field \"Email\".");
				return false;
			}
			if (email.indexOf("@")==-1 || email.indexOf(".")==-1){
			alert("ATTENZIONE! \"e-mail\" non valida.");
			return false; 
			}
		
			else
		return true
		
		}
		
		</SCRIPT>
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
                        <%if mode=0 or mode=2 then%>                        
                        	<h3 style="font-size: 14px; display: inline; border: none;">Retrieval password client</h3>
                            <p>From this page you can obtain the <strong>password</strong> used at the moment of <strong>your registration</strong> with Cristalensi.<br>
						<strong>Important information</strong>:  it is necessary that the <strong>e-mail</strong> you use be the same as that used at registration.<br>
						We remind you moreover that the <strong>e-mail</strong> address will need to be used as a <strong>Login</strong> for future orders.</p>
							<div class="iscrizione clearfix">
                            	<form method="post" action="recupero_pw.asp?mode=1" name="newsform" onSubmit="return verifica();">
                                <div class="table">
                                    <div class="tr" style="text-align:center;">
                                            <strong>E-mail</strong> (compulsory)
                                            
                                    </div>
                                    <div class="tr" style="text-align:center;">
                                    <input name="email" type="text" id="email" size="30" maxlength="30" value="<% if pkid > 0 then %><%=rs("email")%><%else%><%if mode=2 or mode=3 then%><%=email%><%end if%><%end if%>" />
                                    </div>
                                    <%if mode=2 then%>
                                    <div class="tr" style="text-align:center;">
                          			<font color="#990000"><b>Attention! The e-mail inserted is not correct</b></font>
                                    </div>
                        			<%end if%>
                                    <div class="tr text-center"><div class="td"><button type="submit" name="accedi" class="button_link">submit</button></div></div>
                                </div>
                                </form>
                            </div>
                        <%else%>
                        	<h3 style="font-size: 14px; display: inline; border: none;">Retrieval password client</h3>
                            <p style="text-align:center; padding-top:20px">The  access password to Cristalensi.it has been successfully sent your e-mail address:  checking it you will be able to retrieve the access data for the internet site <br /><br /><a href="prodotti.asp">To return to the product gallery, click here.</a>
                            </p>    
                        <%end if%>
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