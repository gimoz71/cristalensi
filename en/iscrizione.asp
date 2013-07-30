<!--#include file="inc_strConn.asp"-->
<%
prov=request("prov")
if prov="" then prov=0
'se = 0 proviene dal sito
'se = 1 proviene dal negozio

	pkid = Session("idCliente")
	if pkid = "" then pkid = 0

	mode = request("mode")
	if mode = "" then mode = 0

	if mode=1 then
		nome=request("nome")
		nominativo=request("nominativo")
		partitaIVA=request("partitaIVA")
		cod_fisc=request("cod_fisc")
		rag_soc=request("rag_soc")
		indirizzo=request("indirizzo")
		cap=request("cap")
		citta=request("citta")
		provincia=request("provincia")
		italia=request("italia")
		nazionediversa=request("nazionediversa")
		telefono=request("telefono")
		fax=request("fax")
		email=request("email")
		aut_email=request("aut_email")
		password=request("password")
		data=now()
		ip=Request.ServerVariables("REMOTE_ADDR")
	end if

	if mode=1 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select email From Clienti where email='"&email&"'"
		if pkid>0 then sql = "Select email,pkid From Clienti where email='"&email&"' and pkid<>"&pkid&""
		rs.Open sql, conn, 1, 1
		if rs.recordcount>0 then mode=3
		rs.close
	end if
	
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Clienti"
	if pkid > 0 then sql = "Select * From Clienti where pkid="&pkid
	rs.Open sql, conn, 3, 3
	
	
	if mode = 1 then
		if pkid = 0 then rs.addnew
		
		rs("nome")=nome
		rs("nominativo")=nominativo
		rs("partitaIVA")=partitaIVA
		rs("cod_fisc")=cod_fisc
		rs("rag_soc")=rag_soc
		rs("indirizzo")=indirizzo
		rs("cap")=cap
		rs("citta")=citta
		rs("provincia")=provincia
		rs("italia")=italia
		rs("nazionediversa")=nazionediversa
		rs("telefono")=telefono
		rs("fax")=fax
		rs("email")=email
		rs("aut_email")=aut_email
		rs("password")=password
		rs("data")=data
		rs("ip")=ip
		rs("aut_privacy")=True
		
		rs.update
		rs.close
		
		if pkid=0 then
			Set rs=Server.CreateObject("ADODB.Recordset")
			sql = "Select @@Identity As pkid"
			rs.Open sql, conn, 1, 1
				pkid_iscritto=rs("pkid")
			rs.close
		
			
			'invio l'email di benvenuto al cliente
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Congratulations "&nome&" "&nominativo&"! Your subscription to Cristalensi.it was successful. <br> From now you can order our products without having to re-enter your information.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Sensitive data and determining access to services Cristalensi.it:<br>Name and Surname: <b>"&nome&" "&nominativo&"</b><br>Login: <b>"&email&"</b><br>Password: <b>"&password&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Iscrizione al sito Cristalensi.it"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuova registrazione al sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti per l'accesso ai servizi di Cristalensi.it:<br>Nome e Cognome: <b>"&nome&" "&nominativo&"</b><br>Login: <b>"&email&"</b><br>Password: <b>"&password&"</b><br>Codice cliente: <b>"&pkid_iscritto&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Nuova iscrizione al sito Cristalensi.it (sito inglese)"
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
			
			Set eMail_cdo = Nothing
			
			Mittente = "info@cristalensi.it"
			Destinatario = "iurymazzoni@hotmail.com"
			Oggetto = "Nuova iscrizione al sito Cristalensi.it (sito inglese)"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set eMail_cdo = Nothing
			
		

		nome_log=nome&" "&nominativo
		session("nome_log")=nome&" "&nominativo
		idsession=pkid_iscritto
		session("idCliente")=pkid_iscritto
		italia_log=italia
		if italia_log="" then italia_log="Si"
		if italia_log="Sì" then italia_log="Si"
		if italia_log="S&igrave;" then italia_log="Si"
		session("italia_log")=italia_log
		
		end if
		
	end if
	
	if mode=2 and pkid=0 then response.Redirect("iscrizione.asp")
%>
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>CRISTALENSI Client Authentication Lamps store online</title>
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
				
			nome=document.newsform.nome.value;
			nominativo=document.newsform.nominativo.value;
			indirizzo=document.newsform.indirizzo.value;
			citta=document.newsform.citta.value;
			telefono=document.newsform.telefono.value;
			email=document.newsform.email.value;
			password=document.newsform.pw.value;	
		
			if (nome==""){
				alert("It has not been filled in the field \"Name\".");
				return false;
			}
			if (nominativo==""){
				alert("It has not been filled in the field \"Surname\".");
				return false;
			}
			if (indirizzo==""){
				alert("It has not been filled in the field \"Address\".");
				return false;
			}
			if (citta==""){
				alert("It has not been filled in the field \"City\".");
				return false;
			}
			if (telefono==""){
				alert("It has not been filled in the field \"Telephone\".");
				return false;
			}
			if (email==""){
				alert("It has not been filled in the field \"Email\".");
				return false;
			}
			if (email.indexOf("@")==-1 || email.indexOf(".")==-1){
			alert("ATTENTION! \"e-mail\" is not correct.");
			return false; 
			}
			if (password==""){
				alert("It has not been filled in the field \"Password\".");
				return false;
			}
		
			else
		return true
		
		}
		
		function accetta(el){
		checkobj=el
			if (document.all||document.getElementById){
				for (i=0;i<checkobj.form.length;i++){
		var tempobj=checkobj.form.elements[i]
			if(tempobj.type.toLowerCase()=="submit")
		tempobj.disabled=!checkobj.checked
									}
								}
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
                        <%if mode=0 or mode=3 then%>
					  
					  	<%if pkid>0 then%>
                        	<h3 style="font-size: 14px; display: inline; border: none;">Changing client data</h3>
                            <p>On this page you can edit the data that you used to sign up Cristalensi.<br>
						Important intormation:  it is necessary that the <strong>Email address</strong> be in function and that it is one that you use frequently given that you will be sent information regarding the state of your order.<br>
We remind you moreover that the <strong>Email address</strong> will be your <strong>Login</strong> to place future orders.</p>
                        
                        <%else%>
                        	
                        	<h3 style="font-size: 14px; display: inline; border: none;">Client Authentication</h3>
                            <p>If you have already signed in, and therefore already have a <strong>Login</strong> and <strong>Password</strong>, it is not necessary to sign on again,
it is sufficient that you enter the access data here: <strong>Login (email)</strong> and <strong>Password</strong></p>
							<div class="iscrizione clearfix" style="background-color: #FEF4C0; padding:5px 0px 8px 8px; border:0px; margin-bottom: 20px;">
                            	<form method="post" action="iscrizione.asp?mode=2&contr=1" name="newsform2">
                                    <div class="half_panel left_p" style="">
                                        <div >
                                            Login <input name="login" type="text" id="login" size="30" />
                                        </div>
                                        <div class="">
                                            Password <input name="password" type="password" id="password" size="30" />
                                        </div>
                                    </div>
                                    <div class="half_panel left_p">
                                        <div style="margin-top: 25px;">
                                            <span style="text-align: center;">&nbsp;&nbsp;&nbsp;<button type="submit" name="accedi" class="button_link_red">sign in</button>&nbsp;&nbsp;&nbsp;</span>
                                            <a href="recupero_pw.asp" style="font-size:11px;">Click here to recover your password</a>
                                        </div>
                                    </div>
                                </form>
                            </div>
                            <hr>
                            <h3 style="font-size: 14px; display: inline; border: none;">Client registration</h3>
                            <p>On this page please insert your data to register with Cristalensi.<br>
						Important intormation:  it is necessary that the <strong>Email address</strong> be in function and that it is one that you use frequently given that you will be sent information regarding the state of your order.<br>
We remind you moreover that the <strong>Email address</strong> will be your <strong>Login</strong> to place future orders.
                            </p>
                            
                          <%end if%>
                            <div class="iscrizione clearfix">
                                
                                <div class="table">
                                    <form method="post" action="iscrizione.asp?mode=1&amp;pkid=<%=pkid%>" name="newsform" onSubmit="return verifica();">
                                    <div class="tr">
                                        <div class="td">
	                                        Name (*) and Surname (*)<br />
                                            <input name="nome" type="text" id="nome" style="width:120px;" maxlength="50" value="<% if pkid > 0 then %><%=rs("nome")%><%else%><%if mode=3 then%><%=nome%><%end if%><%end if%>" />&nbsp;<input name="nominativo" type="text" id="nominativo" style="width:120px;" maxlength="50" value="<% if pkid > 0 then %><%=rs("nominativo")%><%else%><%if mode=3 then%><%=nominativo%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
                                        	Company name (in the case that the order is being placed by a business)<br />
                                            <input name="Rag_Soc" type="text" id="Rag_Soc" maxlength="50" value="<% if pkid > 0 then %><%=rs("Rag_Soc")%><%else%><%if mode=3 then%><%=Rag_Soc%><%end if%><%end if%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        Tax Code<br />
                                            <input name="cod_fisc" type="text" id="cod_fisc"  size="20" maxlength="16" value="<% if pkid > 0 then %><%=rs("cod_fisc")%><%else%><%if mode=3 then%><%=cod_fisc%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
                                        Value Added Tax registration number or equivalent (in the case of a business)<br />
                                            <input name="PartitaIVA" type="text" id="PartitaIVA"  size="20" maxlength="11" value="<% if pkid > 0 then %><%=rs("PartitaIVA")%><%else%><%if mode=3 then%><%=PartitaIVA%><%end if%><%end if%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        	Address (*)<br />
                                            <input name="indirizzo" type="text" id="indirizzo"  size="30" maxlength="100" value="<% if pkid > 0 then %><%=rs("indirizzo")%><%else%><%if mode=3 then%><%=indirizzo%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
                                        	Post Code<br />
                                            <input name="cap" type="text" id="cap"  size="7" maxlength="5" value="<% if pkid > 0 then %><%=rs("cap")%><%else%><%if mode=3 then%><%=cap%><%end if%><%end if%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
	                                        City (*)<br />
                                            <input name="citta" type="text" id="citta"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("citta")%><%else%><%if mode=3 then%><%=citta%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
	                                        Province/Region<br />
                                            <input type="text" name="provincia" id="provincia" value="<% if pkid > 0 then %><%=rs("provincia")%><%else%><%if mode=3 then%><%=provincia%><%end if%><%end if%>" size="3" maxlength="2" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>Nation</div>
                                    </div>
                                    <div class="tr">
                                        Italy:&nbsp;&nbsp;Si&nbsp;<input type="radio" name="italia" value="Si" <% if pkid > 0 then %><%if rs("italia")="Si" then%> checked<%end if %><%else%> checked<%end if %> />&nbsp;&nbsp;No&nbsp;<input type="radio" name="italia" value="No" <% if pkid > 0 then %><%if rs("italia")="No" then%> checked<%end if %><%end if %> />&nbsp;Other nation
                <input name="nazionediversa" type="text" id="nazionediversa"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("nazionediversa")%><%else%><%if mode=3 then%><%=nazionediversa%><%end if%><%end if%>" />
                                        
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        Phone (*)<br />
                                            <input name="telefono" type="text" id="telefono"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("telefono")%><%else%><%if mode=3 then%><%=telefono%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
	                                        Fax<br />
                                            <input name="fax" type="text" id="fax"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("fax")%><%else%><%if mode=3 then%><%=fax%><%end if%><%end if%>" />
                                        </div>
                                    </div>
                                    <%if mode=3 then%>
                                    <div class="tr">
                                        <div class="td" style="background-color:#F90;">
                                        <font color="#990000"><b>Warning! The e-mail is not acceptable </b></font>
                                        </div>
                                        <div class="td" style="background-color:#F90;">
                                        &nbsp;
                                        </div>
                                    </div>
                                    <%end if%>
                                    <div class="tr">
                                        <div class="td">
                                        	<strong>E-mail</strong> (*) - Will be used as the <strong>Login</strong> for future orders<br />
                                            <input name="email" type="text" id="email"  size="30" maxlength="100" value="<% if pkid > 0 then %><%=rs("email")%><%else%><%if mode=3 then%><%=email%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
                                        <strong>Password</strong> (*)<br />
                                            <input name="password" type="password" id="pw"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("password")%><%else%><%if mode=3 then%><%=password%><%end if%><%end if%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>Authorization to receive an Email </div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <input type="radio" name="aut_email" value=True <% if pkid > 0 then %><%if rs("aut_email")=True then%> checked<%end if %><%else%> checked<%end if %> />
                                            &nbsp;Yes&nbsp;&nbsp;
                                                    <input type="radio" name="aut_email" value=False <% if pkid > 0 then %><%if rs("aut_email")=False then%> checked<%end if %><%end if %> />
                                            &nbsp;No
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>&nbsp;</div>
                                    </div>
                                    <div class="tr">
                                        <div>Conditions regading the use of personal data</div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <textarea name="privacy" style="width: 700px;" rows="5" readonly>INFORMATION RELATIVE TO THE TREATMENT OF PERSONAL DATA.  According to the sense of article 10 of the Law number 675 of the 31/12/1996, the Company informs the interested party that the data regarding it, furnished by the same, will be subject to treatment with  respect to the above mentioned norm.  These data will be used to scopes of a gestional, commercial, and promotional nature. The releasing of data to our Company is entirely optional. Data acquired may be communicated and diffused in observation of the dispositions contained in article 20 of Law 675/96 in persuance of the finalities above mentioned.  The owner of the treatment is Cristalensi s.n.c. whose seat is in via arti e mestieri 1 Montelupo F.no (Fi) where, moreover, the responsable  pro tempore of the treatment resides, whose identity can be obtained from the Public Register held by the Garante, or from the legal offices of the Company.  In addition, the Company informs interested parties that they may exercise the rights  foreseen in article 13 of the law 675/96, that is:  know  without charge, through the General Register of the Garante, the treatments of data which concern the interested party;  there can be obtained from Cristalensi s.n.c., - with a contribution towards the costs only in the case of a negative response- the confirmation or negation of the existence, in the company archives, of data which regard the interested party, and can have access to information regarding the finalities to which the data have been put. The request is renewable after ninty days;  Obtain the cancellation, the transformation into anonymous form and the blocking of data treated in violation of the sense of the law; Obtain the updating, the correction or the integration of the data;  Object, without charge, to the treatment of data which concerns the interested party. </textarea>
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <input name="chekka" type="checkbox" onClick="accetta(this)" />
                                            Accept the conditions </div>
                                    </div>
                                    <div class="tr">
                                        <div>&nbsp;</div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <button name="Submit" type="submit" class="button_link_red" align="absmiddle" disabled>Submit</button>
                                            &nbsp;
                                            <button name="Submit2" type="reset" class="button_link">Reset</button>
                                            (*) this space must be filled </div>
                                    </div>
                                    <div class="tr">
                                        <div>&nbsp;</div>
                                    </div>
                                    </form>
                                </div>
                                </div>
                                <%else%>
                                 	<h3 style="font-size: 14px; display: inline; border: none;">Client Authentication</h3>
                                    <div class="iscrizione clearfix">
									  <%if mode=2 then%>
                                          <br /><br />
                                          <font color="#990000"><b>Welcome&nbsp;<%=nome_log%></b></font>
                                          <br /><br /><br /><br />
                                          <a href="<%if italia_log="Si" then%>carrello2.asp<%end if%><%if italia_log="No" then%>carrello2extra.asp<%end if%>">clicking here will return directly to the order to complete:<br />your cart on-line &raquo;</a>
                                          <br /><br /><br /><br />
                          					<a href="commenti_form.asp">clicking here will post a comment on web site:<br />
						  					post a comment &raquo;</a>
                                      <%else%>
										  <%if pkid=0 then%>
                                              <br /><br />Thank yuo for signing up to Cristalensi.it<br />
						  					  From this moment you can complete and send your orders or post comments without having to enter all your information.
                                              
                                              <br /><br /><br /><br />
                                              <a href="<%if italia_log="Si" then%>carrello2.asp<%end if%><%if italia_log="No" then%>carrello2extra.asp<%end if%>">clicking here will return directly to the order to complete:<br />your cart on-line &raquo;</a>
                                              <br /><br /><br /><br />
                                                <a href="commenti_form.asp">clicking here will post a comment on web site:<br />
						  					post a comment &raquo;</a>
                                          <%else%>
                                              <br /><br />
                                              Your data has been updated regularly.
                                          <%end if%>
                              		  <%end if%>
                              		</div>
                            	
                            <%end if%>    
                        	
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