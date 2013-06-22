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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Complimenti "&nominativo&"! La tua iscrizione a Cristalensi.it &egrave; avvenuta correttamente.<br>Da adesso potrai ordinare i nostri prodotti senza dover inserire nuovamente i tuoi dati.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti per l'accesso ai servizi di Cristalensi.it:<br>Nominativo: <b>"&nominativo&"</b><br>Login: <b>"&email&"</b><br>Password: <b>"&password&"</b></font><br>"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti per l'accesso ai servizi di Cristalensi.it:<br>Nominativo: <b>"&nominativo&"</b><br>Login: <b>"&email&"</b><br>Password: <b>"&password&"</b><br>Codice cliente: <b>"&pkid_iscritto&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Nuova iscrizione al sito Cristalensi.it"
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
			Oggetto = "Nuova iscrizione al sito Cristalensi.it"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set eMail_cdo = Nothing
			
		end if

		nome_log=nominativo
		session("nome_log")=nominativo
		idsession=pkid_iscritto
		session("idCliente")=pkid_iscritto
		italia_log=italia
		session("italia_log")=italia_log
		
	end if
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
        <SCRIPT language="JavaScript">

		function verifica() {
				
			nominativo=document.newsform.nominativo.value;
			indirizzo=document.newsform.indirizzo.value;
			citta=document.newsform.citta.value;
			telefono=document.newsform.telefono.value;
			email=document.newsform.email.value;
			password=document.newsform.pw.value;	
		
			if (nominativo==""){
				alert("Non  e\' stato compilato il campo \"Nominativo\".");
				return false;
			}
			if (indirizzo==""){
				alert("Non  e\' stato compilato il campo \"Indirizzo\".");
				return false;
			}
			if (citta==""){
				alert("Non  e\' stato compilato il campo \"Citta\".");
				return false;
			}
			if (telefono==""){
				alert("Non  e\' stato compilato il campo \"Telefono\".");
				return false;
			}
			if (email==""){
				alert("Non  e\' stato compilato il campo \"Email\".");
				return false;
			}
			if (email.indexOf("@")==-1 || email.indexOf(".")==-1){
			alert("ATTENZIONE! \"e-mail\" non valida.");
			return false; 
			}
			if (password==""){
				alert("Non  e\' stato compilato il campo \"Password\".");
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
                        	<h3 style="font-size: 14px; display: inline; border: none;">Modifica dati cliente</h3>
                            <p>In questa pagina puoi modificare i dati che hai usato per registrarti a Cristalensi.<br />
						Informazione importante: &egrave; necessario che l'indirizzo <strong>Email</strong> sia un'indirizzo funzionante e che usi normalmente, in quanto ti verranno spedite comunicazioni relativamente agli ordini e ai prodotti.<br />
						Ti ricordiamo inoltre che l'indirizzo <strong>Email</strong> lo dovrai utilizzare come <strong>Login</strong> per accedere ai tuoi futuri ordini.</p>
                        
                        <%else%>
                        	
                        	<h3 style="font-size: 14px; display: inline; border: none;">Autenticazione cliente</h3>
                            <p>Se sei gi&agrave; iscritto, e quindi hai gi&agrave; Login (Email) e Password, non &egrave; necessario che ti iscriva nuovamente, &egrave; sufficiente inserire i dati di accesso qui sotto e sarai riconosciuto immediatamente.</p>
							<div class="iscrizione clearfix">
                            	<form method="post" action="iscrizione.asp?mode=2&contr=1" name="newsform2">
                                <div class="table">
                                    <div class="tr">
                                        <div class="td">
                                            Login <input name="login" type="text" class="form" id="login" size="30" />
                                        </div>
                                        <div class="td">
                                            Password <input name="password" type="password" class="form" id="password" size="30" />
                                            
                                        </div>
                                       
                                    </div>
                                    <div class="tr"><p style="text-align: center;"><input type="submit" value="accedi" />&nbsp;&nbsp;&nbsp;<a href="recupero_pw.asp">Clicca qui per recuperare la password</a></p></div>
                                </div>
                                </form>
                            </div>
                            <hr>
                            <h3 style="font-size: 14px; display: inline; border: none;">Registrazione cliente</h3>
                            <p>In questa pagina puoi inserire i tuoi dati per registrarti a Cristalensi.<br />
                                Informazione importante: &egrave; necessario che l'indirizzo Email sia un'indirizzo funzionante e che usi normalmente, in quanto ti verranno spedite comunicazioni relativamente agli ordini e ai prodotti.<br />Ti ricordiamo inoltre che l'indirizzo Email lo dovrai utilizzare come Login per accedere ai tuoi futuri ordini.
                            </p>
                            
                          <%end if%>
                            <div class="iscrizione clearfix">
                                
                                <div class="table">
                                    <form method="post" action="iscrizione.asp?mode=1&amp;pkid=0" name="newsform" onSubmit="return verifica();">
                                    <div class="tr">
                                        <div class="td">
	                                        Nome e Cognome (*)<br />
                                            <input name="nominativo" type="text" class="form" id="nominativo"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("nominativo")%><%else%><%if mode=3 then%><%=nominativo%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
                                        	Ragione sociale ( nel caso in cui si tratti di un'Azienda )<br />
                                            <input name="Rag_Soc" type="text" class="form" id="Rag_Soc"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("Rag_Soc")%><%else%><%if mode=3 then%><%=Rag_Soc%><%end if%><%end if%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        Codice Fiscale<br />
                                            <input name="cod_fisc" type="text" class="form" id="cod_fisc"  size="20" maxlength="16" value="<% if pkid > 0 then %><%=rs("cod_fisc")%><%else%><%if mode=3 then%><%=cod_fisc%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
                                        Partita IVA ( nel caso in cui si tratti di un'Azienda )<br />
                                            <input name="PartitaIVA" type="text" class="form" id="PartitaIVA"  size="20" maxlength="11" value="<% if pkid > 0 then %><%=rs("PartitaIVA")%><%else%><%if mode=3 then%><%=PartitaIVA%><%end if%><%end if%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        	Indirizzo (*)<br />
                                            <input name="indirizzo" type="text" class="form" id="indirizzo"  size="30" maxlength="100" value="<% if pkid > 0 then %><%=rs("indirizzo")%><%else%><%if mode=3 then%><%=indirizzo%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
                                        	CAP<br />
                                            <input name="cap" type="text" class="form" id="cap"  size="7" maxlength="5" value="<% if pkid > 0 then %><%=rs("cap")%><%else%><%if mode=3 then%><%=cap%><%end if%><%end if%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
	                                        Citt&agrave; (*)<br />
                                            <input name="citta" type="text" class="form" id="citta"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("citta")%><%else%><%if mode=3 then%><%=citta%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
	                                        Provincia<br />
                                            <input type="text" name="provincia" id="provincia" value="<% if pkid > 0 then %><%=rs("provincia")%><%else%><%if mode=3 then%><%=provincia%><%end if%><%end if%>" size="3" maxlength="2" class="form" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>Nazione</div>
                                    </div>
                                    <div class="tr">
                                        Italia:&nbsp;&nbsp;Si&nbsp;<input type="radio" name="italia" value="S&igrave;" <% if pkid > 0 then %><%if rs("italia")="S&igrave;" then%> checked<%end if %><%else%> checked<%end if %> />&nbsp;&nbsp;No&nbsp;<input type="radio" name="italia" value="No" <% if pkid > 0 then %><%if rs("italia")="No" then%> checked<%end if %><%end if %> />&nbsp;Altra nazione
                <input name="nazionediversa" type="text" class="form" id="nazionediversa"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("nazionediversa")%><%else%><%if mode=3 then%><%=nazionediversa%><%end if%><%end if%>" />
                                        
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        Telefono (*)<br />
                                            <input name="telefono" type="text" class="form" id="telefono"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("telefono")%><%else%><%if mode=3 then%><%=telefono%><%end if%><%end if%>" />
                                        </div>
                                        <div class="td">
	                                        Fax<br />
                                            <input name="fax" type="text" class="form" id="fax"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("fax")%><%else%><%if mode=3 then%><%=fax%><%end if%><%end if%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        	<strong>E-mail</strong> (*) - Verr&agrave; usata come <strong>Login</strong> per i futuri ordini<br />
                                            <input name="email" type="text" class="form" id="email"  size="30" maxlength="100" value="<% if pkid > 0 then %><%=rs("email")%><%else%><%if mode=3 then%><%=email%><%end if%><%end if%>" />
									  	<%if mode=3 then%>
                                          &nbsp;&nbsp;<font color="#990000"><b>Attenzione! L'e-mail inserita non pu&ograve; essere accettata </b></font>
                                        <%end if%>
                                        </div>
                                        <div class="td">
                                        <strong>Password</strong> (*)<br />
                                            <input name="password" type="password" class="form" id="pw"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("password")%><%else%><%if mode=3 then%><%=password%><%end if%><%end if%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>Autorizzazione a ricevere Email </div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <input type="radio" name="aut_email" value=True <% if pkid > 0 then %><%if rs("aut_email")=True then%> checked<%end if %><%else%> checked<%end if %> />
                                            &nbsp;Si&nbsp;&nbsp;
                                                    <input type="radio" name="aut_email" value=False <% if pkid > 0 then %><%if rs("aut_email")=False then%> checked<%end if %><%end if %> />
                                            &nbsp;No
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>&nbsp;</div>
                                    </div>
                                    <div class="tr">
                                        <div>Condizioni sul trattamento dei dati personali</div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <textarea name="privacy" cols="80" rows="5" readonly class="form">INFORMAZIONI RELATIVE AL TRATTAMENTO DI DATI PERSONALI 
Ai sensi dell'art. 10 della L. n&deg;675 del 31/12/1996, l'Azienda informa l'interessato che i dati che lo riguardano, forniti dall'interessato medesimo, formeranno oggetto di trattamento nel rispetto della normativa sopra richiamata. Tali dati verranno trattati per finalita' gestionali, commerciali, promozionali. Il conferimento dei dati alla nostra Azienda e' assolutamente facoltativo. 
I dati acquisiti potranno essere comunicati e diffusi in osservanza di quanto disposto all'articolo 20 della legge 675/96 allo scopo di perseguire le finalita' sopra indicate. 

Il titolare del trattamento e' Cristalensi s.n.c. 
 con sede in via arti e mestieri, 1  
Montelupo F.no (FI)
, ove e' altres&igrave; domiciliato il responsabile protempore del trattamento, i cui dati identificativi possono essere acquisiti presso il Registro pubblico tenuto dal Garante, o presso la sede legale dell'Azienda. 

L'Azienda informa altres&igrave; l'Interessato che questi potra' esercitare i diritti previsti dall'articolo 13 della legge 675/96, ossia:
Conoscere gratuitamente, mediante accesso al Registro Generale del Garante, l'esistenza di trattamenti di dati che possono riguardarlo; 
Ottenere da Cristalensi s.n.c., - con un contributo spese solo in caso di risposta negativa - la conferma dell'esistenza o meno nei propri archivi di dati che lo riguardino, ed avere la loro comunicazione e l'indicazione della logica e delle finalita' su cui si basa il trattamento. La richiesta e' rinnovabile dopo novanta giorni; 
Ottenere la cancellazione, la trasformazione in forma anonima ed il blocco dei dati trattati in violazione di legge; 
Ottenere l'aggiornamento, la rettifica o l'integrazione dei dati; 
Ottenere l'attestazione che la cancellazione, l'aggiornamento, la rettifica o l'integrazione siano portate a conoscenza di coloro che abbiano avuto comunicazione dei dati; 
Opporsi gratuitamente al trattamento dei dati che lo riguardano.</textarea>
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <input name="chekka" type="checkbox" onClick="accetta(this)" />
                                            Accetta le condizioni </div>
                                    </div>
                                    <div class="tr">
                                        <div>&nbsp;</div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <input name="Submit" type="submit" class="form" value="Salva" align="absmiddle" disabled />
                                            &nbsp;
                                            <input name="Submit2" type="reset" class="form" value="Annulla" />
                                            (*) campo obbligatorio </div>
                                    </div>
                                    <div class="tr">
                                        <div>&nbsp;</div>
                                    </div>
                                    </form>
                                </div>
                                
                                <%else%>
                                 	<h3 style="font-size: 14px; display: inline; border: none;">Autenticazione cliente</h3>
                                    <div class="iscrizione clearfix">
									  <%if mode=2 then%>
                                          <br /><br />
                                          <font color="#990000"><b>Benvenuto/a&nbsp;<%=nome_log%></b></font>
                                          <br /><br />
                                          <a href="<%if italia_log="S&igrave;" then%>carrello2.asp<%end if%><%if italia_log="No" then%>carrello2extra.asp<%end if%>">cliccando qui tornerai direttamente all'ordine da completare: Carrello on-line &raquo;</a>
                                      <%else%>
                                  
                                      <%if pkid=0 then%>
                                          Grazie per esserti iscritto/a a Cristalensi.it<br />
                                          Da questo momento potrai completare e inviare i tuoi ordini senza dover inserire tutti i tuoi dati.
                                          <br /><br />
                                          <a href="<%if italia_log="S&igrave;" then%>carrello2.asp<%end if%><%if italia_log="No" then%>carrello2extra.asp<%end if%>">Cliccando qui tornerai direttamente all'ordine da completare: Carrello on-line &raquo;</a>
                                      <%else%>
                                          I tuoi dati sono stati aggiornati regolarmente.
                                      <%end if%>
                                    </div>
                              <%end if%>

                            	
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