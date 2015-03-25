<%
produttore=request("produttore")
id=request("id")

mode=request("mode")
if mode="" then mode=0
if mode=1 then
	email=request("email")
	nome=request("nome")
	cognome=request("cognome")
	telefono=request("telefono")
	richiesta=request("richiesta")
	
	'Response.CacheControl = "no-cache"
	'Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	Function CheckCAPTCHA(valCAPTCHA)
		SessionCAPTCHA = Trim(Session("CAPTCHA"))
		Session("CAPTCHA") = vbNullString
		if Len(SessionCAPTCHA) < 1 then
			CheckCAPTCHA = False
			exit function
		end if
		if CStr(SessionCAPTCHA) = CStr(valCAPTCHA) then
			CheckCAPTCHA = True
		else
			CheckCAPTCHA = False
		end if
	End Function
	
	strCAPTCHA = Trim(Request.Form("strCAPTCHA"))
	if CheckCAPTCHA(strCAPTCHA) = true then
		mode=1
		'response.Write("captcha fatto!<br>")
	else
		mode=2
	end if
end if


if mode=1 then
	if email<>"" then
		ip=Request.ServerVariables("REMOTE_ADDR")
		data=date()
		
		HTML1 = ""
		HTML1 = HTML1 & "<html>"
		HTML1 = HTML1 & "<head>"
		HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
		HTML1 = HTML1 & "<title>Cristalensi</title>"
		HTML1 = HTML1 & "</head>"
		HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		HTML1 = HTML1 & "<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td>"
		HTML1 = HTML1 & "<font face=Verdana size=1 color=#000000>E' stata fatta la seguente richiesta di informazioni sui prodotti di un produttore dal sito il "&data&"<br><br>Dati della richiesta:<br>Nome: <b>"&nome&"</b><br>Cognome: <b>"&cognome&"</b><br>Telefono: <b>"&telefono&"</b><br>E-mail: <b>"&email&"</b><br>IP connessione: <b>"&ip&"</b><br><br>Produttore: <b>"&produttore&"</b><br>Codice progressivo produttore: <b>"&id&"</b><br><br>Richiesta:<br><b>"&richiesta&"</b></font>"
		HTML1 = HTML1 & "</td></tr></table>"
		HTML1 = HTML1 & "</body>"
		HTML1 = HTML1 & "</html>"
			
		Destinatario = "info@cristalensi.it"
		Mittente = "info@cristalensi.it"
		Oggetto = "Richiesta informazioni sul produttore: "&produttore
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
		
		Destinatario = "viadeimedici@gmail.com"
		Mittente = "info@cristalensi.it"
		Oggetto = "Richiesta informazioni sul produttore: "&produttore
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
		
		
		'email di conferma per il cliente
		HTML1 = ""
		HTML1 = HTML1 & "<html>"
		HTML1 = HTML1 & "<head>"
		HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
		HTML1 = HTML1 & "<title>Cristalensi</title>"
		HTML1 = HTML1 & "</head>"
		HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		HTML1 = HTML1 & "<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td>"
		HTML1 = HTML1 & "<font face=Verdana size=1 color=#000000>E' stata inviata la seguente richiesta di informazioni dal sito Cristalensi.it il "&data&"<br><br>Dati della richiesta:<br>Nome: <b>"&nome&"</b><br>Cognome: <b>"&cognome&"</b><br>Telefono: <b>"&telefono&"</b><br>E-mail: <b>"&email&"</b><br><br>Produttore: <b>"&produttore&"</b><br><br>Richiesta:<br><b>"&richiesta&"</b><br><br><br><br>Questa è un'email di conferma dell'invio della richiesta di preventivo.<br><br>La ringraziamo per aver scelto i prodotti di Cristalensi</font>"
		HTML1 = HTML1 & "</td></tr></table>"
		HTML1 = HTML1 & "</body>"
		HTML1 = HTML1 & "</html>"
			
		Destinatario = email
		Mittente = "info@cristalensi.it"
		Oggetto = "Richiesta informazioni sul produttore: "&produttore
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
		
	else
		mode=3
	end if

end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Cristalensi</title>
<link href="stile_stampa.css" rel="stylesheet" type="text/css">
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
<SCRIPT language="JavaScript">

function control() {
		
	email = document.modulo.email.value;
	privacy = document.getElementById("chekka");

	if (email==""){
		alert("Non  e\' stato compilato il campo \"E-mail\".");
		return false;
	}
	if (email.indexOf("@")==-1 || email.indexOf(".")==-1){
    alert("ATTENZIONE! \"L'E-mail\" deve essere completa di tutti i caratteri.");
    return false; 
    }
	if (! (privacy.checked)){
		alert("Non  e\' possibile proseguire senza aver accettato le condizioni.");
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
	<div>
		<h1>RICHIESTA INFORMAZIONI E PREVENTIVI</h1>
		<%if mode=1 then%>
		<br /><br /><br /><br /><br /><br /><b>La richiesta è stata inoltrata correttamente,<br />
			il nostro staff ti contatter&agrave; il prima possibile.<br />Saluti da CRISTALENSI</b>
		<%else%>	
		<p>Al momento non sono esposti sul sito internet prodotti di <%=Produttore%>, ma abbiamo comunque a disposizione il loro catagolo e vendiamo i loro prodotti nel nostro negozio. Quindi se conosci un articolo di questo produttore e vuoi avere un preventivo di prezzo <b>riempi il seguente modulo</b> indicandoci il prodotto, <br />oppure <b>contattaci direttamente</b>, il nostro staff sar&agrave; a Tua disposizione per qualsiasi chiarimento. <br /><img src="immagini/numeroverde_orizzontale_cristalensi.gif" width="330" height="85" hspace="5" vspace="10" align="middle"> 
		  <br />
	      <br />Produttore di cui stai chiedendo informazioni: <b><%=Produttore%></b></p>
		<table width="100%" border="0" cellpadding="5" cellspacing="0">
		<form method="post" name="modulo" action="richiesta_informazioni_produttore.asp" onSubmit="return control();">
			<input type="hidden" name="mode" value="1" />
			<input type="hidden" name="produttore" value="<%=produttore%>" />
			<input type="hidden" name="id" value="<%=id%>" />
			<tr> 
			<td align="center" valign="top">
				<table cellpadding="0" cellspacing="0" width="100%">
				<%if mode=3 then%>
				<tr> 
				<td align="center" height="30" colspan="4">
				<font color="#CC0000"><b>Attenzione! ci sono stati dei problemi al sistema: riprovare l'inserimento, grazie.</b>			</font></td>
				</tr>
				<%end if%>
				<tr>
					<td width="27%" height="30" align="right">
					Nome:&nbsp;</td>
					<td width="26%" height="30" align="left"><input type="text" name="nome" id="nome" size="30" value="<%=nome%>" class="form" /></td>
				    <td width="12%" align="right">Cognome:&nbsp;</td>
			      <td width="35%" align="left"><input type="text" name="cognome" id="cognome" size="30" value="<%=cognome%>" class="form" /></td>
				</tr>
				<tr>
					<td width="27%" height="30" align="right">
					<strong>Email</strong> (OBBLIGATORIA):&nbsp;					</td>
					<td width="26%" height="30" align="left"><input type="text" name="email" id="email" size="30" value="<%=email%>" class="form" />					</td>
				    <td width="12%" align="right">Telefono:&nbsp;</td>
			      <td width="35%" align="left"><input type="text" name="telefono" id="telefono" size="30" value="<%=telefono%>" class="form" /></td>
				</tr>
				
				<tr>
				<td height="30" align="center" colspan="4">Ulteriore richiesta informazioni</td>
				</tr>
				<tr>
				<td colspan="4" align="center">
				<textarea name="richiesta" cols="60" rows="5" class="form"><%=richiesta%></textarea>				</td>
				</tr>	
				<%if mode=2 then%>
				<tr> 
				<td align="center" height="30" colspan="4">
				<font color="#CC0000"><b>Attenzione! Il codice inserito non è corretto: riprovare l'invio, grazie.</b>			</font></td>
				</tr>
				<%end if%>
				<tr>
				<td height="40" align="center" colspan="4">
				<img src="aspcaptcha.asp" alt="This Is CAPTCHA Image" width="86" height="21" />				</td>
				</tr>
				<tr>
				<td height="30" align="center" colspan="4">Per effettuare la richiesta in sicurezza è necessario inserire il codice indicato sopra				</td>
				</tr>
				<tr>
				<td height="30" align="right" colspan="2">Codice: </td>
			  	<td height="40" colspan="2" align="left"><input name="strCAPTCHA" type="text" id="strCAPTCHA" maxlength="8" class="form" /></td>
				</tr>
				<tr>
				<td colspan="4" align="center">
				<textarea name="privacy" cols="60" rows="5" readonly="readonly" class="form">Informativa sulla Privacy. La informiamo  che i dati da lei conferiti come mittente, mediante la compilazione dei campi elettronici sopra individuati, saranno trattati da Cristalensi s.n.c. , in qualità di Titolare del trattamento, ai sensi dell ´art. 13 del Codice in materia di protezione dei dati personali.
I suoi dati saranno trattati con mezzi informatici nel rispetto dei principi stabiliti dal Codice della Privacy D.Lgs. 196/2003.
Per informazioni sulle modalità del trattamento e per esercitare i diritti a lei riconosciuti dall’art 7 del Codice della Privacy, potrà rivolgersi al Titolare della Cristalensi s.n.c. presso la sede di via arti e mestieri, 1 - 50056 Montelupo Fiorentino (FI), oppure all’indirizzo di posta elettronica info@cristalensi.it. Potrà comunque successivamente, in ogni momento, chiedere la cancellazione dei suoi dati, rivolgendosi al Titolare o al Responsabile del Trattamento.</textarea>				</td>
				</tr>		  
				<tr>
				<td height="25" align="right" colspan="2"><input type="submit" name="entra" value="Invia la richiesta" class="button" disabled /></td>
				<td height="40" colspan="2" align="left"><input name="chekka" id="chekka" type="checkbox" onClick="accetta(this)"> Accetta le condizioni</td>
				</tr>
				</table>
			</td>
			</tr>
		  </form>
	  </table>
			<%end if%>
	<p align="center"><a href="#" onclick="self.close();">[CHIUDI LA FINESTRA]</a></p>
</div>
</body>
</html>
