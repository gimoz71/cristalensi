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
		HTML1 = HTML1 & "<font face=Verdana size=1 color=#000000>E' stata fatta la seguente richiesta di informazioni sui prodotti di un produttore dal sito (in inglese) il "&data&"<br><br>Dati della richiesta:<br>Nome: <b>"&nome&"</b><br>Cognome: <b>"&cognome&"</b><br>Telefono: <b>"&telefono&"</b><br>E-mail: <b>"&email&"</b><br>IP connessione: <b>"&ip&"</b><br><br>Produttore: <b>"&produttore&"</b><br>Codice progressivo produttore: <b>"&id&"</b><br><br>Richiesta:<br><b>"&richiesta&"</b></font>"
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
		HTML1 = HTML1 & "<font face=Verdana size=1 color=#000000>The following request for informations has been sent by the site Cristalensi.it on the "&data&"<br><br>Request data:<br>Name: <b>"&nome&"</b><br>Surname: <b>"&cognome&"</b><br>Telephone: <b>"&telefono&"</b><br>E-mail: <b>"&email&"</b><br><br>Producer: <b>"&produttore&"</b><br><br>Request:<br><b>"&richiesta&"</b><br><br><br><br>This e-mail is confirmation that a request for  an estimate has been sent.<br><br>Thank you for having chosen Cristalensi's products</font>"
		HTML1 = HTML1 & "</td></tr></table>"
		HTML1 = HTML1 & "</body>"
		HTML1 = HTML1 & "</html>"
			
		Destinatario = email
		Mittente = "info@cristalensi.it"
		Oggetto = "Request informations about producer: "&produttore
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
		alert("It has not been filled in the field \"E-mail\".");
		return false;
	}
	if (email.indexOf("@")==-1 || email.indexOf(".")==-1){
    alert("ATTENTION! E-mail address must be complete with all the characters.");
    return false; 
    }
	if (! (privacy.checked)){
		alert("You can not go on without accepting the conditions.");
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
		<h1>REQUEST FOR INFORMATION AND ESTIMATES</h1>
		<%if mode=1 then%>
		<br /><br /><br /><br /><br /><br /><b>The request has been sent correctly.<br />Our staff will contact you as soon as possible.<br />Best wishes from the staff of Cristalensi</b>
		<%else%>	
		<p>At present product <%=Produttore%> is not shown on the site, nonetheless we have it available in our catalogue, and we sell their products in our shop.  Therefore, if you know an article made by this producer and you would like a price quote please fill-in the following form, our staff is happy to provide any clarifications.
          
		  <br />
	      <br />Producer about whom you are asking information: <b><%=Produttore%></b></p>
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
				<font color="#CC0000"><b>Attention! There have been problems with the system:  please reinsert data and try again, thanks.</b>			</font></td>
				</tr>
				<%end if%>
				<tr>
					<td width="27%" height="30" align="right">
					Name:&nbsp;</td>
					<td width="26%" height="30" align="left"><input type="text" name="nome" id="nome" size="30" value="<%=nome%>" class="form" /></td>
				    <td width="12%" align="right">Surname:&nbsp;</td>
			      <td width="35%" align="left"><input type="text" name="cognome" id="cognome" size="30" value="<%=cognome%>" class="form" /></td>
				</tr>
				<tr>
					<td width="27%" height="30" align="right">
					<strong>Email</strong> (OBLIGATORY):&nbsp;					</td>
					<td width="26%" height="30" align="left"><input type="text" name="email" id="email" size="30" value="<%=email%>" class="form" />					</td>
				    <td width="12%" align="right">Telephone:&nbsp;</td>
			      <td width="35%" align="left"><input type="text" name="telefono" id="telefono" size="30" value="<%=telefono%>" class="form" /></td>
				</tr>
				
				<tr>
				<td height="30" align="center" colspan="4">Further information request</td>
				</tr>
				<tr>
				<td colspan="4" align="center">
				<textarea name="richiesta" cols="60" rows="5" class="form"><%=richiesta%></textarea>				</td>
				</tr>	
				<%if mode=2 then%>
				<tr> 
				<td align="center" height="30" colspan="4">
				<font color="#CC0000"><b>Attention!  The code inserted is incorrect: please send again, thank you.</b>			</font></td>
				</tr>
				<%end if%>
				<tr>
				<td height="40" align="center" colspan="4">
				<img src="aspcaptcha.asp" alt="This Is CAPTCHA Image" width="86" height="21" />				</td>
				</tr>
				<tr>
				<td height="30" align="center" colspan="4">To make a request you must enter the security code above </td>
				</tr>
				<tr>
				<td height="30" align="right" colspan="2">Code: </td>
			  	<td height="40" colspan="2" align="left"><input name="strCAPTCHA" type="text" id="strCAPTCHA" maxlength="8" class="form" /></td>
				</tr>
				<tr>
				<td colspan="4" align="center">
				<textarea name="privacy" cols="60" rows="5" readonly="readonly" class="form">INFORMATION RELATIVE TO THE TREATMENT OF PERSONAL DATA.  According to the sense of article 10 of the Law number 675 of the 31/12/1996, the Company informs the interested party that the data regarding it, furnished by the same, will be subject to treatment with  respect to the above mentioned norm.  These data will be used to scopes of a gestional, commercial, and promotional nature. The releasing of data to our Company is entirely optional. Data acquired may be communicated and diffused in observation of the dispositions contained in article 20 of Law 675/96 in persuance of the finalities above mentioned.  The owner of the treatment is Cristalensi s.n.c. whose seat is in via arti e mestieri 1 Montelupo F.no (Fi) where, moreover, the responsable  pro tempore of the treatment resides, whose identity can be obtained from the Public Register held by the Garante, or from the legal offices of the Company.  In addition, the Company informs interested parties that they may exercise the rights  foreseen in article 13 of the law 675/96, that is:  know  without charge, through the General Register of the Garante, the treatments of data which concern the interested party;  there can be obtained from Cristalensi s.n.c., - with a contribution towards the costs only in the case of a negative response- the confirmation or negation of the existence, in the company archives, of data which regard the interested party, and can have access to information regarding the finalities to which the data have been put. The request is renewable after ninty days;  Obtain the cancellation, the transformation into anonymous form and the blocking of data treated in violation of the sense of the law; Obtain the updating, the correction or the integration of the data;  Object, without charge, to the treatment of data which concerns the interested party.</textarea>				</td>
				</tr>		  
				<tr>
				<td height="25" align="right" colspan="2"><input type="submit" name="entra" value="Submit" class="button" disabled /></td>
				<td height="40" colspan="2" align="left"><input name="chekka" id="chekka" type="checkbox" onClick="accetta(this)"> Accept the conditions</td>
				</tr>
				</table>
			</td>
			</tr>
		  </form>
	  </table>
			<%end if%>
	<p align="center"><a href="#" onclick="self.close();">[CLOSE THE WINDOW]</a></p>
</div>
</body>
</html>
