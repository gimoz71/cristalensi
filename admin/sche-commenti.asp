<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
<%
	pkid = request("pkid")
	if pkid = "" then pkid = 0
	
	p = request("p")
	if p = "" then p = 1
	ordine = request("ordine")
	if ordine = "" then ordine = 0
	
	mode = request("mode")
	if mode = "" then mode = 0

	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Commenti_Clienti"
	if pkid > 0 then sql = "Select * From Commenti_Clienti where pkid="&pkid
	rs.Open sql, conn, 3, 3
	
	if mode = 1 then
		if pkid = 0 then rs.addnew
		
		rs("fkiscritto")=request("fkiscritto")
		
			Set cs=Server.CreateObject("ADODB.Recordset")
			sql = "Select PkId, Nominativo From Clienti WHERE PkId="&request("fkiscritto")
			cs.Open sql, conn, 1, 1
			if cs.recordcount>0 then
				Nominativo=cs("Nominativo")
				Nome=cs("Nome")
			end if
			cs.close
		
		pubblicato=request("pubblicato")
		if pubblicato="si" then rs("pubblicato")=True
		if pubblicato="no" then rs("pubblicato")=False
		rs("pubblicato")=pubblicato
		
		risposta=request("risposta")
		if risposta="si" then rs("risposta")=True
		if risposta="no" then rs("risposta")=False
		rs("risposta")=risposta
		
		rs("testo")=request("testo")
		rs("data")=now()
		
		Notifica_pub=request("Notifica_pub")
		if Notifica_pub="si" then
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Spett.le "&nome&" "&nominativo&", lo staff di Cristalensi ha pubblicato il commento inserito.<br><br>Potrà vederlo andando direttamente sul sito internet alla <a href=""http://www.cristalensi.it"">pagina dei commenti</a>.</font>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td><br><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dear "&nome&" "&nominativo&", the staff of Cristalensi has published the comment inserted.<br><br> You could see it by going directly to the website<a href=""http://www.cristalensi.it""> in the page of feed-back</a></font>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best regards, from the staff of Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Cristalensi.it: pubblicato il commento - published the comment"
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
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Cristalensi.it: pubblicato il commento - published the comment"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set eMail_cdo = Nothing
			'fine invio email
			
			'invio l'email al webmaster
			Mittente = "info@cristalensi.it"
			Destinatario = "iurymazzoni@hotmail.com"
			Oggetto = "Cristalensi.it: pubblicato il commento - published the comment"
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
		
		if request("C1") = "ON" then
			
			'qui devono essere inserite tutte le tabelle dove compare FkCat_Prod per cancellare il record oppure metterlo a 0
			Set ss=Server.CreateObject("ADODB.Recordset")
			sql = "Select FkCommento From Commenti_Risposte where FkCommento="&pkid&""
			ss.Open sql, conn, 3, 3
				if ss.recordcount>0 then
					Do while not ss.EOF
						ss("FkCommento")=0
						ss.update
					ss.movenext
					loop
				end if
			ss.close
			
			rs.delete
		end if
		rs.update
		
		rs.close
		
		'aggiunta/modifica risposta
		if risposta="si" then
			pkid_risposta=request("pkid_risposta")
			if pkid_risposta="" then pkid_risposta=0
			
			if pkid=0 then
				Set os2 = Server.CreateObject("ADODB.Recordset")
				sql = "Select @@Identity As pkid"
				os2.Open sql, conn, 1, 1
				pkid=os2("pkid")
				os2.close
			end if
			
			Set risp_rs=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From Commenti_Risposte"
			if pkid_risposta > 0 then sql = "Select * From Commenti_Risposte where pkid="&pkid_risposta
			risp_rs.Open sql, conn, 3, 3
			
			if pkid_risposta = 0 then risp_rs.addnew
			
			risp_rs("fkamministratore")=idadmin
			risp_rs("fkcommento")=pkid
			
			risp_rs("pubblicato")=True
			
			risp_rs("testo")=request("testo_risposta")
			risp_rs("data")=now()
			
			risp_rs.update
			risp_rs.close
			
			Notifica_risp=request("Notifica_risp")
			if Notifica_risp="si" then
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
				HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Spett.le "&nome&" "&nominativo&", lo staff di Cristalensi ha pubblicato una risposta al commento inserito.<br><br>Potrà vederla andando direttamente sul sito internet alla <a href=""http://www.cristalensi.it"">pagina dei commenti</a>.</font>"
				HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"
				HTML1 = HTML1 & "<tr>"
				HTML1 = HTML1 & "<td><br><br>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"
				HTML1 = HTML1 & "<tr>"
				HTML1 = HTML1 & "<td>"
				HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dear "&nome&" "&nominativo&", the staff of Cristalensi has published a response to the comment inserted.<br><br> You could see it by going directly to the website<a href=""http://www.cristalensi.it""> in the page of feed-back</a></font>"
				HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Best regards, from the staff of Cristalensi</font>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"
				HTML1 = HTML1 & "</table>"
				HTML1 = HTML1 & "</body>"
				HTML1 = HTML1 & "</html>"
	
				Mittente = "info@cristalensi.it"
				Destinatario = email
				Oggetto = "Cristalensi.it: inserita risposta al commento - inserted response to the comment"
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
				Mittente = "info@cristalensi.it"
				Destinatario = "info@cristalensi.it"
				Oggetto = "Cristalensi.it: inserita risposta al commento - inserted response to the comment"
				Testo = HTML1
	
				Set eMail_cdo = CreateObject("CDO.Message")
	
				eMail_cdo.From = Mittente
				eMail_cdo.To = Destinatario
				eMail_cdo.Subject = Oggetto
	
				eMail_cdo.HTMLBody = Testo
	
				eMail_cdo.Send()
	
				Set eMail_cdo = Nothing
				'fine invio email
				
				'invio l'email al webmaster
				Mittente = "info@cristalensi.it"
				Destinatario = "iurymazzoni@hotmail.com"
				Oggetto = "Cristalensi.it: inserita risposta al commento - inserted response to the comment"
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
		end if
	end if
	
	if pkid>0 then
		if rs("risposta")=True then
			Set risp_rs=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From Commenti_Risposte where fkcommento="&pkid
			risp_rs.Open sql, conn, 3, 3
		end if
	end if
	
	
	
%>
<html>
<head>
<title>Cristalensi Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0" class="TAB_centrale">
  <!--#include file="testata.asp"-->
  <tr>
    <td height="20" colspan="2" valign="middle"><table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="159" class="menu-celle">&nbsp;Menu</td>
          <td width="267" class="menu-celle">Gestione Commenti</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-commenti.asp">Elenco Commenti &raquo;</a>&nbsp;&nbsp;<a href="sche-commenti.asp">Nuovo commento &raquo;</a></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td colspan="2" valign="top"><table width="750" height="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="150" class="admin-menu" valign="top">
		<!--#include file="sinistra.asp"-->
		 </td>
        <td align="center" valign="top">
          <!--tab centrale-->
			<% if request("C1") <> "ON" then %>
        <% if mode = 1 and pkid = 0 then %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Commento Inserito ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-commenti.asp?ordine=<%=ordine%>";
			   					}
							//-->
							</script>
                <% else %>
        <% if mode = 1 then %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Commento Aggiornato ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
								<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-commenti.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
								//-->
								</script>
                <% else %>
				<table cellpadding="0" cellspacing="0" border="0" width="95%" class="admin-righe">
				  <tr> 
                	<td colspan="2">&nbsp;</td>
              	</tr> 	
					<form method="post" action="sche-commenti.asp?mode=1&pkid=<%=pkid%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                  <tr align="left">
                    <td width="40%" height="15"><strong>Pubblicato</strong></td>
                    <td width="60%" height="15"><strong>Invia Email per pubblicazione</strong> </td>
                  </tr>
				  <tr align="left">
                    <td height="15">
					Si 
				    <input name="Pubblicato" type="radio" value="si" <% if pkid > 0 then %><%if rs("Pubblicato")=True then%>checked<%end if%><%end if%>>&nbsp;&nbsp;No <input name="Pubblicato" type="radio" value="no" <% if pkid > 0 then %><%if rs("Pubblicato")=False then%>checked<%end if%><%else%>checked<%end if%>></td>
                    <td height="15">Si 
				    <input name="Notifica_pub" type="radio" value="si" />&nbsp;&nbsp;No <input name="Notifica_pub" type="radio" value="no" checked /></td>
				  </tr>
				  <tr align="left">
                    <td colspan="2" height="15"><strong>Cliente</strong> </td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><%
					Set cs=Server.CreateObject("ADODB.Recordset")
					sql = "Select * From Clienti order by Nominativo ASC"
					cs.Open sql, conn, 1, 1
					%>
					<select name="FkIscritto" class="form">
                        <%
						if cs.recordcount>0 then
						Do While Not cs.EOF
						%>
                        <option value=<%=cs("pkid")%> <% if pkid > 0 then %><%if rs("FkIscritto")=cs("pkid") then%> selected<%end if%><%end if%>><%=cs("Nominativo")%>&nbsp;<%=cs("Nome")%></option>
                        <%
						cs.movenext
						loop
						end if
						%>
                     </select>
					 <%cs.close%></td>
				  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Testo commento</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<textarea name="testo" cols="78" rows="5" class="form"><%if pkid>0 then%><%=rs("testo")%><%end if%></textarea></td>
                  </tr>
				  
                  <tr align="left">
                    <td height="15"><strong>Risposta</strong></td>
                    <td height="15"><strong>Invia Email per risposta</strong></td>
                    </tr>
				  <tr align="left">
                    <td height="15">
					Si 
				    <input name="Risposta" type="radio" value="si" <% if pkid > 0 then %><%if rs("Risposta")=True then%>checked<%end if%><%end if%>>&nbsp;&nbsp;No <input name="Risposta" type="radio" value="no" <% if pkid > 0 then %><%if rs("Risposta")=False then%>checked<%end if%><%else%>checked<%end if%>><%if rs("Risposta")=True then%><input type="hidden" name="pkid_risposta" value="<%=risp_rs("PkId")%>" /><%end if%></td>
                    <td height="15">Si 
				    <input name="Notifica_risp" type="radio" value="si" />&nbsp;&nbsp;No <input name="Notifica_risp" type="radio" value="no" checked /></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Testo risposta</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<textarea name="testo_risposta" cols="78" rows="5" class="form"><%if pkid>0 and rs("Risposta")=True and risp_rs("PkId")>0 then%><%=risp_rs("Testo")%><%end if%></textarea></td>
                  </tr>
                  
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">
					<input name="Submit" type="submit" class="form" value="Salva" align="absmiddle"> 
                          &nbsp; <input name="Submit2" type="reset" class="form" value="Annulla"> 
                          &nbsp; <input type="checkbox" name="C1" value="ON" > 
                          &nbsp; Per cancellare il commento </td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                </form>
				</table>
				<% end if %>
                <% end if %>
        <% else %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Commento Cancellato ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-commenti.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
							//-->
						</script>
                <% end if %>
			<!--fine tab-->
		  </td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<!--#include file="strClose.asp"-->
<!--#include file="chiusura.asp"-->