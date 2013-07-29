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

	if mode=1 then
		nominativo=request("nominativo")
		nome=request("nome")
		Rag_Soc=request("Rag_Soc")
		Cod_Fisc=request("Cod_Fisc")
		PartitaIVA=request("PartitaIVA")
		indirizzo=request("indirizzo")
		cap=request("cap")
		citta=request("citta")
		provincia=request("fkprovincia")
		italia=request("italia")
		nazionediversa=request("nazionediversa")
		Telefono=request("Telefono")
		Fax=request("Fax")
		email=request("email")
		Aut_email=request("Aut_email")
		password=request("Password")
		Data=request("Data")
		ip=request("ip")
		Aut_privacy=request("Aut_privacy")
	end if
	'if mode=1 then
		'Set rs=Server.CreateObject("ADODB.Recordset")
		'sql = "Select nick From Iscritti where nick='"&nick&"'"
		'if pkid>0 then sql = "Select nick,pkid From Iscritti where nick='"&nick&"' and pkid<>"&pkid&""
		'rs.Open sql, conn, 1, 1
		'if rs.recordcount>0 then mode=2
		'rs.close
	'end if
	'if mode=1 then
		'Set rs=Server.CreateObject("ADODB.Recordset")
		'sql = "Select email From Iscritti where email='"&email&"'"
		'if pkid>0 then sql = "Select email,pkid From Iscritti where email='"&email&"' and pkid<>"&pkid&""
		'rs.Open sql, conn, 1, 1
		'if rs.recordcount>0 then mode=3
		'rs.close
	'end if
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Clienti"
	if pkid > 0 then sql = "Select * From Clienti where pkid="&pkid
	rs.Open sql, conn, 3, 3
	
	if mode = 1 then
		if pkid = 0 then rs.addnew
		
		rs("nominativo")=nominativo
		rs("nome")=nome
		rs("Rag_Soc")=Rag_Soc
		rs("Cod_Fisc")=Cod_Fisc
		rs("PartitaIVA")=PartitaIVA
		rs("indirizzo")=indirizzo
		rs("cap")=cap
		rs("citta")=citta
		rs("provincia")=provincia
		rs("italia")=italia
		rs("nazionediversa")=nazionediversa
		rs("Telefono")=Telefono
		rs("Fax")=Fax
		rs("email")=email
		rs("Aut_email")=Aut_email
		rs("password")=password
		rs("Data")=Data
		rs("ip")=ip
		rs("Aut_privacy")=Aut_privacy
		
		if request("C1") = "ON" then
			
			'qui devono essere inserite tutte le tabelle dove compare fkcliente per cancellare il record oppure metterlo a 0
			Set ss=Server.CreateObject("ADODB.Recordset")
			sql = "Select fkcliente From Ordini where fkcliente="&pkid&""
			ss.Open sql, conn, 3, 3
				if ss.recordcount>0 then
					Do while not ss.EOF
						ss("fkcliente")=0
						ss.update
					ss.movenext
					loop
				end if
			ss.close
			
			Set ss=Server.CreateObject("ADODB.Recordset")
			sql = "Select fkcliente From RigheOrdine where fkcliente="&pkid&""
			ss.Open sql, conn, 3, 3
				if ss.recordcount>0 then
					Do while not ss.EOF
						ss("fkcliente")=0
						ss.update
					ss.movenext
					loop
				end if
			ss.close
			
			Set ss=Server.CreateObject("ADODB.Recordset")
			sql = "Select fkiscritto From Commenti_Clienti where fkiscritto="&pkid&""
			ss.Open sql, conn, 3, 3
				if ss.recordcount>0 then
					Do while not ss.EOF
						ss("fkiscritto")=0
						ss.update
					ss.movenext
					loop
				end if
			ss.close
			
			rs.delete
		end if
		rs.update
	end if
	
%>
<html>
<head>
<title>Cristalensi Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0" class="TAB_centrale">
  <!--#include file="testata.asp"-->
  <tr>
    <td height="30" colspan="2" valign="middle"><table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="159" class="menu-celle">&nbsp;Menu</td>
          <td width="267" class="menu-celle">Gestione clienti</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-iscritti.asp">Elenco clienti &raquo;</a>&nbsp;&nbsp;<a href="sche-iscritti.asp">Nuovo cliente &raquo;</a></td>
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
                <p> Cliente Inserito ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-iscritti.asp?ordine=<%=ordine%>";
			   					}
							//-->
							</script>
                <% else %>
                <% if mode = 1 then %>
                <p>&nbsp;</p>
                <p> Cliente Aggiornato ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
								<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-iscritti.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
								//-->
								</script>
                <% else %>
				<table cellpadding="0" cellspacing="0" border="0" width="98%" class="admin-righe">
				  <tr> 
                	<td colspan="2">&nbsp;</td>
              	</tr> 	
					<form method="post" action="sche-iscritti.asp?mode=1&pkid=<%=pkid%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                  <tr align="left">
                    <td width="329">Nome e Cognome (Nominativo)</td>
                    <td width="258">Ragione sociale</td>
                  </tr>
                  <tr align="left">
                    <td><input name="Nome" type="text" class="form" id="Nome"  size="20" maxlength="50" <% if pkid > 0 then %> value="<%=rs("Nome")%>"<%end if %>>&nbsp;<input name="Nominativo" type="text" class="form" id="Nominativo"  size="20" maxlength="50" <% if pkid > 0 then %> value="<%=rs("Nominativo")%>"<%end if %>></td>
                    <td><input name="Rag_Soc" type="text" class="form" id="Rag_Soc"  size="20" maxlength="50" <% if pkid > 0 then %> value="<%=rs("Rag_Soc")%>"<%end if %>></td>
                  </tr>
				  <tr align="left">
                    <td>Codice Fiscale</td>
					<td>Partita IVA</td>
                  </tr>
				  <tr align="left">
                    <td><input name="Cod_Fisc" type="text" class="form" id="Cod_Fisc"  size="20" maxlength="50" <% if pkid > 0 then %> value="<%=rs("Cod_Fisc")%>"<%end if %>></td>
					<td><input name="PartitaIVA" type="text" class="form" id="PartitaIVA"  size="20" maxlength="50" <% if pkid > 0 then %> value="<%=rs("PartitaIVA")%>"<%end if %>></td>
                  </tr>
				  <tr align="left">
                    <td width="329">Indirizzo</td>
                    <td width="258">Cap</td>
                  </tr>
                  <tr align="left">
                    <td><input name="indirizzo" type="text" class="form" id="indirizzo"  size="20" maxlength="100" <% if pkid > 0 then %> value="<%=rs("indirizzo")%>"<%end if %>></td>
                    <td><input name="cap" type="text" class="form" id="cap"  size="7" maxlength="5" <% if pkid > 0 then %> value="<%=rs("cap")%>"<%end if %>></td>
                  </tr>
				  <tr align="left">
                    <td width="329">Citt&agrave;</td>
                    <td width="258">Provincia</td>
                  </tr>
                  <tr align="left">
                    <td><input name="citta" type="text" class="form" id="citta"  size="20" maxlength="50" <% if pkid > 0 then %> value="<%=rs("citta")%>"<%end if %>></td>
                    <td><input name="provincia" type="text" class="form" id="provincia"  size="3" maxlength="2" <% if pkid > 0 then %> value="<%=rs("provincia")%>"<%end if %>></td>
                  </tr>
				  <tr align="left">
				  <td height="20" colspan="2">Nazione</td>
				  </tr>
				  <tr align="left">
				  <td height="20" colspan="2">Italia:&nbsp;&nbsp;Si&nbsp;<input type="radio" name="italia" value="Sì" <% if pkid > 0 then %><%if rs("italia")="Sì" then%> checked<%end if %><%else%> checked<%end if %>>&nbsp;&nbsp;No&nbsp;<input type="radio" name="italia" value="No" <% if pkid > 0 then %><%if rs("italia")="No" then%> checked<%end if %><%end if %>>&nbsp;Altra nazione
			<input name="nazionediversa" type="text" class="form" id="nazionediversa"  size="30" maxlength="50" value="<% if pkid > 0 then %><%=rs("nazionediversa")%><%end if%>"></td>
				  </tr>
				  <tr align="left">
                    <td width="329">Telefono</td>
                    <td width="258">Fax</td>
                  </tr>
                  <tr align="left">
                    <td><input name="Telefono" type="text" class="form" id="Telefono"  size="20" maxlength="20" <% if pkid > 0 then %> value="<%=rs("Telefono")%>"<%end if %>></td>
                    <td><input name="Fax" type="text" class="form" id="Fax"  size="20" maxlength="20" <% if pkid > 0 then %> value="<%=rs("Fax")%>"<%end if %>></td>
                  </tr>
				  <tr align="left">
                    <td width="329">E-mail</td>
                    <td width="258">Password</td>
                  </tr>
                  <tr align="left">
                    <td><input name="email" type="text" class="form" id="email"  size="20" maxlength="50" <% if pkid > 0 then %> value="<%=rs("email")%>"<%end if %>></td>
                    <td><input name="password" type="text" class="form" id="password"  size="20" maxlength="20" <% if pkid > 0 then %> value="<%=rs("password")%>"<%end if %>></td>
                  </tr>
				  <tr align="left">
                    <td>Aut. Email</td>
					<td>Privacy</td>
                  </tr>
                  <tr align="left">
                    <td><input type="radio" name="Aut_email" value=True <% if pkid > 0 then %><%if rs("Aut_email")=True then%> checked<%end if %><%else%> checked<%end if %>>&nbsp;Si&nbsp;&nbsp;<input type="radio" name="Aut_email" value=False <% if pkid > 0 then %><%if rs("Aut_email")=False then%> checked<%end if %><%end if %>>&nbsp;No</td>
					<td><input type="radio" name="Aut_privacy" value=True <% if pkid > 0 then %><%if rs("Aut_privacy")=True then%> checked<%end if %><%else%> checked<%end if %>>&nbsp;Si&nbsp;&nbsp;<input type="radio" name="Aut_privacy" value=False <% if pkid > 0 then %><%if rs("Aut_privacy")=False then%> checked<%end if %><%end if %>>&nbsp;No</td>
                  </tr>
                  <tr align="left">
                    <td width="329">Iscritto il</td>
                    <td width="258">Ip</td>
                  </tr>
                  <tr align="left">
                    <td><input name="Data" type="text" class="form" id="Data"  size="20" maxlength="20" readonly  value="<% if pkid > 0 then %><%=rs("Data")%><%else%><%=now()%><%end if %>"></td>
                    <td><input name="ip" type="text" class="form" id="ip"  size="20" maxlength="15" readonly  value="<% if pkid > 0 then %><%=rs("ip")%><%else%><%=Request.ServerVariables("REMOTE_ADDR")%><%end if %>"></td>
                  </tr>
				  
                  
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">
					<input name="Submit" type="submit" class="form" value="Salva" align="absmiddle"> 
                          &nbsp; <input name="Submit2" type="reset" class="form" value="Annulla"> 
                          &nbsp; <input type="checkbox" name="C1" value="ON" > 
                          &nbsp; Per cancellare il cliente </td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                </form>
				</table>
                <!--ordine-->
                <br />
                <table width="98%"  border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
              <tr class="admin-intestazione" align="left"> 
                <td width="9%">&nbsp;<a href="ges-ordini.asp?ordine=0">Cod.</a></td>
                <td width="24%">Cliente</td>
				<td width="16%" align="center">Totale</td>
				<td width="19%" align="center">Stato</td>
                <td width="23%" align="center">Data&nbsp;<a href="ges-ordini.asp?ordine=1">0/1</a>&nbsp;<a href="ges-ordini.asp?ordine=2">1/0</a></td>
                <td width="9%" align="center">Elimina</td>
              </tr>
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
              <%
			  	Set nrs=Server.CreateObject("ADODB.Recordset")
				sql = "SELECT * "
				sql = sql + "FROM Ordini "
				sql = sql + "WHERE FkCliente="&pkid&" "
				sql = sql + "ORDER BY PkId DESC"
				nrs.Open sql, conn, 1, 1
				
			  if nrs.recordcount>0 then	
			  	Do While Not nrs.EOF
			  %>
              <tr align="left" class="admin-righe" <% if t = 1 then %>bgcolor="#CFCFCF"<% end if %>> 
                <td>&nbsp;<a href="sche-ordini.asp?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>"><font color="#CC0000"><%=nrs("pkid")%></font></a></td>
                <td><%'if Nominativo<>"" then%><%=rs("Nominativo")%>&nbsp;<%=rs("Nome")%><%'else%><!--Non iscritto--><%'end if%></td>
                <td align="center"><%if nrs("TotaleGenerale")<>"" then%><%=FormatNumber(nrs("TotaleGenerale"),2)%><%else%>0,00<%end if%>€</td>
				<td align="center">
				<%if nrs("Stato")=0 then%>iniziato<%end if%>
				<%if nrs("Stato")=1 then%>assegnato<%end if%>
				<%if nrs("Stato")=2 then%>fase spedizione<%end if%>
				<%if nrs("Stato")=12 then%>fase spedizione int.<%end if%>
				<%if nrs("Stato")=22 then%>fase pagamento int.<%end if%>
				<%if nrs("Stato")=3 then%>fase pagamento<%end if%>
				<%if nrs("Stato")=4 then%>pagato paypal<%end if%>
				<%if nrs("Stato")=5 then%>no pagato<%end if%>
				<%if nrs("Stato")=6 then%>in pagamento<%end if%>
				<%if nrs("Stato")=7 then%>in lavorazione<%end if%>
				<%if nrs("Stato")=8 then%>spedito/evaso<%end if%>
				</td>
                <td align="center">
					<%=nrs("dataAggiornamento")%>
                </td>
                <td align="center"><a href="sche-ordini.asp?mode=1&pkid=<%=nrs("pkid")%>&C1=ON&ordine=<%=ordine%>&p=<%=p%>"><font color="#CC0000">X</font></a></td>
               </tr>
              <% if t = 1 then t = 0 else t = 1 %>
              <%
				nrs.movenext
			  	loop
			  %>
              <%else%>
              <tr> 
                <td colspan="6">Nessun ordine presente</td>
              </tr>
              <%end if%>
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
              <%nrs.close%>
            </table>
            <!--fine ordine-->
				<% end if %>
                <% end if %>
                <% else %>
                <p>&nbsp;</p>
                <p> Cliente Cancellato ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-iscritti.asp?p=<%=p%>&ordine=<%=ordine%>";
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