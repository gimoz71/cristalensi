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
	sql = "Select * From Produttori"
	if pkid > 0 then sql = "Select * From Produttori where pkid="&pkid
	rs.Open sql, conn, 3, 3
	
	if mode = 1 then
		if pkid = 0 then rs.addnew
		
		rs("titolo")=request("titolo")
		rs("descrizione")=request("descrizione")
		rs("logo")=request("allegato")
		rs("prodotti")=request("prodotti")
		
		if request("C1") = "ON" then
			
			'qui devono essere inserite tutte le tabelle dove compare FkCat_Prod per cancellare il record oppure metterlo a 0
			Set ss=Server.CreateObject("ADODB.Recordset")
			sql = "Select FkProduttore From Prodotti where FkProduttore="&pkid&""
			ss.Open sql, conn, 3, 3
				if ss.recordcount>0 then
					Do while not ss.EOF
						ss("FkProduttore")=0
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
          <td width="267" class="menu-celle">Gestione Produttori</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-produttori.asp">Elenco Produttori &raquo;</a>&nbsp;&nbsp;<a href="sche-produttori.asp">Nuova produttore &raquo;</a></td>
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
                <p class="admin-righe"> Produttore Inserito ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-produttori.asp?ordine=<%=ordine%>";
			   					}
							//-->
							</script>
                <% else %>
                <% if mode = 1 then %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Produttore Aggiornato ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
								<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-produttori.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
								//-->
								</script>
                <% else %>
				<table cellpadding="0" cellspacing="0" border="0" width="95%" class="admin-righe">
				  <tr> 
                	<td colspan="2">&nbsp;</td>
              	</tr> 	
					<form method="post" action="sche-produttori.asp?mode=1&pkid=<%=pkid%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Titolo</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="titolo" class="form" size="80" maxlength="100" <%if pkid>0 then%> value="<%=rs("titolo")%>"<%end if%>></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Descrizione</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<textarea name="descrizione" cols="78" rows="5" class="form"><%if pkid>0 then%><%=rs("descrizione")%><%end if%></textarea></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Logo</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="Allegato" id="Allegato" class="form" size="30" maxlength="100" <%if pkid>0 then%> value="<%=rs("Logo")%>"<%end if%>> per inserire un file, <a href="#" onClick="MM_openBrWindow('upload-file.asp','','width=300,height=300')">cliccare qui.</a></td>
                  </tr>
				  <tr align="left">
                    <td colspan="2"><strong>Il produttore ha prodotti esposti?</strong></td>
                  </tr>
				  <tr align="left">
                    <td width="328">Si 
                    <input name="Prodotti" type="radio" value="1" <% if pkid > 0 then %><%if rs("Prodotti")=1 then%>checked<%end if%><%end if%>>&nbsp;&nbsp;No <input name="Prodotti" type="radio" value="0" <% if pkid > 0 then %><%if rs("Prodotti")=0 then%>checked<%end if%><%else%>checked<%end if%>></td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">
					<input name="Submit" type="submit" class="form" value="Salva" align="absmiddle"> 
                          &nbsp; <input name="Submit2" type="reset" class="form" value="Annulla"> 
                          &nbsp; <input type="checkbox" name="C1" value="ON" > 
                          &nbsp; Per cancellare il produttore </td>
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
                <p class="admin-righe"> Produttore Cancellato ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-produttori.asp?p=<%=p%>&ordine=<%=ordine%>";
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