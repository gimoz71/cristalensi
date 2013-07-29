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
	sql = "Select * From CostiTrasporto"
	if pkid > 0 then sql = "Select * From CostiTrasporto where pkid="&pkid
	rs.Open sql, conn, 3, 3
	
	if mode = 1 then
		if pkid = 0 then rs.addnew
		
		rs("nome")=request("nome")
		rs("descrizione")=request("descrizione")
		rs("nome_en")=request("nome_en")
		rs("descrizione_en")=request("descrizione_en")
		
		if request("C1") = "ON" then
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
    <td height="20" colspan="2" valign="middle"><table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="159" class="menu-celle">&nbsp;Menu</td>
          <td width="267" class="menu-celle">Gestione Costi Trasporto</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-pagamenti.asp">Elenco Costi &raquo;</a>&nbsp;&nbsp;<a href="sche-pagamenti.asp">Nuovo Costo &raquo;</a></td>
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
                <p class="admin-righe"> Categoria Inserita ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-spedizioni.asp?ordine=<%=ordine%>";
			   					}
							//-->
							</script>
                <% else %>
                <% if mode = 1 then %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Categoria Aggiornata ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
								<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-spedizioni.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
								//-->
								</script>
                <% else %>
				<table cellpadding="0" cellspacing="0" border="0" width="95%" class="admin-righe">
				  <tr> 
                	<td colspan="2">&nbsp;</td>
              	</tr> 	
					<form method="post" action="sche-spedizioni.asp?mode=1&pkid=<%=pkid%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Nome</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="nome" class="form" size="80" maxlength="100" <%if pkid>0 then%> value="<%=rs("nome")%>"<%end if%>></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Descrizione</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="descrizione" class="form" size="80" maxlength="250" <%if pkid>0 then%> value="<%=rs("descrizione")%>"<%end if%>></td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Nome EN</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="nome_en" class="form" size="80" maxlength="100" <%if pkid>0 then%> value="<%=rs("nome_en")%>"<%end if%>></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Descrizione EN</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="descrizione_en" class="form" size="80" maxlength="250" <%if pkid>0 then%> value="<%=rs("descrizione_en")%>"<%end if%>></td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">
					<input name="Submit" type="submit" class="form" value="Salva" align="absmiddle"> 
                          &nbsp; <input name="Submit2" type="reset" class="form" value="Annulla"> 
                          &nbsp; <input type="checkbox" name="C1" value="ON" > 
                          &nbsp; Per cancellare il costo </td>
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
                <p class="admin-righe"> Categoria Cancellata ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-spedizioni.asp?p=<%=p%>&ordine=<%=ordine%>";
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