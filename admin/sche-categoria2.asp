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
	sql = "Select * From Categorie2"
	if pkid > 0 then sql = "Select * From Categorie2 where pkid="&pkid
	rs.Open sql, conn, 3, 3
	
	if mode = 1 then
		if pkid = 0 then rs.addnew
		
		rs("fkcategoria1")=request("fkcategoria1")
		rs("posizione")=request("posizione")
		rs("titolo")=request("titolo")
		rs("descrizione")=request("descrizione")
		rs("logo")=request("allegato")
		
		rs("testo1")=request("testo1")
		rs("testo2")=request("testo2")
		
		rs("titolo_en")=request("titolo_en")
		rs("descrizione_en")=request("descrizione_en")
		rs("testo1_en")=request("testo1_en")
		rs("testo2_en")=request("testo2_en")
		
		if request("C1") = "ON" then
			
			'qui devono essere inserite tutte le tabelle dove compare FkCat_Prod per cancellare il record oppure metterlo a 0
			Set ss=Server.CreateObject("ADODB.Recordset")
			sql = "Select FkCategoria2 From Prodotti where FkCategoria2="&pkid&""
			ss.Open sql, conn, 3, 3
				if ss.recordcount>0 then
					Do while not ss.EOF
						ss("FkCategoria2")=0
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
          <td width="267" class="menu-celle">Gestione Categorie liv.2</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-categoria2.asp">Elenco Categorie liv.2 &raquo;</a>&nbsp;&nbsp;<a href="sche-categoria2.asp">Nuova categoria &raquo;</a></td>
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
			        			document.location.href = "ges-categoria2.asp?ordine=<%=ordine%>";
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
			        			document.location.href = "ges-categoria2.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
								//-->
								</script>
                <% else %>
				<table cellpadding="0" cellspacing="0" border="0" width="95%" class="admin-righe">
				  <tr> 
                	<td colspan="2">&nbsp;</td>
              	</tr> 	
					<form method="post" action="sche-categoria2.asp?mode=1&pkid=<%=pkid%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                  <tr align="left">
                    <td width="40%" height="15"><strong>Posizione</strong></td>
                    <td width="60%" height="15"><strong>Categoria liv.1</strong> </td>
                  </tr>
				  <tr align="left">
                    <td height="15">
					<input type="text" name="posizione" class="form" size="3" maxlength="3" <%if pkid>0 then%> value="<%=rs("posizione")%>"<%end if%>></td>
                    <td height="15"><%
					Set cs=Server.CreateObject("ADODB.Recordset")
					sql = "Select * From Categorie1 order by titolo ASC"
					cs.Open sql, conn, 1, 1
					%>
					<select name="FkCategoria1" class="form">
                        <%
						if cs.recordcount>0 then
						Do While Not cs.EOF
						%>
                        <option value=<%=cs("pkid")%> <% if pkid > 0 then %><%if rs("FkCategoria1")=cs("pkid") then%> selected<%end if%><%end if%>><%=cs("titolo")%></option>
                        <%
						cs.movenext
						loop
						end if
						%>
                     </select>
					 <%cs.close%></td>
				  </tr>
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
                    <td height="15" colspan="2"><strong>Fotografia</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="Allegato" id="Allegato" class="form" size="30" maxlength="100" <%if pkid>0 then%> value="<%=rs("Logo")%>"<%end if%>> per inserire un file, <a href="#" onClick="MM_openBrWindow('upload-file.asp','','width=300,height=300')">cliccare qui.</a></td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Testo 1</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="testo1" class="form" size="80" maxlength="255" <%if pkid>0 then%> value="<%=rs("testo1")%>"<%end if%>></td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Testo 2</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="testo2" class="form" size="80" maxlength="255" <%if pkid>0 then%> value="<%=rs("testo2")%>"<%end if%>></td>
                  </tr>
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Titolo ENG</strong></td>
                    </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="titolo_en" class="form" size="80" maxlength="100" <%if pkid>0 then%> value="<%=rs("titolo_en")%>"<%end if%>></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Descrizione ENG</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<textarea name="descrizione_en" cols="78" rows="5" class="form"><%if pkid>0 then%><%=rs("descrizione_en")%><%end if%></textarea></td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Testo 1 ENG</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="testo1_en" class="form" size="80" maxlength="255" <%if pkid>0 then%> value="<%=rs("testo1_en")%>"<%end if%>></td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Testo 2 ENG</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="testo2_en" class="form" size="80" maxlength="255" <%if pkid>0 then%> value="<%=rs("testo2_en")%>"<%end if%>></td>
                  </tr>
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">
					<input name="Submit" type="submit" class="form" value="Salva" align="absmiddle"> 
                          &nbsp; <input name="Submit2" type="reset" class="form" value="Annulla"> 
                          &nbsp; <input type="checkbox" name="C1" value="ON" > 
                          &nbsp; Per cancellare la categoria </td>
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
			        			document.location.href = "ges-categoria2.asp?p=<%=p%>&ordine=<%=ordine%>";
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