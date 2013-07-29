<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
<%
p=request("p")
if p="" then p=1
					
ordine=request("ordine")
if ordine="" then ordine=0
if ordine=0 then ord="PkId DESC"
if ordine=1 then ord="Titolo ASC"
if ordine=2 then ord="Titolo DESC"
if ordine=3 then ord="CodiceArticolo ASC"
if ordine=4 then ord="CodiceArticolo DESC"

titolo=request("titolo")
codice=request("codice")
FkProduttore=request("FkProduttore")
if FkProduttore="" then FkProduttore=0
Offerta=request("Offerta")
if Offerta="" then Offerta=0
			
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Prodotti "
if titolo<>"" then
	ricerca = "WHERE Titolo LIKE '%"&titolo&"%' "
	if FkProduttore>0 then ricerca = ricerca + "AND FkProduttore="&FkProduttore&" "
	if Offerta=10 then ricerca = ricerca + "AND Offerta=10 "
	if Offerta=1 then ricerca = ricerca + "AND (Offerta=1 OR Offerta=2) "
end if
if codice<>"" then
	ricerca = "WHERE CodiceArticolo LIKE '%"&codice&"%' "
	if FkProduttore>0 then ricerca = ricerca + "AND FkProduttore="&FkProduttore&" "
	if Offerta=10 then ricerca = ricerca + "AND Offerta=10 "
	if Offerta=1 then ricerca = ricerca + "AND (Offerta=1 OR Offerta=2) "
end if
if titolo<>"" and codice<>"" then
	ricerca = "WHERE Titolo LIKE '%"&titolo&"%' AND CodiceArticolo LIKE '%"&codice&"%' "
	if FkProduttore>0 then ricerca = ricerca + "AND FkProduttore="&FkProduttore&" "
	if Offerta=10 then ricerca = ricerca + "AND Offerta=10 "
	if Offerta=1 then ricerca = ricerca + "AND (Offerta=1 OR Offerta=2) "
end if
if titolo="" and codice="" and FkProduttore>0 then
	ricerca = "WHERE FkProduttore="&FkProduttore&" "
	if Offerta=10 then ricerca = ricerca + "AND Offerta=10 "
	if Offerta=1 then ricerca = ricerca + "AND (Offerta=1 OR Offerta=2) "
end if
if titolo="" and codice="" and FkProduttore=0 and Offerta>0 then
	if Offerta=10 then ricerca = "WHERE Offerta=10 "
	if Offerta=1 then ricerca = "WHERE (Offerta=1 OR Offerta=2) "
end if
sql = sql + ricerca
sql = sql + "ORDER BY "&ord&""
nrs.Open sql, conn, 1, 1
					
nrs.PageSize = 25
if nrs.recordcount > 0 then 
nrs.AbSolutePage = p 
maxPage = nrs.PageCount 
End if
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
          <td width="267" class="menu-celle">Gestione prodotti</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-prodotti.asp">Elenco prodotti &raquo;</a>&nbsp;&nbsp;<a href="sche-prodotti.asp">Nuovo prodotto &raquo;</a></td>
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
			<table width="98%"  border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
			  <form method="post" action="ges-prodotti.asp">
			  <tr> 
                <td colspan="6" bgcolor="#CCCCCC" class="admin-righe" height="30">&nbsp;Cerca per:&nbsp;&nbsp;<strong>Azienda</strong>
              <%
						Set cs=Server.CreateObject("ADODB.Recordset")
						sql = "Select * From Produttori order by titolo ASC"
						cs.Open sql, conn, 1, 1
						if cs.recordcount>0 then
						%>
                        <select name="FkProduttore" id="FkProduttore" class="form">
                        <option value="0">Seleziona un produttore</option>
                        <%
                        Do While Not cs.EOF
                        %>
                        <option value="<%=cs("pkid")%>" <%if cInt(FkProduttore)=cs("pkid") then%> selected<%end if%>><%=cs("titolo")%></option>
                        <%
                        cs.movenext
                        loop
                        %>
                        </select>
                        <%end if%>
						<%cs.close%>
                        &nbsp;&nbsp;<strong>Vis.</strong>: 
                    &nbsp;Tutti<input name="Offerta" type="radio" value="0" <% if offerta=0 then %>checked<%end if%>>
                    &nbsp;Non visib.<input name="Offerta" type="radio" value="10" <% if offerta=10 then %>checked<%end if%>>
                    &nbsp;In offerta<input name="Offerta" type="radio" value="1" <% if offerta=1 then %>checked<%end if%>>
                  </td>
              </tr>
              <tr> 
                <td colspan="6" bgcolor="#CCCCCC" class="admin-righe" height="30">&nbsp;&nbsp;&nbsp;<strong>Nome</strong> 
                  <input type="text" name="titolo" class="form" size="30" value="<%=titolo%>" />&nbsp;&nbsp;<strong>Codice</strong> 
                  <input type="text" name="codice" class="form" size="15" value="<%=codice%>" />
                  &nbsp;&nbsp;<input type="submit" name="cerca" value="cerca" class="form" /><%if FkProduttore>0 or titolo<>"" or codice<>"" then%>&nbsp;&nbsp;<a href="ges-prodotti-stampa.asp?ordine=<%=ordine%>&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>" target="_blank">[STAMPA]</a><%end if%></td>
              </tr>
			  </form>
			  <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
              <tr class="admin-intestazione" align="left"> 
                <td width="33%">&nbsp;<a href="ges-prodotti.asp?ordine=0&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>">Cod.</a>&nbsp;Titolo&nbsp;<a href="ges-prodotti.asp?ordine=1&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>">A/Z</a>&nbsp;<a href="ges-prodotti.asp?ordine=2&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>">Z/A</a></td>
                <td width="20%">Categoria</td>
				<td width="18%" align="center">Codice&nbsp;<a href="ges-prodotti.asp?ordine=3&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>">A/Z</a>&nbsp;<a href="ges-prodotti.asp?ordine=4&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>">Z/A</a></td>
				<td width="10%" align="center">Visualizz.</td>
                <td width="10%" align="center">Primo p.</td>
                <td width="9%" align="center">Elimina</td>
              </tr>
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
              <%
			  if nrs.recordcount>0 then	
			  	Do While Not nrs.EOF and rowCount < nrs.PageSize 
				Rowcount = rowCount + 1
				
				FkCategoria2=nrs("FkCategoria2")
				if FkCategoria2>0 then
					Set rs=Server.CreateObject("ADODB.Recordset")
					'sql = "Select * From Cat_Prod where pkid="&FkCat_Prod
					sql = "SELECT Categorie1.PkId as PkId_1, Categorie1.Titolo as Titolo_1, Categorie2.PkId as PkId_2, Categorie2.Titolo as Titolo_2 "
					sql = sql + "FROM Categorie1 INNER JOIN Categorie2 ON Categorie1.PkId = Categorie2.Fkcategoria1 "
					sql = sql + "WHERE Categorie2.PkId = "&FkCategoria2&""
					rs.Open sql, conn, 3, 3
					if rs.recordcount>0 then
						'titolocat=rs("Titolo_1")&"<br>"&rs("Titolo_2")
						titolocat=rs("Titolo_2")
					end if
					rs.close
				else
					titolocat="Nessuna categoria"
				end if
			  %>
              <tr align="left" class="admin-righe" <% if t = 1 then %>bgcolor="#CFCFCF"<% end if %>> 
                <td>&nbsp;<a href="sche-prodotti.asp?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>&p=<%=p%>"><font color="#CC0000"><%=nrs("pkid")%>.</font><%=Left(nrs("Titolo"),18)%><%if Len(nrs("Titolo"))>18 then%>...<%end if%></a></td>
                <td><%=titolocat%></td>
                <td align="center"><%=nrs("CodiceArticolo")%><%if Len(nrs("CodiceArticolo_Azienda"))>0 then%><br /><%=nrs("CodiceArticolo_Azienda")%><%end if%></td>
				<td align="center">
				<%if nrs("Offerta")=0 then%>Solo prod.<%end if%>
				<%if nrs("Offerta")=1 then%>Solo off.<%end if%>
				<%if nrs("Offerta")=2 then%>Prod./Off.<%end if%>
                <%if nrs("Offerta")=10 then%>Non visibile<%end if%>
                </td>
                <td align="center">
<%if nrs("PrimoPiano")=True then%>
                  Si <%else%>
                  No <%end if%> </td>
                <td align="center"><a href="sche-prodotti.asp?mode=1&pkid=<%=nrs("pkid")%>&C1=ON&ordine=<%=ordine%>&p=<%=p%>"><font color="#CC0000">X</font></a></td>
              </tr>
              <% if t = 1 then t = 0 else t = 1 %>
              <%
				nrs.movenext
			  	loop
			  %>
              <%else%>
              <tr> 
                <td colspan="6">Nessun record presente</td>
              </tr>
              <%end if%>
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
              <tr class="admin-intestazione" align="left"> 
                <td colspan="6">&nbsp; <% if nrs.recordcount > 0 then %>
                  Pag. <strong><%=p%></strong> di <%=nrs.PageCount%> Vai alla pagina&nbsp; <% if p > 5 then %>
                  [<a href="ges-prodotti.asp?p=<%=p-5%>&ordine=<%=ordine%>&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>">&lt;&lt; 
                  5 prec</a>] 
                  <% end if %> <% if p > 1 then %>
                  [<a href="ges-prodotti.asp?p=<%=p-1%>&ordine=<%=ordine%>&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>">&lt; 
                  prec</a>] 
                  <% end if %> <% for page = p to p+4 %> <a href="ges-prodotti.asp?p=<%=Page%>&ordine=<%=ordine%>&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>" class="testo"><%=page%></a> <% if page = nrs.PageCount then
		   		 		page = p+4
   		 				end if
	    				%> <% next %> <% if page-1 < nrs.PageCount then %>
                  [<a href="ges-prodotti.asp?p=<%=p+1%>&ordine=<%=ordine%>&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>">succ 
                  &gt;</a>] 
                  <% end if %> <% if nrs.PageCount-page > 5 then %>
                  [<a href="ges-prodotti.asp?p=<%=p+5%>&ordine=<%=ordine%>&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>">5 
                  succ &gt;&gt;</a>] 
                  <% end if%> &nbsp; &nbsp;[<a href="ges-prodotti.asp?p=<%=nrs.PageCount%>&ordine=<%=ordine%>&titolo=<%=titolo%>&codice=<%=codice%>&FkProduttore=<%=FkProduttore%>&Offerta=<%=Offerta%>">ultima 
                  pagina</a>] 
                  <% end if %> </td>
              </tr>
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
              <%nrs.close%>
            </table>
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