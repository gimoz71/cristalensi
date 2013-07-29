<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
<%
ordine=request("ordine")
if ordine="" then ordine=4

if ordine=1 then ord="Prodotti.Titolo ASC"
if ordine=2 then ord="Prodotti.Titolo DESC"
if ordine=3 then ord="Sum(RigheOrdine.Quantita) ASC"
if ordine=4 then ord="Sum(RigheOrdine.Quantita) DESC"
if ordine=5 then ord="Categorie2.Titolo ASC"
if ordine=6 then ord="Categorie2.Titolo DESC"

			
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT RigheOrdine.FkProdotto, Prodotti.Titolo AS Titolo, Prodotti.CodiceArticolo, Produttori.Titolo  AS Produttore, Categorie2.Titolo AS Categoria, Prodotti.NomePagina, Sum(RigheOrdine.Quantita) AS SommaDiQuantita, Sum(RigheOrdine.TotaleRiga) AS SommaDiTotaleRiga "
sql = sql + "FROM RigheOrdine LEFT JOIN (Categorie2 RIGHT JOIN (Prodotti LEFT JOIN Produttori ON Prodotti.FkProduttore = Produttori.PkId) ON Categorie2.PkId = Prodotti.FkCategoria2) ON RigheOrdine.FkProdotto = Prodotti.PkId "
sql = sql + "GROUP BY RigheOrdine.FkProdotto, Prodotti.Titolo, Prodotti.CodiceArticolo, Produttori.Titolo, Categorie2.Titolo, Prodotti.NomePagina "
sql = sql + "ORDER BY "&ord&""
nrs.Open sql, conn, 1, 1
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
          <td width="220" class="menu-celle">Gestione Report</td>
          <td width="371" class="menu-celle" align="right"><a href="ges-report.asp">Elenco Report &raquo;</a>&nbsp;</td>
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
                <td colspan="5">&nbsp;</td>
              </tr>
              <tr class="admin-intestazione" align="left"> 
                <td width="38%">&nbsp;Titolo&nbsp;(Cod. - Prod.)&nbsp;<a href="ges-report2.asp?ordine=1&tipo=<%=tipo%>">A/Z</a>&nbsp;<a href="ges-report2.asp?ordine=2&tipo=<%=tipo%>">Z/A</a></td>
                <td width="26%">Categoria&nbsp;<a href="ges-report2.asp?ordine=5&tipo=<%=tipo%>">A/Z</a>&nbsp;<a href="ges-report2.asp?ordine=6&tipo=<%=tipo%>">Z/A</a></td>
                <td width="14%" align="center">Tot. ord.</td>
                <td colspan="2">Tot. art.&nbsp;<a href="ges-report2.asp?ordine=3&tipo=<%=tipo%>">0/1</a>&nbsp;<a href="ges-report2.asp?ordine=4&tipo=<%=tipo%>">1/0</a></td>
              </tr>
              <tr> 
                <td colspan="5">&nbsp;</td>
              </tr>
              <%
			  if nrs.recordcount>0 then	
			  	Do While Not nrs.EOF
				
				'visualizzazioni=nrs("visualizzazioni")
				'if visualizzazioni="" or IsNull(visualizzazioni) then visualizzazioni=0
			  %>
              <tr align="left" class="admin-righe" <% if t = 1 then %>bgcolor="#CFCFCF"<% end if %>> 
                <td height="30">&nbsp;<font color="#CC0000"><%=nrs("FkProdotto")%></font>.<%=nrs("Titolo")%><br />(<%=nrs("CodiceArticolo")%> - <%=nrs("Produttore")%>)</td>
                <td><%=nrs("Categoria")%></td>
                <td align="center"><%=FormatNumber(nrs("SommaDiTotaleRiga"),2)%></td>
                <td width="8%" align="center"><%=nrs("SommaDiQuantita")%></td>
                <td width="14%" align="center"><a href="sche-prodotti.asp?pkid=<%=nrs("FkProdotto")%>" target="_blank">Scheda</a><br /><a href="http://www.cristalensi.it/public/pagine/<%=nrs("NomePagina")%>" target="_blank">Sito</a></td>
              </tr>
              <% if t = 1 then t = 0 else t = 1 %>
              <%
				nrs.movenext
			  	loop
			  %>
              <%else%>
              <tr> 
                <td colspan="5">Nessuna visualizzazione presente</td>
              </tr>
              <%end if%>
              <tr> 
                <td colspan="5">&nbsp;</td>
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