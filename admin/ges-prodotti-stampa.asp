<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
<%
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
			
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Prodotti "
if titolo<>"" then
	ricerca = "WHERE Titolo LIKE '%"&titolo&"%' "
	if FkProduttore>0 then ricerca = ricerca + "AND FkProduttore="&FkProduttore&" "
end if
if codice<>"" then
	ricerca = "WHERE CodiceArticolo LIKE '%"&codice&"%' "
	if FkProduttore>0 then ricerca = ricerca + "AND FkProduttore="&FkProduttore&" "
end if
if titolo<>"" and codice<>"" then
	ricerca = "WHERE Titolo LIKE '%"&titolo&"%' AND CodiceArticolo LIKE '%"&codice&"%' "
	if FkProduttore>0 then ricerca = ricerca + "AND FkProduttore="&FkProduttore&" "
end if
if titolo="" and codice="" and FkProduttore>0 then
	ricerca = "WHERE FkProduttore="&FkProduttore&" "
end if
sql = sql + ricerca
sql = sql + "ORDER BY "&ord&""
nrs.Open sql, conn, 1, 1
%>
<html>
<head>
<title>Cristalensi Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
</head>

<body onLoad="print();">
<table width="750" border="0" align="left" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="top">
      <!--tab centrale-->
    <table width="98%"  border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td colspan="5">&nbsp;</td>
          </tr>
          <tr class="admin-righe" align="left" height="25"> 
            <td width="30%"><strong>&nbsp;Cod.&nbsp;Titolo&nbsp;</strong></td>
            <td width="20%"><strong>Codice</strong></td>
            <td width="33%"><strong>Categoria</strong></td>
            <td width="9%" align="right"><strong>P. Listino</strong></td>
            <td width="8%" align="right"><strong>P. Crist.</strong></td>
          </tr>
          <tr> 
            <td colspan="5"><hr style="padding:0px; margin:0px;" /></td>
          </tr>
          <tr> 
            <td colspan="5">&nbsp;</td>
          </tr>
          <%
          if nrs.recordcount>0 then	
            Do While Not nrs.EOF
            
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
          <tr align="left" class="admin-righe" height="20"> 
            <td>&nbsp;<font color="#CC0000"><%=nrs("pkid")%>.</font><%=nrs("Titolo")%></td>
            <td><%=nrs("CodiceArticolo")%></td>
            <td><%=titolocat%></td>
            <td align="right"><%=nrs("PrezzoListino")%></td>
            <td align="right"><%=nrs("PrezzoProdotto")%></td>
          </tr>
          <tr> 
            <td colspan="5" style="border-top-color:#000; border-top-style: solid; border-top-width:1px; height:1px; font-size:1px;">&nbsp;</td>
          </tr>
          <!--<tr> 
            <td colspan="5"><hr style="padding:0px; margin:0px;" /></td>
          </tr>-->
          <% if t = 1 then t = 0 else t = 1 %>
          <%
            nrs.movenext
            loop
          %>
          <%else%>
          <tr> 
            <td colspan="5">Nessun record presente</td>
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
</table>
</body>
</html>
<!--#include file="strClose.asp"-->
<!--#include file="chiusura.asp"-->