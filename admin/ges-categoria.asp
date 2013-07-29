<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
<%
p=request("p")
if p="" then p=1
					
ordine=request("ordine")
if ordine="" then ordine=0
if ordine=0 then ord="pkid DESC"
if ordine=1 then ord="Titolo ASC"
if ordine=2 then ord="Titolo DESC"

			
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Cat_Prod "
sql = sql + "ORDER BY "&ord&""
nrs.Open sql, conn, 1, 1
					
nrs.PageSize = 20
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
          <td width="267" class="menu-celle">Gestione Categorie</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-categoria.asp">Elenco Categorie &raquo;</a>&nbsp;&nbsp;<a href="sche-categoria.asp">Nuova categria &raquo;</a></td>
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
                <td colspan="2">&nbsp;</td>
              </tr>
              <tr class="admin-intestazione" align="left"> 
                <td>&nbsp;<a href="ges-categoria.asp?ordine=0">Cod.</a>&nbsp;Categoria&nbsp;<a href="ges-categoria.asp?ordine=1">A/Z</a>&nbsp;<a href="ges-categoria.asp?ordine=2">Z/A</a></td>
                <td width="8%" align="center">Elimina</td>
              </tr>
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
              <%
			  if nrs.recordcount>0 then	
			  	Do While Not nrs.EOF and rowCount < nrs.PageSize 
				Rowcount = rowCount + 1
			  %>
              <tr align="left" class="admin-righe" <% if t = 1 then %>bgcolor="#CFCFCF"<% end if %>> 
                <td>&nbsp;<a href="sche-categoria.asp?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>"><font color="#CC0000"><%=nrs("pkid")%>.</font><%=nrs("titolo")%></a></td>
                <td align="center"><a href="sche-categoria.asp?mode=1&pkid=<%=nrs("pkid")%>&C1=ON&ordine=<%=ordine%>&p=<%=p%>"><font color="#CC0000">X</font></a></td>
               </tr>
              <% if t = 1 then t = 0 else t = 1 %>
              <%
				nrs.movenext
			  	loop
			  %>
              <%else%>
              <tr> 
                <td colspan="2">Nessuna categoria presente</td>
              </tr>
              <%end if%>
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
              <tr class="admin-intestazione" align="left"> 
                <td colspan="2">&nbsp; <% if nrs.recordcount > 0 then %>
                  Pag. <strong><%=p%></strong> di <%=nrs.PageCount%> Vai alla pagina&nbsp; <% if p > 5 then %>
                  [<a href="ges-categoria.asp?p=<%=p-5%>&ordine=<%=ordine%>">&lt;&lt; 
                  5 prec</a>] 
                  <% end if %> <% if p > 1 then %>
                  [<a href="ges-categoria.asp?p=<%=p-1%>&ordine=<%=ordine%>">&lt; 
                  prec</a>] 
                  <% end if %> <% for page = p to p+4 %> <a href="ges-categoria.asp?p=<%=Page%>&ordine=<%=ordine%>" class="testo"><%=page%></a> <% if page = nrs.PageCount then
		   		 		page = p+4
   		 				end if
	    				%> <% next %> <% if page-1 < nrs.PageCount then %>
                  [<a href="ges-categoria.asp?p=<%=p+1%>&ordine=<%=ordine%>">succ 
                  &gt;</a>] 
                  <% end if %> <% if nrs.PageCount-page > 5 then %>
                  [<a href="ges-categoria.asp?p=<%=p+5%>&ordine=<%=ordine%>">5 
                  succ &gt;&gt;</a>] 
                  <% end if%> &nbsp; &nbsp;[<a href="ges-categoria.asp?p=<%=nrs.PageCount%>&ordine=<%=ordine%>">ultima 
                  pagina</a>] <% end if %> </td>
              </tr>
              <tr> 
                <td colspan="2">&nbsp;</td>
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