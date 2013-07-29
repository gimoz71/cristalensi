<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
<%
p=request("p")
if p="" then p=1
					
ordine=request("ordine")
if ordine="" then ordine=0
if ordine=0 then ord="PkId DESC"
if ordine=1 then ord="Data ASC"
if ordine=2 then ord="Data DESC"

			
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Commenti_Clienti "
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
          <td width="267" class="menu-celle">Gestione commenti</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-ordini.asp">Elenco commenti &raquo;</a>&nbsp;&nbsp;<a href="sche-commenti.asp">Nuovo commento &raquo;</a></td>
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
              <tr class="admin-intestazione" align="left"> 
                <td width="27%">&nbsp;<a href="ges-commenti.asp?ordine=0">Cod.Cliente</a></td>
                <td width="28%">Testo</td>
				<td width="10%" align="center">Pubblicato</td>
				<td width="11%" align="center">Risposta</td>
                <td width="16%" align="center">Data&nbsp;<a href="ges-commenti.asp?ordine=1">0/1</a>&nbsp;<a href="ges-commenti.asp?ordine=2">1/0</a></td>
                <td width="8%" align="center">Elimina</td>
              </tr>
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
              <%
			  if nrs.recordcount>0 then	
			  	Do While Not nrs.EOF and rowCount < nrs.PageSize 
				Rowcount = rowCount + 1
				
				FkCliente=nrs("FkIscritto")
				Set rs=Server.CreateObject("ADODB.Recordset")
				sql = "Select * From Clienti where pkid="&FkCliente
				rs.Open sql, conn, 3, 3
				if rs.recordcount>0 then
					Nominativo=rs("Nominativo")
					Nome=rs("Nome")
				else
					Nominativo="Non iscritto"
				end if
				rs.close
			  %>
              <tr align="left" class="admin-righe" <% if t = 1 then %>bgcolor="#CFCFCF"<% end if %>> 
                <td>&nbsp;<a href="sche-commenti.asp?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>"><font color="#CC0000"><%=nrs("pkid")%>.<%=Nominativo%>&nbsp;<%=Nome%></font></a></td>
                <td><%=Left(nrs("Testo"), 20)%>...</td>
                <td align="center"><%if nrs("Pubblicato")=True then%>Si<%else%>No<%end if%></td>
				<td align="center"><%if nrs("Risposta")=True then%>Si<%else%>No<%end if%></td>
                <td align="center">
					<%=Left(nrs("Data"), 10)%>
                </td>
                <td align="center"><a href="sche-commenti.asp?mode=1&pkid=<%=nrs("pkid")%>&C1=ON&ordine=<%=ordine%>&p=<%=p%>"><font color="#CC0000">X</font></a></td>
               </tr>
              <% if t = 1 then t = 0 else t = 1 %>
              <%
				nrs.movenext
			  	loop
			  %>
              <%else%>
              <tr> 
                <td colspan="6">Nessun commento presente</td>
              </tr>
              <%end if%>
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
              <tr class="admin-intestazione" align="left"> 
                <td colspan="6">&nbsp; <% if nrs.recordcount > 0 then %>
                  Pag. <strong><%=p%></strong> di <%=nrs.PageCount%> Vai alla pagina&nbsp; <% if p > 5 then %>
                  [<a href="ges-commenti.asp?p=<%=p-5%>&ordine=<%=ordine%>">&lt;&lt; 
                  5 prec</a>] 
                  <% end if %> <% if p > 1 then %>
                  [<a href="ges-commenti.asp?p=<%=p-1%>&ordine=<%=ordine%>">&lt; 
                  prec</a>] 
                  <% end if %> <% for page = p to p+4 %> <a href="ges-commenti.asp?p=<%=Page%>&ordine=<%=ordine%>" class="testo"><%=page%></a> <% if page = nrs.PageCount then
		   		 		page = p+4
   		 				end if
	    				%> <% next %> <% if page-1 < nrs.PageCount then %>
                  [<a href="ges-commenti.asp?p=<%=p+1%>&ordine=<%=ordine%>">succ 
                  &gt;</a>] 
                  <% end if %> <% if nrs.PageCount-page > 5 then %>
                  [<a href="ges-commenti.asp?p=<%=p+5%>&ordine=<%=ordine%>">5 
                  succ &gt;&gt;</a>] 
                  <% end if%> &nbsp; &nbsp;[<a href="ges-commenti.asp?p=<%=nrs.PageCount%>&ordine=<%=ordine%>">ultima 
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