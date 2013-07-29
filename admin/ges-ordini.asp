<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
<%
p=request("p")
if p="" then p=1
					
ordine=request("ordine")
if ordine="" then ordine=0
if ordine=0 then ord="PkId DESC"
if ordine=1 then ord="DataOrdine ASC"
if ordine=2 then ord="DataOrdine DESC"

			
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Ordini "
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
          <td width="267" class="menu-celle">Gestione ordini</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-ordini.asp">Elenco ordini &raquo;</a></td>
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
			  if nrs.recordcount>0 then	
			  	Do While Not nrs.EOF and rowCount < nrs.PageSize 
				Rowcount = rowCount + 1
				
				FkCliente=nrs("FkCliente")
				Set rs=Server.CreateObject("ADODB.Recordset")
				sql = "Select * From Clienti where pkid="&FkCliente
				rs.Open sql, conn, 3, 3
				if rs.recordcount>0 then
					Nominativo=rs("Nominativo")
					Nome=rs("Nome")
				else
					Nominativo="Non iscritto"
					Nome=""
				end if
				rs.close
			  %>
              <tr align="left" class="admin-righe" <% if t = 1 then %>bgcolor="#CFCFCF"<% end if %>> 
                <td>&nbsp;<a href="sche-ordini.asp?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>"><font color="#CC0000"><%=nrs("pkid")%></font></a></td>
                <td><%'if Nominativo<>"" then%><%=Nominativo%>&nbsp;<%=Nome%><%'else%><!--Non iscritto--><%'end if%></td>
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
              <tr class="admin-intestazione" align="left"> 
                <td colspan="6">&nbsp; <% if nrs.recordcount > 0 then %>
                  Pag. <strong><%=p%></strong> di <%=nrs.PageCount%> Vai alla pagina&nbsp; <% if p > 5 then %>
                  [<a href="ges-ordini.asp?p=<%=p-5%>&ordine=<%=ordine%>">&lt;&lt; 
                  5 prec</a>] 
                  <% end if %> <% if p > 1 then %>
                  [<a href="ges-ordini.asp?p=<%=p-1%>&ordine=<%=ordine%>">&lt; 
                  prec</a>] 
                  <% end if %> <% for page = p to p+4 %> <a href="ges-ordini.asp?p=<%=Page%>&ordine=<%=ordine%>" class="testo"><%=page%></a> <% if page = nrs.PageCount then
		   		 		page = p+4
   		 				end if
	    				%> <% next %> <% if page-1 < nrs.PageCount then %>
                  [<a href="ges-ordini.asp?p=<%=p+1%>&ordine=<%=ordine%>">succ 
                  &gt;</a>] 
                  <% end if %> <% if nrs.PageCount-page > 5 then %>
                  [<a href="ges-ordini.asp?p=<%=p+5%>&ordine=<%=ordine%>">5 
                  succ &gt;&gt;</a>] 
                  <% end if%> &nbsp; &nbsp;[<a href="ges-ordini.asp?p=<%=nrs.PageCount%>&ordine=<%=ordine%>">ultima 
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