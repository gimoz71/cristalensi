<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
<%
ordine=request("ordine")
if ordine="" then ordine=0
if ordine=0 then ord="pkid DESC"
if ordine=1 then ord="Nome ASC"
if ordine=2 then ord="Nome DESC"

			
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM CostiTrasporto "
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
          <td width="267" class="menu-celle">Gestione Costi Trasporto</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-spedizioni.asp">Elenco Costi &raquo;</a>&nbsp;&nbsp;<a href="sche-spedizioni.asp">Nuovo Costo &raquo;</a></td>
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
                <td>&nbsp;<a href="ges-spedizioni.asp?ordine=0">Cod.</a>&nbsp;Costo&nbsp;<a href="ges-spedizioni.asp?ordine=1">A/Z</a>&nbsp;<a href="ges-spedizioni.asp?ordine=2">Z/A</a></td>
                <td width="8%" align="center">Elimina</td>
              </tr>
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
              <%
			  if nrs.recordcount>0 then	
			  	Do While Not nrs.EOF
			  %>
              <tr align="left" class="admin-righe" <% if t = 1 then %>bgcolor="#CFCFCF"<% end if %>> 
                <td>&nbsp;<a href="sche-spedizioni.asp?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>"><font color="#CC0000"><%=nrs("pkid")%>.</font><%=nrs("nome")%></a></td>
                <td align="center"><a href="sche-spedizioni.asp?mode=1&pkid=<%=nrs("pkid")%>&C1=ON&ordine=<%=ordine%>&p=<%=p%>"><font color="#CC0000">X</font></a></td>
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