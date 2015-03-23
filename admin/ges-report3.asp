<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
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
              <%for anno=2010 to 2015%>
              <tr class="admin-intestazione" align="left"> 
                <td width="27%"><strong>&nbsp;ANNO&nbsp;<%=anno%></strong></td>
                <td width="15%" align="right">n. ordini</td>
                <td width="20%" align="right">totale&nbsp;carrello&nbsp;</td>
                <td width="20%" align="right">carrello medio</td>
                <td width="18%" align="right">totale&nbsp;ordine&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="5">&nbsp;</td>
              </tr>
              <%
				mm=1
				for mm=1 to 12
				
				if mm=1 then mese="GENNAIO"
				if mm=2 then mese="FEBBRAIO"
				if mm=3 then mese="MARZO"
				if mm=4 then mese="APRILE"
				if mm=5 then mese="MAGGIO"
				if mm=6 then mese="GIUGNO"
				if mm=7 then mese="LUGLIO"
				if mm=8 then mese="AGOSTO"
				if mm=9 then mese="SETTEMBRE"
				if mm=10 then mese="OTTOBRE"
				if mm=11 then mese="NOVEMBRE"
				if mm=12 then mese="DICEMBRE"
				
				if mm=1 then fine=31
				if mm=2 then fine=28
				if mm=2 and (anno=2012 or anno=2016 or anno=2020) then fine=29
				if mm=3 then fine=31
				if mm=4 then fine=30
				if mm=5 then fine=31
				if mm=6 then fine=30
				if mm=7 then fine=31
				if mm=8 then fine=31
				if mm=9 then fine=30
				if mm=10 then fine=31
				if mm=11 then fine=30
				if mm=12 then fine=31
				
				Set nrs=Server.CreateObject("ADODB.Recordset")
				sql = "SELECT Sum([Ordini.TotaleCarrello]) AS totale_carrello, Sum([Ordini.TotaleGenerale]) AS totale_generale, Count(*) AS n_ordini "
				sql = sql + "FROM Ordini "
				sql = sql + "WHERE (((Ordini.DataOrdine)>=#"&mm&"/1/"&anno&" 00:00:00# And (Ordini.DataOrdine)<=#"&mm&"/"&fine&"/"&anno&" 23:59:59#) AND ((Ordini.Stato)=7 Or (Ordini.Stato)=8))"
				nrs.Open sql, conn, 1, 1
				
			  %>
              <tr align="left" class="admin-righe"> 
                <td height="20">&nbsp;<%=mese%></td>
                <td align="right"><%=nrs("n_ordini")%></td>
                <td align="right"><%=FormatNumber(nrs("totale_carrello"),2)%></td>
                <td align="right"><%=FormatNumber((nrs("totale_carrello")/nrs("n_ordini")),2)%></td>
                <td align="right"><%=FormatNumber(nrs("totale_generale"),2)%></td>
              </tr>
              <%nrs.close%>
              <%next%>
              <tr> 
                <td colspan="5">&nbsp;</td>
              </tr>
              <%
			  	'Set trs=Server.CreateObject("ADODB.Recordset")
				'sql = "SELECT Sum([Ordini.TotaleCarrello]) AS totale_carrello, Sum([Ordini.TotaleGenerale]) AS totale_generale, Count(*) AS n_ordini "
				'sql = sql + "FROM Ordini "
				'sql = sql + "WHERE (((Ordini.DataOrdine)>=#1/1/"&anno&" 00:00:00# And (Ordini.DataOrdine)<=#12/31/"&anno&" 23:59:59#) AND ((Ordini.Stato)=7 Or (Ordini.Stato)=8))"
				'trs.Open sql, conn, 1, 1
			  %>
              <!--<tr align="left" class="admin-righe"> 
                <td height="20">&nbsp;Anno <%'=anno%></td>
                <td align="right"><%'=trs("n_ordini")%></td>
                <td align="right"><%'=FormatNumber(trs("totale_carrello"),2)%></td>
                <td align="right"><%'=FormatNumber((trs("totale_carrello")/trs("n_ordini")),2)%></td>
                <td align="right"><%'=FormatNumber(trs("totale_generale"),2)%></td>
              </tr>
              <%'trs.close%>
              <tr> 
                <td colspan="5">&nbsp;</td>
              </tr>-->
              <%next%>
              
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