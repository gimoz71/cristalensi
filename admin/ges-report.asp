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
          <td width="236" class="menu-celle">Gestione Report e Visualizzazioni </td>
          <td width="355" class="menu-celle" align="right">&nbsp;</td>
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
                <td width="43%">&nbsp;Tipologia</td>
                <td>Data</td>
              </tr>
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
              <tr align="left" class="admin-righe" bgcolor="#CFCFCF"> 
                <td height="25">&nbsp;<a href="ges-report1.asp?tipo=2">Report Prodotti + visti</a></td>
                <td height="25">Dal 12/09/2011</td>
              </tr>
			   
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
              
              <tr align="left" class="admin-righe" bgcolor="#CFCFCF"> 
                <td height="25" colspan="2">&nbsp;<a href="ges-report2.asp?tipo=2">Report Prodotti + acquistati</a></td>
              </tr>
			   
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
              
              <tr align="left" class="admin-righe" bgcolor="#CFCFCF"> 
                <td height="25" colspan="2">&nbsp;<a href="ges-report3.asp">Report Ordini</a></td>
              </tr>
			   
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
              
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