<%
err=request("err")
%>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
<!--
function closeWin()
{
		self.history.back();
}
//-->
</SCRIPT>
<html>
<head>
<title>:: Control Panel ::</title>
<link href="stile.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#FFFFFF">
<table cellpadding="0" cellspacing="0" border="0" align="center" class="admin-righe">
  <tr><td align="center" valign="middle">
	<table>
        <%if err=1 then%>
		<tr>
	      <td align="center"><b><br>
            <br>Attenzione!!!<br> <br>
            Il File scelto è già presente nel Database.<br>
            La preghiamo di rinominare il file e ripetere 
            l'operazione.<br>
            <br>
            Grazie.<br>
            <br></b>
          </td>
		</tr>
		<%else%>
		<tr>
	      <td align="center"><b><br>
            <br>Attenzione!!!<br> <br>
            Il File immesso ha dimensioni maggiori del massimo consentito quindi 
            non sarà inserito nel DataBase.<br>
            La preghiamo di realizzare il file con la giusta dimensione e ripetere 
            l'operazione.<br>
            <br>
            Grazie.<br>
            <br></b>
          </td>
		</tr>
		<%end if%>
		<tr>
          <td align="center"><br><br><a href="#" onclick="closeWin();return false">Torna 
            indietro</a></td>
        </tr>
	</table>
</td></tr>
</table>
</body>
</html>
