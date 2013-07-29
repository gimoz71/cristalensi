<!--#include file="strConn.asp"-->
<%
mode=request("mode")
if mode="" then mode=0
if mode=1 then
	file_server=request("file_server")
end if
%>
<html>
<head>
<title>Documento senza titolo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
				<script language="Javascript">
				function reload() {
			 		self.opener.document.forms['server'].elements['file_server'].value = '<%=file_server%>';
			 		self.close();
			 	}
				</script>
</head>

<body <%if mode=1 then%> onLoad="reload();"<%end if%>>
<table width="380" height="250" border="0" cellpadding="0" cellspacing="0" >
<tr><td valign="top" align="center">
	<table width="100%" align="center" cellpadding="0" cellspacing="0" border="0" class="admin-righe">
	
	<tr>
	  <td colspan="3" height="25">&nbsp;&nbsp;<strong>ELENCO IMMAGINI PRESENTI SUL SERVER</strong></td>
	</tr>
	<tr>
	  <td height="25" colspan="2">&nbsp;&nbsp;scegliere un'immagine e cliccare su "invia"</td>
	  <td width="28%" height="25">&nbsp;</td>
	</tr>
	<form method="post" action="upload-file-server.asp?mode=1">
	<%
	Set objFso=Server.CreateObject("scripting.FileSystemObject")
	%>
	<%
	path="d:\inetpub\webs\cristalensiit\public\"
	'Set folder= objFso.getFolder( Server.MapPath("../public/") )
	Set folder= objFso.getFolder( path )
	Set files=folder.files
	
	for each file in files
	nome = file.name
	%>
	<%if Right(nome,3)="jpg" or Right(nome,3)="gif" then%>
	<tr><td width="9%" height="22" valign="middle">&nbsp;
	  <input type="radio" name="file_server" value="<%=nome%>"></td>
	  <td width="63%" valign="middle">&nbsp;<a href="../public/<%=nome%>" target="_blank"><%=nome%></a></td>
	  <td align="left"><a href="../public/<%=nome%>" target="_blank">VISUALIZZA</a>&nbsp;</td>
	</tr>
	<%end if%>
	
	<%next%>
	<tr>
	  <td colspan="3" height="5"><img src="immagini/spacer.gif" height="3"></td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
	  <td colspan="2" height="3"><input type="submit" name="invia" value="invia" class="form"></td>
	</tr>
	<tr>
	  <td colspan="3" height="5"><img src="immagini/spacer.gif" height="3"></td>
	</tr>
	</form>
	<tr>
	  <td colspan="3" height="25" align="right" bgcolor="#E6E6E6"><a href="upload-file1.asp?id=<%=id%>&tab=<%=tab%>">CHIUDI FINESTRA</a>&nbsp;</td>
	</tr>
	<tr>
	  <td colspan="3" height="5"><img src="immagini/spacer.gif" height="3"></td>
	</tr>
	</table>
</td></tr>
</table>
</body>
</html>
<!--#include file="strClose.asp"-->
