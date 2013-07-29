<!--#include file="strConn.asp"-->
<%
id=request("id")
tab=request("tab")

elim=request("elim")
if elim="" then elim=0
if elim=1 then
	idfile=request("idfile")
	Set pps=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Immagini where pkid="&idfile
	pps.Open sql, conn, 3, 3
	pps.delete
	pps.update
	pps.close
end if
%>
<html>
<head>
<title>Documento senza titolo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
<SCRIPT language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body style="border-style: none;">
<table width="548" height="100" border="0" cellpadding="0" cellspacing="0">
<tr><td valign="top" align="center">
	<table width="100%" align="center" cellpadding="0" cellspacing="0" border="0" class="admin-righe">
	
	<tr>
	  <td colspan="5" height="25">&nbsp;<strong>ELENCO FOTO COLLEGATE</strong></td>
	</tr>
	<tr>
	  <td height="25">&nbsp;FOTO</td>
	  <td>ZOOM</td>
	  <td height="25">&nbsp;TITOLO</td>
	  <td colspan="2" height="25">&nbsp;</td>
	</tr>
	<%
	Set pps=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Immagini where Record="&id&" and tabella='"&tab&"' Order by PkId ASC"
	pps.Open sql, conn, 1, 1
	if pps.recordcount>0 then
	Do while not pps.EOF 
	%>
	<tr><td width="23%" height="20">&nbsp;<a href="../public/<%=pps("file")%>" target="_blank"><%=pps("file")%></a></td>
	  <td width="12%" height="20">&nbsp;
	    <%if pps("zoom")<>"" then%><a href="../public/<%=pps("zoom")%>" target="_blank">SI</a><%else%>NO<%end if%></td>
	  <td width="34%" height="22" align="left">&nbsp;<%=pps("titolo")%></td>
	  <td width="20%" align="left"><a href="upload-file2.asp?id=<%=id%>&tab=<%=tab%>&idfile=<%=pps("pkid")%>">MODIFICA TESTO</a>&nbsp;</td>
	  <td width="11%" align="right"><a href="ins_file.asp?elim=1&id=<%=id%>&tab=<%=tab%>&idfile=<%=pps("pkid")%>">ELIMINA</a>&nbsp;</td>
	</tr>
	<tr>
	  <td colspan="5" height="3"><img src="immagini/spacer.gif" height="3"></td>
	</tr>
	<%
	pps.movenext
	loop
	else
	%>
	<tr>
	  <td colspan="5" height="30"><span>&nbsp;Nessuna foto collegata</span></td>
	</tr>
	<%
	end if
	pps.close
	%>
	<tr>
	  <td colspan="5" height="25" align="right" bgcolor="#E6E6E6"><a href="upload-file1.asp?id=<%=id%>&tab=<%=tab%>">PER INSERIRE UNA FOTO, CLICCA QUI</a>&nbsp;</td>
	</tr>
	</table>
</td></tr>
</table>
</body>
</html>
<!--#include file="strClose.asp"-->
