<!--#include file="strConn.asp"-->
<%
mode=request("mode")
if mode="" then mode=0
id=request("id")
tab=request("tab")
idfile=request("idfile")
%>
<%
if mode=0 then
	Set pps=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Immagini where pkid="&idfile
	pps.Open sql, conn, 3, 3
		foto=pps("file")
		titolo_file_it=pps("titolo")
	pps.close
end if
%>
<%
if mode=1 then
	pkid_fileUpload=request("pkid_fileUpload")
	titolo_file_it=request("titolo_file_it")
	Set pps=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Immagini where pkid="&pkid_fileUpload
	pps.Open sql, conn, 3, 3
		pps("titolo")=titolo_file_it
	pps.update
	pps.close
	
	response.Redirect("ins_file.asp?id="&id&"&tab="&tab&"")
end if
%>
<!--#include file="strClose.asp"-->
<html>
<head>
<title>:: Control Panel ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
</head>

<body style="border-style: none;">
<table width="545" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
      <table width="98%" border="0" cellspacing="0" cellpadding="0" height="100%">
        <tr> 
          <td width="2"><img src="immagini/spacer.gif" width="5" height="1"></td>
          <td width="540" align="left"> 
			<%if mode=0 then%>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" class="admin-righe">
              <tr> 
                <td height="5" colspan="3" align="center"><img src="immagini/spacer.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td colspan="3" align="left">
				<span class="testo">
				  Nome del File: <b><%=foto%></b>                  </span><br>                </td>
              </tr>
			  <form method="post" action="upload-file2.asp?mode=1&id=<%=id%>&tab=<%=tab%>">
			  <input type="hidden" name="pkid_fileUpload" value="<%=idfile%>">
			  <tr class="testo">
	  			<td height="20" align="left" colspan="3">
				Se vuoi, puoi modificare il Titolo all'immagine inserita:				</td>
			  </tr>
			  <tr>
	  			<td width="15%" height="20" align="left">
				<b>Titolo:</b>&nbsp;				</td>
			    <td height="20" colspan="2" align="left"><input type="text" name="titolo_file_it" value="<%=titolo_file_it%>" class="form" size="40"></td>
			  </tr>
			  <tr>
	  			<td height="25" align="left">&nbsp;				</td>
			    <td colspan="2" align="left"><input type="submit" name="invia" value="invia" class="form"></td>
			  </tr>
			  </form>
			  <tr>
	  			<td height="15" align="left" bgcolor="#EAEAEA" colspan="2">&nbsp;<a href="ins_file.asp?id=<%=id%>&tab=<%=tab%>">&raquo;ELENCO FOTO COLLEGATE</a></td>
			    <td width="48%" align="right" bgcolor="#EAEAEA"><a href="upload-file1.asp?id=<%=id%>&tab=<%=tab%>">&raquo;COLLEGA UN'ALTRA FOTO</a>&nbsp;</td>
			  </tr>
            </table>
		  <%end if%>
		  <%if mode=1 then%>
		  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="admin-righe">
              <tr> 
                <td height="5" colspan="2" align="center"><img src="immagini/space.gif" width="1" height="1"></td>
              </tr>
			  <tr> 
                <td height="30" colspan="2" align="center">Aggiornamento riuscito con successo</td>
              </tr>
			  <tr>
	  			<td height="15" align="left" bgcolor="#EAEAEA">&nbsp;<a href="ins_file.asp?id=<%=id%>&tab=<%=tab%>">&raquo;ELENCO FOTO COLLEGATE</a></td>
			    <td align="right" bgcolor="#EAEAEA"><a href="upload-file1.asp?id=<%=id%>&tab=<%=tab%>">&raquo;COLLEGA UN'ALTRA FOTO</a>&nbsp;</td>
			  </tr>
		  </table>
		  <%end if%>
		  </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
