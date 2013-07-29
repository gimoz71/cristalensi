
<%
mode=request("mode")
if mode="" then mode=0
%>
<!-- #include file="Upload2-file.asp" -->
<%	
if mode=1 then	
	Dim objUpload, objFSO, lngLoop
	Set objUpload = New clsUpload
	img=objUpload.Files.Item(lngLoop).FileName
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	'If objFSO.FileExists(Server.MapPath("../public/"&img&"")) Then
	If objFSO.FileExists("d:\inetpub\webs\cristalensiit\public\"&img&"") Then
		Response.Redirect("file_grande.asp?err=1")
		Response.End
	End if
	Set objFSO = Nothing
	
	
	'if (Request.TotalBytes > 0) then
'	sz=500
'	if objUpload.Files.Count > 0 then
'		if objUpload.Files.Item(file).Size > (cInt(sz)*1000) then
'		Response.Redirect("file_grande.asp?err=2")
'		Response.End
'		end if
'	end if
'	end if
end if
%>
<html>
<head>
<title>Pannello di controllo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#E0E4DA" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="300" border="0" cellspacing="0" cellpadding="0" bgcolor="#E0E4DA">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
        <tr> 
          <td colspan="2" height="10"><img src="immagini/img_piccola.gif" width="1" height="1">
          </td>
        </tr>
        <tr> 
          <td width="7"><img src="immagini/img_piccola.gif" width="1" height="1"></td>
          <td width="300"> 
            <form enctype="multipart/form-data" action="upload-file.asp?mode=1" method="post" >
            <table width="300" border="0" cellspacing="0" cellpadding="0" class="admin-righe">
              <tr> 
                  <td width="468" align="center" class="testograssetto">File da inserire nel Database 
                    (massimo 500KB)</td>
              </tr>
              <tr> 
                <td height="10" width="470" align="center"><img src="immagini/img_piccola.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                  <td height="15" width="468" align="center" class="testo">Indirizzo del File: 
                    <input type="file" name="file" size="35" class="form" >
                  </td>
              </tr>
			  <tr align="center"> 
                <td height="2"> 
                  <img src="immagini/img_piccola.gif" width="1" height="1">
                </td>
              </tr>
              <tr align="center"> 
                <td align="center" >
				<input name="invia" type="submit" value="invia" class="form">
                  </td>
              </tr>
            </table>
            </form>
            <table width="300" border="0" cellspacing="0" cellpadding="0" class="admin-righe">
              <tr> 
                <td height="5" width="470" align="center"><img src="immagini/img_piccola.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td align="center" class="testo"> File Uploaded: 
                  <%
				'Dim lngLoop

				If Request.TotalBytes > 0 Then
				%>
                  <b><%= objUpload.Files.Count %></b> 
                <%
					For lngLoop = 0 to objUpload.Files.Count - 1
						mypath = cstr(Server.MapPath ("upload-foto1"))
						objUpload.Files.Item(lngLoop).Save mypath
				%>
                  <br>
                  Nome File: <b><%= objUpload.Files.Item(lngLoop).FileName %></b><br> 
                  <br>
                  Operazione Riuscita con successo ...<br>
                  Cliccare "Chiudi la Pagina", grazie. 
                <%
						file = objUpload.Files.Item(lngLoop).FileName
					Next
				End If
				%>
                  <script language="Javascript">
				function reload() {
			 		self.opener.document.forms['newsform'].elements['Allegato'].value = '<%=file%>';
			 		self.close();
			 	}
				</script> <br> 
                  <p align="center"><a href="#" onClick="reload()" class="testo">Chiudi la Pagina</a></p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
