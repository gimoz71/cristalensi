<%
mode=request("mode")
if mode="" then mode=0
id=request("id")
tab=request("tab")
%>
<!--#include file="strConn.asp"-->
<%
if mode=2 then
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
<%
if mode=1 or mode=4 then
%>
<!-- #include file="Upload2-file.asp" -->
<%	
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
	
	
	if (Request.TotalBytes > 0) then
	sz=500
	if objUpload.Files.Count > 0 then
		if objUpload.Files.Item(file).Size > (cInt(sz)*1000) then
		Response.Redirect("file_grande.asp?err=2")
		Response.End
		end if
	end if
	end if
end if
%>
<%if mode=3 or mode=5 then%>
<!--#include file="strConn.asp"-->
<%
	if mode=3 then
		Set pps=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Immagini"
		pps.Open sql, conn, 3, 3
		pps.addnew
			pps("file")=request("file_server")
			pps("Record")=id
			pps("tabella")=tab
		pps.update
		pps.close
							
		Set pps=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Immagini order by pkid desc"
		pps.Open sql, conn, 1, 1
			pkid_fileUpload=pps("pkid")
		pps.close
	end if
	if mode=5 then
		pkid_fileUpload=request("pkid_fileUpload")
		Set pps=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Immagini where pkid="&pkid_fileUpload
		pps.Open sql, conn, 3, 3
			file_prec=pps("file")
			pps("zoom")=request("file_server")
		pps.update
		pps.close
	end if
%>
<!--#include file="strClose.asp"-->
<%end if%>
<html>
<head>
<title>:: Control Panel ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body style="border-style: none;">
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
      <table width="98%" border="0" cellspacing="0" cellpadding="0" height="100%" class="admin-righe">
        <tr> 
          <td width="2"><img src="immagini/spacer.gif" width="5" height="1"></td>
          <td width="530" align="left"> 
            <%if mode=0 then%>
			
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="admin-righe">
              <form enctype="multipart/form-data" action="upload-file1.asp?mode=1&id=<%=id%>&tab=<%=tab%>" method="post" >
			  <tr> 
                  <td colspan="2" align="left">Per inserire un File che risiede sul tuo PC (il seguente file sarà visualizzato sul sito)</td>
              </tr>
			  <tr> 
                  <td colspan="2" align="left">Il File da inserire pu&ograve; raggiungere massimo 500KB con misure w: 120px e h: 90px, grazie </td>
              </tr>
              <tr> 
                <td height="5" colspan="2" align="center"><img src="immagini/spacer.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                  <td width="19%" height="25" align="left">Indirizzo del File:</td>
                  <td width="81%" align="left"><input type="file" name="file" size="40" class="form" >&nbsp;&nbsp;<input name="invia" type="submit" value="invia" class="form"></td>
              </tr>
			  <tr> 
                <td height="5" colspan="2" align="center"><img src="immagini/spacer.gif" width="1" height="1"></td>
              </tr>
			  </form>
			  <form name="server" action="upload-file1.asp?mode=3&id=<%=id%>&tab=<%=tab%>" method="post" >
			  <tr> 
                  <td colspan="2" align="left">Per inserire un File che sta sul sito </td>
              </tr>
              <tr> 
                <td height="5" colspan="2" align="center"><img src="immagini/spacer.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                  <td height="25" align="left">File:</td>
                  <td align="left"><input type="text" name="file_server" size="40" class="form" > <a href="#" onClick="MM_openBrWindow('upload-file-server.asp','','width=400,height=250,scrollbars=yes')">Elenco file</a>&nbsp;&nbsp;<input name="invia" type="submit" value="invia" class="form"></td>
              </tr>
			  <tr> 
                <td height="5" colspan="2" align="center"><img src="immagini/spacer.gif" width="1" height="1"></td>
              </tr>
			  </form>
			  <tr>
	  			<td height="15" colspan="2" align="right" bgcolor="#EAEAEA"><a href="ins_file.asp?id=<%=id%>&tab=<%=tab%>">&raquo;ELENCO FOTO COLLEGATE</a>&nbsp;</td>
			  </tr>
            </table>
            
			<%end if%>
			<%if mode=1 or  mode=3 or mode=4 or  mode=5 then%>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" class="admin-righe">
              <tr> 
                <td height="5" colspan="3" align="center"><img src="immagini/spacer.gif" width="1" height="1"></td>
              </tr>
			  <%if mode=1 then%>
              <tr> 
                <td colspan="3" align="left">
				<span>
				
                  <%'Dim objUpload
					'Dim lngLoop

				If Request.TotalBytes > 0 Then
					'Set objUpload = New clsUpload
					%>
                  Operazione Riuscita con successo ...<br>
				  <b><%'= objUpload.Files.Count %></b> 
                  <%
					For lngLoop = 0 to objUpload.Files.Count - 1
						'If accessing this page annonymously,
						'the internet guest account must have
						'write permission to the path below.
						mypath = cstr(Server.MapPath ("upload-foto"))
						objUpload.Files.Item(lngLoop).Save mypath
						%>
                  <%pathfile=objUpload.Files.Item(lngLoop).FileName%>
				  <%
				  	if pathfile<>"" then
				  %>
				  <!--#include file="strConn.asp"-->
				  <%
						Set pps=Server.CreateObject("ADODB.Recordset")
						sql = "Select * From Immagini"
						pps.Open sql, conn, 3, 3
						pps.addnew
							pps("file")=pathfile
							pps("Record")=id
							pps("tabella")=tab
						pps.update
						pps.close
						
						Set pps=Server.CreateObject("ADODB.Recordset")
						sql = "Select * From Immagini order by pkid desc"
						pps.Open sql, conn, 1, 1
							pkid_fileUpload=pps("pkid")
						pps.close
				  %>
				  <!--#include file="strClose.asp"-->
				  <%	
					end if
				  %>
				  File inserito: <b><%=pathfile%></b>                  </span><br><br>
                  <%
						'file = objUpload.Files.Item(lngLoop).FileName
					Next
				End If
				%>                </td>
              </tr>
			  <%end if%>
			  <%if mode=3 then%>
			  <tr> 
                <td colspan="3" align="left">
				Operazione Riuscita con successo ...<br>Nome del File: <b><%=request("file_server")%></b><br><br>
			  	</td>
			  </tr>
			  <%end if%>
			  
			  <%if mode=1 or mode=3 then%>
			  <!--zoom-->
			  <form enctype="multipart/form-data" action="upload-file1.asp?mode=4&id=<%=id%>&tab=<%=tab%>&pkid_fileUpload=<%=pkid_fileUpload%>" method="post" >
			  <tr> 
                  <td colspan="3" align="left"><b>Immagine Zoom</b> (è l'immagine che si ottiene cliccando su una piccola)</td>
              </tr>
			  <tr> 
                  <td colspan="3" height="20" align="left">Per inserire un File che risiede sul tuo PC (massimo 500KB)</td>
              </tr>
              <tr> 
                  <td width="19%" height="20" align="left">Indirizzo del File:</td>
                  <td width="81%" align="left" colspan="2"><input type="file" name="file" size="40" class="form" >&nbsp;&nbsp;<input name="invia" type="submit" value="invia" class="form"></td>
              </tr>
			  <tr> 
                <td height="5" colspan="3" align="center"><img src="immagini/spacer.gif" width="1" height="1"></td>
              </tr>
			  </form>
			  <form name="server" action="upload-file1.asp?mode=5&id=<%=id%>&tab=<%=tab%>&pkid_fileUpload=<%=pkid_fileUpload%>" method="post" >
			  <tr> 
                  <td colspan="3" height="20" align="left">Per inserire un File che sta sul sito </td>
              </tr>
              <tr> 
                  <td height="20" align="left">File:</td>
                  <td align="left" colspan="3"><input type="text" name="file_server" size="40" class="form" > <a href="#" onClick="MM_openBrWindow('upload-file-server.asp','','width=400,height=250,scrollbars=yes')">Elenco file</a>&nbsp;&nbsp;<input name="invia" type="submit" value="invia" class="form"></td>
              </tr>
			  <tr> 
                <td height="5" colspan="3" align="center"><hr></td>
              </tr>
			  </form>
			  
			  <!--fine zoom-->
			  <%end if%>
			  <%if mode=4 or mode=5 then%>
			  
			  <%if mode=4 then%>
              <tr> 
                <td colspan="3" align="left">
				<span>
				
                  <%'Dim objUpload
					'Dim lngLoop

				If Request.TotalBytes > 0 Then
					'Set objUpload = New clsUpload
					%>
                  Operazione Riuscita con successo ...<br>
				  <b><%'= objUpload.Files.Count %></b> 
                  <%
					For lngLoop = 0 to objUpload.Files.Count - 1
						'If accessing this page annonymously,
						'the internet guest account must have
						'write permission to the path below.
						mypath = cstr(Server.MapPath ("upload-foto"))
						objUpload.Files.Item(lngLoop).Save mypath
						%>
                  <%pathfile=objUpload.Files.Item(lngLoop).FileName%>
				  <%
				  	if pathfile<>"" then
				  %>
				  <!--#include file="strConn.asp"-->
				  <%
						Set pps=Server.CreateObject("ADODB.Recordset")
						sql = "Select * From Immagini order by pkid desc"
						pps.Open sql, conn, 1, 1
							pkid_fileUpload=pps("pkid")
						pps.close
						
						Set pps=Server.CreateObject("ADODB.Recordset")
						sql = "Select * From Immagini where pkid="&pkid_fileUpload
						pps.Open sql, conn, 3, 3
							file_prec=pps("file")
							pps("zoom")=pathfile
						pps.update
						pps.close
				  %>
				  <!--#include file="strClose.asp"-->
				  <%	
					end if
				  %>
				  Nome del File: <b><%=pathfile%></b>                  </span><br><br>
                  <%
						'file = objUpload.Files.Item(lngLoop).FileName
					Next
				End If
				%>                </td>
              </tr>
			  <%end if%>
			  <%if mode=5 then%>
			  <tr> 
                <td colspan="3" align="left">
				Operazione Riuscita con successo ...<br>File inserito: <b><%=request("file_server")%></b><br>
				File inserito precedentemente: <b><%=file_prec%></b><br><br>
			  	</td>
			  </tr>
			  <%end if%>
			  
			  <%end if%>
			  <form method="post" action="upload-file1.asp?mode=2&id=<%=id%>&tab=<%=tab%>">
			  <input type="hidden" name="pkid_fileUpload" value="<%=pkid_fileUpload%>">
			  <tr>
	  			<td height="20" align="left" colspan="3">
				Se vuoi, puoi aggiungere un Titolo/Commento all'immagine/i inserita/e:				</td>
			  </tr>
			  <tr>
	  			<td width="15%" height="20" align="left">
				<b>Titolo:</b>&nbsp;				</td>
			    <td height="20" colspan="2" align="left"><input type="text" name="titolo_file_it" class="form" size="40">&nbsp;<input type="submit" name="invia" value="invia" class="form"></td>
			  </tr>
			  <tr> 
                <td height="5" colspan="3" align="center"><img src="immagini/spacer.gif" width="1" height="1"></td>
              </tr>
			  </form>
			  <tr>
	  			<td height="15" colspan="2" align="left" bgcolor="#EAEAEA">&nbsp;<a href="ins_file.asp?id=<%=id%>&tab=<%=tab%>">&raquo;ELENCO FOTO COLLEGATE</a>&nbsp;</td>
			    <td width="46%" align="right" bgcolor="#EAEAEA"><a href="upload-file1.asp?id=<%=id%>&tab=<%=tab%>">&raquo;COLLEGA UN'ALTRA FOTO</a>&nbsp;</td>
			  </tr>
            </table>
		  <%end if%>
		  <%if mode=2 then%>
		  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="admin-righe">
              <tr> 
                <td height="5" colspan="2" align="center"><img src="immagini/spacer.gif" width="1" height="1"></td>
              </tr>
			  <tr> 
                <td height="30" colspan="2" align="center">Aggiornamento riuscito con successo</td>
              </tr>
			  <tr>
	  			<td height="15" align="left" bgcolor="#EAEAEA">&nbsp;<a href="ins_file.asp?id=<%=id%>&tab=<%=tab%>">&raquo;ELENCO FOTO COLLEGATE</a>&nbsp;</td>
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
