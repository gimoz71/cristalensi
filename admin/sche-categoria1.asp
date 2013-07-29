<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="session.asp"-->
<!--#include file="strConn.asp"-->
<%
	Function NoHTML(Stringa)
		Set RegEx = New RegExp
		RegEx.Pattern = "<[^>]*>"
		RegEx.Global = True
		RegEx.IgnoreCase = True
		NoHTML = RegEx.Replace(Stringa, "")
	End Function

	Function ConvertiTitoloInNomeScript(Titolo, IDArticolo)
		Risultato = Titolo
		Risultato = NoHTML(Risultato)
		Risultato = LCase(Risultato)
		Risultato = Replace(Risultato, " ", "-")
		Risultato = Replace(Risultato, "\", "-")
		Risultato = Replace(Risultato, "/", "-")
		Risultato = Replace(Risultato, ":", "-")
		Risultato = Replace(Risultato, "*", "-")
		Risultato = Replace(Risultato, "?", "-")
		Risultato = Replace(Risultato, "<", "-")
		Risultato = Replace(Risultato, ">", "-")
		Risultato = Replace(Risultato, "|", "-")
		Risultato = Replace(Risultato, """", "")
		Risultato = Replace(Risultato, "'", "-")
		Risultato = IDArticolo & "-" & Risultato & ".asp"
		ConvertiTitoloInNomeScript = Risultato
	End Function
	
	Function ConvertiTitoloInNomeScript_en(Titolo_en, IDArticolo)
		Risultato=""
		Risultato = Titolo_en
		Risultato = NoHTML(Risultato)
		Risultato = LCase(Risultato)
		Risultato = Replace(Risultato, " ", "-")
		Risultato = Replace(Risultato, "\", "-")
		Risultato = Replace(Risultato, "/", "-")
		Risultato = Replace(Risultato, ":", "-")
		Risultato = Replace(Risultato, "*", "-")
		Risultato = Replace(Risultato, "?", "-")
		Risultato = Replace(Risultato, "<", "-")
		Risultato = Replace(Risultato, ">", "-")
		Risultato = Replace(Risultato, "|", "-")
		Risultato = Replace(Risultato, """", "")
		Risultato = Replace(Risultato, "'", "-")
		Risultato = IDArticolo & "-" & Risultato & ".asp"
		ConvertiTitoloInNomeScript_en = Risultato
	End Function
%>
<%
	pkid = request("pkid")
	if pkid = "" then pkid = 0
	
	p = request("p")
	if p = "" then p = 1
	ordine = request("ordine")
	if ordine = "" then ordine = 0
	
	mode = request("mode")
	if mode = "" then mode = 0

	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Categorie1"
	if pkid > 0 then sql = "Select * From Categorie1 where pkid="&pkid
	rs.Open sql, conn, 3, 3
	
	if mode = 1 then
		if pkid = 0 then rs.addnew
		
		rs("posizione")=request("posizione")
		Titolo=request("Titolo")
		rs("titolo")=Titolo
		rs("descrizione")=request("descrizione")
		
		rs("testo1")=request("testo1")
		rs("testo2")=request("testo2")
		
		NomePagina_Old=request("NomePagina")
		
		Titolo_en=request("Titolo_en")
		rs("titolo_en")=Titolo_en
		rs("descrizione_en")=request("descrizione_en")
		
		rs("testo1_en")=request("testo1_en")
		rs("testo2_en")=request("testo2_en")
		
		NomePagina_Old_en=request("NomePagina_en")
		
		if pkid>0 then
			rs("NomePagina")=ConvertiTitoloInNomeScript(Titolo, PkId)
			rs("NomePagina_en")=ConvertiTitoloInNomeScript_en(Titolo_en, PkId)
		end if
		
		if request("C1") = "ON" then
			
			NomePaginaDaEliminare=rs("NomePagina")
			Set FSO = CreateObject("Scripting.FileSystemObject")
			If FSO.FileExists(Server.MapPath("/public/pagine/") & "\" & NomePaginaDaEliminare) Then
				Set Documento = FSO.GetFile(Server.MapPath("/public/pagine/") & "\" & NomePaginaDaEliminare)
				Documento.Delete
				Set Documento = Nothing
			End If
			Set FSO = Nothing
			
			NomePaginaDaEliminare_en=rs("NomePagina_en")
			Set FSO = CreateObject("Scripting.FileSystemObject")
			If FSO.FileExists(Server.MapPath("/public/pagine/") & "\" & NomePaginaDaEliminare_en) Then
				Set Documento = FSO.GetFile(Server.MapPath("/public/pagine/") & "\" & NomePaginaDaEliminare_en)
				Documento.Delete
				Set Documento = Nothing
			End If
			Set FSO = Nothing
			
			'qui devono essere inserite tutte le tabelle dove compare FkCat_Prod per cancellare il record oppure metterlo a 0
			Set ss=Server.CreateObject("ADODB.Recordset")
			sql = "Select FkCategoria1 From Categorie2 where FkCategoria1="&pkid&""
			ss.Open sql, conn, 3, 3
				if ss.recordcount>0 then
					Do while not ss.EOF
						ss("FkCategoria1")=0
						ss.update
					ss.movenext
					loop
				end if
			ss.close
			
			rs.delete
			
			
		end if
		rs.update
		
		rs.close
	end if
	
	
	'modifica
	if mode=1 then
		if PkId=0 then
			SQL = " SELECT TOP 1 PkId FROM Categorie1 ORDER BY PkId DESC "
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.Open SQL, Conn, 1, 3
			If NOT RS.EOF Then
				RS.MoveFirst
				PkId = RS("PkId")
			End If
			Set RS = Nothing
			
			SQL = " UPDATE Categorie1 SET NomePagina = '"& ConvertiTitoloInNomeScript(Titolo, PkId) &"' WHERE PkId = "& PkId &" "
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.Open SQL, Conn, 3, 3
			Set RS = Nothing
			
			SQL = " UPDATE Categorie1 SET NomePagina_en = '"& ConvertiTitoloInNomeScript_en(Titolo_en, PkId) &"' WHERE PkId = "& PkId &" "
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.Open SQL, Conn, 3, 3
			Set RS = Nothing
		end if
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		If FSO.FileExists(percorso_pagine & "\" & NomePagina_Old) Then
			Set Documento = FSO.GetFile(percorso_pagine & "\" & NomePagina_Old)
			Documento.Delete
			Set Documento = Nothing
		End If
		Set FSO = Nothing
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		If FSO.FileExists(percorso_pagine & "\" & NomePagina_Old_en) Then
			Set Documento = FSO.GetFile(percorso_pagine & "\" & NomePagina_Old_en)
			Documento.Delete
			Set Documento = Nothing
		End If
		Set FSO = Nothing
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		Set Documento = FSO.OpenTextFile(percorso_pagine & "\" & ConvertiTitoloInNomeScript(Titolo, PkId), 2, True)
		ContenutoFile = ""
		ContenutoFile = ContenutoFile & "<" & "%" & vbCrLf
		ContenutoFile = ContenutoFile & "cat = "& PkId &"" & vbCrLf
		ContenutoFile = ContenutoFile & "%" & ">" & vbCrLf
		ContenutoFile = ContenutoFile & "<!--#include file=""inc_categorie.asp""-->"
		Documento.Write ContenutoFile
		Set Documento = Nothing
		Set FSO = Nothing
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		Set Documento = FSO.OpenTextFile(percorso_pagine & "\" & ConvertiTitoloInNomeScript_en(Titolo_en, PkId), 2, True)
		ContenutoFile = ""
		ContenutoFile = ContenutoFile & "<" & "%" & vbCrLf
		ContenutoFile = ContenutoFile & "cat = "& PkId &"" & vbCrLf
		ContenutoFile = ContenutoFile & "%" & ">" & vbCrLf
		ContenutoFile = ContenutoFile & "<!--#include file=""inc_categorie_en.asp""-->"
		Documento.Write ContenutoFile
		Set Documento = Nothing
		Set FSO = Nothing
		
		'response.Write("percorso:"&percorso_pagine)
		'response.Write("ContenutoFile:"&ContenutoFile)
		'response.End()
	
	end if
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
    <td height="20" colspan="2" valign="middle"><table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="159" class="menu-celle">&nbsp;Menu</td>
          <td width="267" class="menu-celle">Gestione Categorie liv.1</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-categoria1.asp">Elenco Categorie liv.1 &raquo;</a>&nbsp;&nbsp;<a href="sche-categoria1.asp">Nuova categoria &raquo;</a></td>
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
			<% if request("C1") <> "ON" then %>
                <% if mode = 1 and pkid = 0 then %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Categoria Inserita ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-categoria1.asp?ordine=<%=ordine%>";
			   					}
							//-->
							</script>
                <% else %>
                <% if mode = 1 then %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Categoria Aggiornata ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
								<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-categoria1.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
								//-->
								</script>
                <% else %>
				<table cellpadding="0" cellspacing="0" border="0" width="95%" class="admin-righe">
				  <tr> 
                	<td colspan="2">&nbsp;</td>
              	</tr> 	
					<form method="post" action="sche-categoria1.asp?mode=1&pkid=<%=pkid%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Posizione</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="posizione" class="form" size="3" maxlength="3" <%if pkid>0 then%> value="<%=rs("posizione")%>"<%end if%>></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Titolo</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="titolo" class="form" size="80" maxlength="100" <%if pkid>0 then%> value="<%=rs("titolo")%>"<%end if%>><input type="hidden" name="NomePagina" <%if pkid>0 then%> value="<%=rs("NomePagina")%>"<%end if%>></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Descrizione</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<textarea name="descrizione" cols="78" rows="5" class="form"><%if pkid>0 then%><%=rs("descrizione")%><%end if%></textarea></td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Testo 1</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="testo1" class="form" size="80" maxlength="255" <%if pkid>0 then%> value="<%=rs("testo1")%>"<%end if%>></td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Testo 2</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="testo2" class="form" size="80" maxlength="255" <%if pkid>0 then%> value="<%=rs("testo2")%>"<%end if%>></td>
                  </tr>
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2">&nbsp;</td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Titolo ENG</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="titolo_en" class="form" size="80" maxlength="100" <%if pkid>0 then%> value="<%=rs("titolo_en")%>"<%end if%>><input type="hidden" name="NomePagina_en" <%if pkid>0 then%> value="<%=rs("NomePagina_en")%>"<%end if%>></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Descrizione ENG</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<textarea name="descrizione_en" cols="78" rows="5" class="form"><%if pkid>0 then%><%=rs("descrizione_en")%><%end if%></textarea></td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Testo 1 ENG</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="testo1_en" class="form" size="80" maxlength="255" <%if pkid>0 then%> value="<%=rs("testo1_en")%>"<%end if%>></td>
                  </tr>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Testo 2 ENG</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="testo2_en" class="form" size="80" maxlength="255" <%if pkid>0 then%> value="<%=rs("testo2_en")%>"<%end if%>></td>
                  </tr>
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">
					<input name="Submit" type="submit" class="form" value="Salva" align="absmiddle"> 
                          &nbsp; <input name="Submit2" type="reset" class="form" value="Annulla"> 
                          &nbsp; <input type="checkbox" name="C1" value="ON" > 
                          &nbsp; Per cancellare la categoria </td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                </form>
				</table>
				<% end if %>
                <% end if %>
                <% else %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Categoria Cancellata ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-categoria1.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
							//-->
						</script>
                <% end if %>
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