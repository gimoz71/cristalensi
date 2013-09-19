<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
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
		Risultato = IDArticolo & "p-" & Risultato & ".asp"
		ConvertiTitoloInNomeScript = Risultato
	End Function
	
	Function ConvertiTitoloInNomeScript_en(Titolo, IDArticolo)
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
		Risultato = IDArticolo & "p-en-" & Risultato & ".asp"
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
%>
<%
	mode = request("mode")
	if mode = "" then mode = 0

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Prodotti"
	if pkid > 0 then sql = "Select * From Prodotti where pkid="&pkid
	rs.Open sql, conn, 3, 3
		
	if mode = 1 then
		if pkid = 0 then rs.addnew
		
		Titolo=request("Titolo")
		rs("titolo")=Titolo
		
		Titolo_en=request("Titolo_en")
		rs("titolo_en")=Titolo_en
		
		NomePagina_Old=request("NomePagina")
		NomePagina_Old_en=request("NomePagina_en")
		
		if pkid>0 then
			rs("NomePagina")=ConvertiTitoloInNomeScript(Titolo, PkId)
			rs("NomePagina_en")=ConvertiTitoloInNomeScript_en(Titolo_en, PkId)
		end if
		
		testo = request("testo")
		testo=Replace(testo, """", "'")
		testo=Replace(testo, vbcrlf, "")	
		rs("descrizione")=testo
		
		testo_en = request("testo_en")
		testo_en=Replace(testo_en, """", "'")
		testo_en=Replace(testo_en, vbcrlf, "")	
		rs("descrizione_en")=testo_en
		
		Offerta=request("Offerta")
		if Offerta="" then Offerta=0
		rs("Offerta")=Offerta
		
		PrimoPiano=request("PrimoPiano")
		if PrimoPiano="si" then rs("PrimoPiano")=True
		if PrimoPiano="no" then rs("PrimoPiano")=False
		
		rs("FkCategoria2") = request("FkCategoria2")
		rs("FkProduttore") = request("FkProduttore")
		rs("CodiceArticolo") = request("CodiceArticolo")
		rs("codicearticolo_azienda") = request("codicearticolo_azienda")
		rs("Allegato") = request("Allegato")
		rs("PrezzoProdotto") = request("PrezzoProdotto")
		rs("PrezzoListino") = request("PrezzoListino")
		
		'aggiornamento colori
		if pkid>0 then
			Set pps=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From [Prodotto-Colore] where FkProdotto="&pkid&" "
			pps.Open sql, conn, 3, 3
			if pps.recordcount>0 then
				Do while not pps.EOF
					pps.delete
				pps.movenext
				loop
			end if
			pps.close
			
			fkcolori=request("fkcolori")
			arrFkcolore=split(fkcolori,", ")
		
			if fkcolori<>"" then
				For iLoop = LBound(arrFkcolore) to UBound(arrFkcolore)	
					fkcolore=arrFkcolore(iLoop)
					fkcolore=cInt(fkcolore)
					Set pps=Server.CreateObject("ADODB.Recordset")
					sql = "Select * From [Prodotto-Colore]"
					pps.Open sql, conn, 3, 3	
					pps.addnew
					pps("fkcolore")=fkcolore
					pps("fkprodotto")=pkid
					pps.update
					pps.close
				Next
			end if
		end if
		
		
		'aggiornamento lampadine
		if pkid>0 then
			Set pps=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From [Prodotto-Lampadina] where FkProdotto="&pkid&" "
			pps.Open sql, conn, 3, 3
			if pps.recordcount>0 then
				Do while not pps.EOF
					pps.delete
				pps.movenext
				loop
			end if
			pps.close
			
			fklampadine=request("fklampadine")
			arrFklampadina=split(fklampadine,", ")
		
			if fklampadine<>"" then
				For iLoop = LBound(arrFklampadina) to UBound(arrFklampadina)	
					fklampadina=arrFklampadina(iLoop)
					fklampadina=cInt(fklampadina)
					Set pps=Server.CreateObject("ADODB.Recordset")
					sql = "Select * From [Prodotto-Lampadina]"
					pps.Open sql, conn, 3, 3	
					pps.addnew
					pps("fklampadina")=fklampadina
					pps("fkprodotto")=pkid
					pps.update
					pps.close
				Next
			end if
		end if
		
		
				
		if request("C1") = "ON" then
			rs.delete
			
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
			
			'elimino i colori
			Set pps=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From [Prodotto-Colore] where FkProdotto="&pkid&" "
			pps.Open sql, conn, 3, 3
			if pps.recordcount>0 then
				Do while not pps.EOF
					pps.delete
				pps.movenext
				loop
			end if
			pps.close
			
			'elimino le lampadine
			Set pps=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From [Prodotto-Lampadina] where FkProdotto="&pkid&" "
			pps.Open sql, conn, 3, 3
			if pps.recordcount>0 then
				Do while not pps.EOF
					pps.delete
				pps.movenext
				loop
			end if
			pps.close
			
		end if
		rs.update
		
		rs.close
	end if
	
	if mode=1 then
		if PkId=0 then
			SQL = " SELECT TOP 1 PkId FROM Prodotti ORDER BY PkId DESC "
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.Open SQL, Conn, 1, 3
			If NOT RS.EOF Then
				RS.MoveFirst
				PkId = RS("PkId")
			End If
			Set RS = Nothing
			
			SQL = " UPDATE Prodotti SET NomePagina = '"& ConvertiTitoloInNomeScript(Titolo, PkId) &"' WHERE PkId = "& PkId &" "
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.Open SQL, Conn, 3, 3
			Set RS = Nothing
			
			SQL = " UPDATE Prodotti SET NomePagina_en = '"& ConvertiTitoloInNomeScript_en(Titolo_en, PkId) &"' WHERE PkId = "& PkId &" "
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.Open SQL, Conn, 3, 3
			Set RS = Nothing
			
			'inserisco i colori
			fkcolori=request("fkcolori")
			arrFkcolore=split(fkcolori,", ")
		
			if fkcolori<>"" then
				For iLoop = LBound(arrFkcolore) to UBound(arrFkcolore)	
					fkcolore=arrFkcolore(iLoop)
					fkcolore=cInt(fkcolore)
					Set pps=Server.CreateObject("ADODB.Recordset")
					sql = "Select * From [Prodotto-Colore]"
					pps.Open sql, conn, 3, 3	
					pps.addnew
					pps("fkcolore")=fkcolore
					pps("fkprodotto")=pkid
					pps.update
					pps.close
				Next
			end if
			
			'inserisco le lampadine
			fklampadine=request("fklampadine")
			arrFklampadina=split(fklampadine,", ")
		
			if fklampadine<>"" then
				For iLoop = LBound(arrFklampadina) to UBound(arrFklampadina)	
					fklampadina=arrFklampadina(iLoop)
					fklampadina=cInt(fklampadina)
					Set pps=Server.CreateObject("ADODB.Recordset")
					sql = "Select * From [Prodotto-Lampadina]"
					pps.Open sql, conn, 3, 3	
					pps.addnew
					pps("fklampadina")=fklampadina
					pps("fkprodotto")=pkid
					pps.update
					pps.close
				Next
			end if
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
		ContenutoFile = ContenutoFile & "id = "& PkId &"" & vbCrLf
		ContenutoFile = ContenutoFile & "%" & ">" & vbCrLf
		ContenutoFile = ContenutoFile & "<!--#include file=""inc_scheda_prodotto.asp""-->"
		Documento.Write ContenutoFile
		Set Documento = Nothing
		Set FSO = Nothing
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		Set Documento = FSO.OpenTextFile(percorso_pagine & "\" & ConvertiTitoloInNomeScript_en(Titolo_en, PkId), 2, True)
		ContenutoFile = ""
		ContenutoFile = ContenutoFile & "<" & "%" & vbCrLf
		ContenutoFile = ContenutoFile & "id = "& PkId &"" & vbCrLf
		ContenutoFile = ContenutoFile & "%" & ">" & vbCrLf
		ContenutoFile = ContenutoFile & "<!--#include file=""inc_scheda_prodotto_en.asp""-->"
		Documento.Write ContenutoFile
		Set Documento = Nothing
		Set FSO = Nothing
		
		'response.Write("percorso:"&percorso_pagine)
		'response.Write("ContenutoFile:"&ContenutoFile)
		'response.End()
	
	end if
	
	if mode=0 AND pkid>0 then
		testo=rs("descrizione")
		testo_en=rs("descrizione_en")
	else
		testo=""
		testo_en=""
	end if
%>
<html>
<head>
<title>Cristalensi Admin</title>
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

<body>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0" class="TAB_centrale">
  <!--#include file="testata.asp"-->
  <tr>
    <td height="30" colspan="2" valign="middle"><table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="159" class="menu-celle">&nbsp;Menu</td>
          <td width="267" class="menu-celle">Gestione prodotti</td>
          <td width="324" class="menu-celle" align="right"><a href="ges-prodotti.asp">Elenco prodotti &raquo;</a>&nbsp;&nbsp;<a href="sche-prodotti.asp">Nuovo prodotto &raquo;</a></td>
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
                <p class="admin-righe"> Record Inserito ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-prodotti.asp?ordine=<%=ordine%>";
			   					}
							//-->
							</script>
                <% else %>
                <% if mode = 1 then %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Record Aggiornato ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
								<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-prodotti.asp?p=<%=p%>&ordine=<%=ordine%>";
			   					}
								//-->
								</script>
                <% else %>
				<table cellpadding="0" cellspacing="0" border="0" class="admin-righe">
					<tr align="left">
                    <td height="15" colspan="2">&nbsp;</td>
                  </tr>
					<form method="post" action="sche-prodotti.asp?mode=1&pkid=<%=pkid%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                  <tr align="left">
                    <td><strong>Codice articolo pubblico</strong></td>
                    <td><strong>Codice articolo azienda</strong></td>
				  </tr>
				  <tr align="left">
                    <td><input type="text" name="codicearticolo" <%if pkid>0 then%> value="<%=rs("codicearticolo")%>"<%end if%> class="form" size="20"></td>
                    <td><input type="text" name="codicearticolo_azienda" <%if pkid>0 then%> value="<%=rs("codicearticolo_azienda")%>"<%end if%> class="form" size="20"></td>
				  </tr>
				  <tr align="left">
                    <td><strong>Categoria liv.2 </strong></td>
                    <td><strong>Produttore</strong></td>
				  </tr>
				  <tr align="left">
                    <td>
					<%
					Set cs=Server.CreateObject("ADODB.Recordset")
					sql = "SELECT Categorie1.PkId as PkId_1, Categorie1.Titolo as Titolo_1, Categorie2.PkId as PkId_2, Categorie2.Titolo as Titolo_2 "
					sql = sql + "FROM Categorie1 INNER JOIN Categorie2 ON Categorie1.PkId = Categorie2.Fkcategoria1 "
					sql = sql + "ORDER BY Categorie1.Titolo ASC, Categorie2.Titolo ASC"
					cs.Open sql, conn, 1, 1
					%>
					<select name="FkCategoria2" class="form">
                        <option value=0 <% if pkid = 0 then %> selected<%end if%>>Nessuna categoria</option>
						<%
						if cs.recordcount>0 then
						Do While Not cs.EOF
						%>
                        <option value=<%=cs("pkid_2")%> <% if pkid > 0 then %><%if cInt(rs("FkCategoria2"))=cInt(cs("pkid_2")) then%> selected<%end if%><%end if%>><%=cs("Titolo_1")%> - <%=cs("Titolo_2")%></option>
                        <%
						cs.movenext
						loop
						end if
						%>
                     </select>
					 <%cs.close%>					</td>
                    <td>
					<%
					Set cs=Server.CreateObject("ADODB.Recordset")
					sql = "Select * From Produttori order by titolo ASC"
					cs.Open sql, conn, 1, 1
					%>
					<select name="FkProduttore" class="form">
                        <option value=0 <% if pkid = 0 then %> selected<%end if%>>Nessun produttore</option>
						<%
						if cs.recordcount>0 then
						Do While Not cs.EOF
						%>
                        <option value=<%=cs("pkid")%> <% if pkid > 0 then %><%if rs("FkProduttore")=cs("pkid") then%> selected<%end if%><%end if%>><%=cs("titolo")%></option>
                        <%
						cs.movenext
						loop
						end if
						%>
                     </select>
					 <%cs.close%>
					</td>
				  </tr>
				  
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Titolo</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="titolo" class="form" size="80" maxlength="100" <%if pkid>0 then%> value="<%=rs("titolo")%>"<%end if%>></td>
                  </tr>
				  
				  
				  <tr align="left">
                    <td height="15"><strong>Prezzo Cristalensi</strong></td>
                    <td height="15"><strong>Prezzo Listino</strong></td>
				  </tr>
				  <tr align="left">
                    <td height="15">
					<input type="text" name="PrezzoProdotto" class="form" size="10" maxlength="50" <%if pkid>0 then%> value="<%=rs("PrezzoProdotto")%>"<%end if%>>
					&euro;</td>
                    <td height="15">
					<input type="text" name="PrezzoListino" class="form" size="10" maxlength="50" <%if pkid>0 then%> value="<%=rs("PrezzoListino")%>"<%end if%>>
					</td>
				  </tr>
				  <tr align="left">
                    <td height="15" colspan="2"><strong>Allegato</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="Allegato" id="Allegato" class="form" size="30" maxlength="100" <%if pkid>0 then%> value="<%=rs("Allegato")%>"<%end if%>> per inserire un file, <a href="#" onClick="MM_openBrWindow('upload-file.asp','','width=300,height=300')">cliccare qui.</a></td>
                  </tr>
				  <tr align="left">
                    <td width="388"><strong>Dove viene visualizzato il prodotto</strong></td>
                    <td width="199"><strong>In primo piano</strong></td>
                  </tr>
				  <tr align="left">
                    <td width="388">Solo nei prodotti 
                    <input name="Offerta" type="radio" value="0" <% if pkid > 0 then %><%if rs("Offerta")=0 then%>checked<%end if%><%else%>checked<%end if%>>
                    &nbsp;&nbsp;Solo nelle offerte
                    <input name="Offerta" type="radio" value="1" <% if pkid > 0 then %><%if rs("Offerta")=1 then%>checked<%end if%><%end if%>>
                    <br>
                    Sia nei prodotti che nelle offerte 
                    <input name="Offerta" type="radio" value="2" <% if pkid > 0 then %><%if rs("Offerta")=2 then%>checked<%end if%><%end if%>>&nbsp;&nbsp;Non visibile 
                    <input name="Offerta" type="radio" value="10" <% if pkid > 0 then %><%if rs("Offerta")=10 then%>checked<%end if%><%end if%>></td>
                    <td width="199">
					Si 
				    <input name="PrimoPiano" type="radio" value="si" <% if pkid > 0 then %><%if rs("PrimoPiano")=True then%>checked<%end if%><%end if%>>&nbsp;&nbsp;No <input name="PrimoPiano" type="radio" value="no" <% if pkid > 0 then %><%if rs("PrimoPiano")=False then%>checked<%end if%><%else%>checked<%end if%>>					</td>
                  </tr>
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  <%
					Set ns=Server.CreateObject("ADODB.Recordset")
					sql = "Select * From Lampadine order by Titolo ASC"
					ns.Open sql, conn, 1, 1
					totale=ns.recordcount
					conta=0
					if ns.recordcount>0 then			  
				   %>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Lampadine</strong></td>
                  </tr>
                  <tr align="left"><td height="15" colspan="2"><table width="100%" class="admin-righe">
                  <%
					Do While not ns.EOF
					If((conta Mod 3)=0) then
				  %>
				  <tr align="left">
                  <%end if%>
                    <td>
                    <%
					pkid_lampadina=ns("pkid")
					pkid_lampadina=cInt(pkid_lampadina)
					if pkid>0 then
						pkid=cInt(pkid)	
						esiste=""
						Set ps=Server.CreateObject("ADODB.Recordset")
						sql = "SELECT [Prodotto-Lampadina].FkProdotto, [Prodotto-Lampadina].FkLampadina FROM [Prodotto-Lampadina] WHERE ((([Prodotto-Lampadina].FkProdotto)="&pkid&") AND (([Prodotto-Lampadina].FkLampadina)="&pkid_lampadina&"))"
						ps.Open sql, conn, 1, 1
						if ps.recordcount=1 then
							esiste="Si"
						else
							esiste="No"
						end if
						ps.close
					else
						esiste="No"
					end if
					%>
					
					<input name="fklampadine" type="checkbox" value=<%=pkid_lampadina%> <%if esiste="Si" then%> checked<%end if%>>&nbsp;<%=ns("Titolo")%> / <%=ns("Titolo_en")%>
					</td>
                    <%If((conta Mod 3)=2) then%>
                  </tr>
                  <tr> </tr>
                  <%
					End if
					conta=conta+1
					ns.movenext
					loop
				  %>
                  </table></td></tr>
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  <%end if%>
                  <%ns.close%>
                  
				   <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  <%
					Set ns=Server.CreateObject("ADODB.Recordset")
					sql = "Select * From Colori order by Titolo ASC"
					ns.Open sql, conn, 1, 1
					totale=ns.recordcount
					conta=0
					if ns.recordcount>0 then			  
				   %>
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Colori</strong></td>
                  </tr>
                  <tr align="left"><td height="15" colspan="2"><table width="100%" class="admin-righe">
                  <%
					Do While not ns.EOF
					If((conta Mod 3)=0) then
				  %>
				  <tr align="left">
                  <%end if%>
                    <td>
                    <%
					pkid_colore=ns("pkid")
					pkid_colore=cInt(pkid_colore)
					if pkid>0 then
						pkid=cInt(pkid)	
						esiste=""
						Set ps=Server.CreateObject("ADODB.Recordset")
						sql = "SELECT [Prodotto-Colore].FkProdotto, [Prodotto-Colore].FkColore FROM [Prodotto-Colore] WHERE ((([Prodotto-Colore].FkProdotto)="&pkid&") AND (([Prodotto-Colore].FkColore)="&pkid_colore&"))"
						ps.Open sql, conn, 1, 1
						if ps.recordcount=1 then
							esiste="Si"
						else
							esiste="No"
						end if
						ps.close
					else
						esiste="No"
					end if
					%>
					
					<input name="fkcolori" type="checkbox" value=<%=pkid_colore%> <%if esiste="Si" then%> checked<%end if%>>&nbsp;<%=ns("Titolo")%> / <%=ns("Titolo_en")%>
					</td>
                    <%If((conta Mod 3)=2) then%>
                  </tr>
                  <tr> </tr>
                  <%
					End if
					conta=conta+1
					ns.movenext
					loop
				  %>
                  </table></td></tr>
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  <%end if%>
                  <%ns.close%>
                  <tr align="left">
                    <td height="20" colspan="2"><strong>Testo</strong></td>
                  </tr>
                  <tr><td colspan="2" align="center"><%
' Automatically calculates the editor base path based on the _samples directory.
' This is usefull only for these samples. A real application should use something like this:
' oFCKeditor.BasePath = '/fckeditor/' ;	// '/fckeditor/' is the default value.
Dim sBasePath
'sBasePath = Request.ServerVariables("PATH_INFO")
'sBasePath = Left( sBasePath, InStrRev( sBasePath, "/admin" ) )
sBasePath = "/admin/FCKeditor/"

Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath	= sBasePath

oFCKeditor.ToolbarSet = Server.HTMLEncode( "Default" )

oFCKeditor.Value	= testo
oFCKeditor.Create "testo"
%></td></tr>
                   <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  
                  <tr align="left">
                    <td height="15" colspan="2"><strong>Titolo ENG</strong></td>
                  </tr>
				  <tr align="left">
                    <td height="15" colspan="2">
					<input type="text" name="titolo_en" class="form" size="80" maxlength="100" <%if pkid>0 then%> value="<%=rs("titolo_en")%>"<%end if%>></td>
                  </tr>
                  <tr align="left">
                    <td height="20" colspan="2"><strong>Testo ENG</strong></td>
                  </tr>
                  <tr><td colspan="2" align="center"><%
' Automatically calculates the editor base path based on the _samples directory.
' This is usefull only for these samples. A real application should use something like this:
' oFCKeditor.BasePath = '/fckeditor/' ;	// '/fckeditor/' is the default value.
'Dim sBasePath
'sBasePath = Request.ServerVariables("PATH_INFO")
'sBasePath = Left( sBasePath, InStrRev( sBasePath, "/admin" ) )
sBasePath = "/admin/FCKeditor/"

'Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath	= sBasePath

oFCKeditor.ToolbarSet = Server.HTMLEncode( "Default" )

oFCKeditor.Value	= testo_en
oFCKeditor.Create "testo_en"
%></td></tr>
                   <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  
                  <tr align="left">
                    <td colspan="2">
                      <input name="Submit" type="submit" class="form" value="Salva" align="absmiddle"> 
                          &nbsp; <input name="Submit2" type="reset" class="form" value="Annulla">
						  <%if pkid>0 then%>&nbsp;<input type="checkbox" name="C1" value="ON" >&nbsp; Per cancellare il prodotto<%end if%>
                    </td>
                  </tr>
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
				  
				  <tr align="left">
                    <td height="20" colspan="2"><strong>Photo gallery</strong></td>
                  </tr>
				  <%if pkid>0 then%>
					  <tr> 
                        <td colspan="2" align="center"><iframe width="548" height="280" src="ins_file.asp?id=<%=pkid%>&tab=Prodotti"></iframe></td>
                      </tr>
					<%else%>
					  <tr> 
                        <td height="20" colspan="2">Per inserire una o più foto/immagini è necessario salvare il prodotto e rientrare nello stesso.</td>
                      </tr>
				  <%end if%>
				  
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                </form>
				</table>
				<% end if %>
                <% end if %>
                <% else %>
                <p>&nbsp;</p>
                <p class="admin-righe"> Record Cancellato ....<br>
    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</p>
                <SCRIPT LANGUAGE="JavaScript">
							<!--
			   					setTimeout("update()",2000);
			   					function update(){
			        			document.location.href = "ges-prodotti.asp?p=<%=p%>&ordine=<%=ordine%>";
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