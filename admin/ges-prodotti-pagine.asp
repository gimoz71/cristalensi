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
%>
<%			
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT pkid, titolo, nomepagina "
sql = sql + "FROM Prodotti "
nrs.Open sql, conn, 3, 3


Do While Not nrs.EOF
	titolo=nrs("titolo")
	pkid=nrs("pkid")
	
	NomePaginaDaEliminare=nrs("NomePagina")
	if NomePaginaDaEliminare<>"" then
		Set FSO = CreateObject("Scripting.FileSystemObject")
		If FSO.FileExists(Server.MapPath("/public/pagine/") & "\" & NomePaginaDaEliminare) Then
			Set Documento = FSO.GetFile(Server.MapPath("/public/pagine/") & "\" & NomePaginaDaEliminare)
			Documento.Delete
			Set Documento = Nothing
		End If
		Set FSO = Nothing
	end if
	
	nrs("NomePagina")=ConvertiTitoloInNomeScript(Titolo, PkId)
	nrs.update
	
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
	
	response.Write("pkid:"&pkid&"<br>")
nrs.movenext
loop

nrs.close
					
%>
<html>
<head>
<title>Cristalensi Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
</head>

<body>
</body>
</html>
<!--#include file="strClose.asp"-->
<!--#include file="chiusura.asp"-->