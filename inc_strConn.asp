<%
'On Error Resume Next

	Set conn = Server.CreateObject("ADODB.Connection")
	'conn.open = "DRIVER={Microsoft Access Driver (*.mdb)};dbq=d:\inetpub\webs\cristalensiit\mdb-database\db_cristalensi.mdb"
	'conn.open = "DRIVER={Microsoft Access Driver (*.mdb)};dbq=d:\inetpub\webs\viadeimediciit\mdb-database\db_cristalensi.mdb"
	conn.open = "DRIVER={Microsoft Access Driver (*.mdb)};dbq="& Server.MapPath("mdb-database/db_cristalensi.mdb")
	'conn.open = "DSN=cristalensiit"
	
	path_img="d:\inetpub\webs\cristalensiit\public\"
	
	fromURL = Request.ServerVariables("HTTP_REFERER")
	toUrl = Request.ServerVariables("SCRIPT_NAME")
	
	
	'strDaDoveVengo = Request.Servervariables("HTTP_REFERER")
	UltimoSlash1 = InStrRev(fromURL,"/")
	fromURL = Mid((fromURL),(UltimoSlash1 + 1), len(fromURL)- UltimoSlash1)
	'Response.Write "La pagina di provenienza è: " &fromURL& ".<br>"
	
	UltimoSlash2 = InStrRev(toUrl,"/")
	toUrl = Mid((toUrl),(UltimoSlash2 + 1), len(toUrl)- UltimoSlash2)
	'Response.Write "La pagina dove sono è: " &toURL& "."
	
'If Err.Number <> 0 Then
	'Response.Redirect("aggiornamento.htm")
'End IF

MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString

'controllo per il log
contr=request("contr")
if contr="" then contr=0

if contr=1 then
	login = Request.form("login")
	lg1=InStr(login, "'")
	if lg1>0 then
		login=Replace(login, "'", " ")	
		'response.End()
	end if
	lg2=InStr(login, "&")
	if lg2>0 then
		login=Replace(login, "&", " ")	
		'response.End()
	end if
	login=Trim(login)
	password = Request.form("Password")
	pw1=InStr(password, "'")
	if pw1>0 then
		password=Replace(password, "'", " ")	
		'response.End()
	end if
	pw2=InStr(password, "&")
	if pw2>0 then
		password=Replace(password, "&", " ")	
		'response.End()
	end if
	password=Trim(password)


	Set log_rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Clienti WHERE Email='" & login & "' AND Password='" & password & "'"
	log_rs.open sql,conn

	if not log_rs.eof then
		idsession=log_rs("PkId")
		nome_log=log_rs("Nominativo")
		if nome_log="" then nome_log="Anonimo"
		italia_log=log_rs("Italia")
		if italia_log="" then italia_log="Si"
		if italia_log="Sì" then italia_log="Si"
	
		Session("idCliente") = idsession
		Session("nome_log") = nome_log
		Session("italia_log") = italia_log
	else
		contr=2
	end if
	log_rs.close
	set log_rs = nothing
else
	nome_log=Session("nome_log")
	italia_log=Session("italia_log")
	idsession=Session("idCliente")
	if idsession="" then idsession=0
	if italia_log="" then italia_log="Si"
end if

'funzione che mi registra il passaggio da un pagina, un'eventuale tabella aperta e il record aperto
Sub Visualizzazione(tabella,record,pagina)
'	Set conn2 = Server.CreateObject("ADODB.Connection")
'	conn2.open = "DRIVER={Microsoft Access Driver (*.mdb)};dbq=d:\inetpub\webs\cristalensiit\mdb-database\db_cristalensi_privato.mdb"
'	
'	Set vis_rs = Server.CreateObject("ADODB.Recordset")
'	sql = "SELECT * FROM Visualizzazioni"
'	vis_rs.open sql,conn2, 3, 3
'	vis_rs.addnew
'		vis_rs("DataOra_Inserimento")=Now()
'		vis_rs("IP")=Request.ServerVariables("REMOTE_ADDR")
'		vis_rs("Tabella")=tabella
'		vis_rs("Record")=record
'		vis_rs("Pagina")=pagina
'		
'	vis_rs.update
'	vis_rs.close
'	
'	conn2.close
'	set conn2 = nothing
End Sub

Function TogliTAG(Stringa)
   Dim RegEx, Temp

   Temp = Stringa
   Set RegEx = New RegExp
   RegEx.Pattern = "<[^>]*>"
   RegEx.Global = True
   RegEx.IgnoreCase = True
   Temp = RegEx.Replace(Temp, "")
   Set RegEx = Nothing

   TogliTAG = Temp
End Function

Function NoHTML(strInput) 
 
 Dim RegEx 
 Set RegEx = New RegExp 
 RegEx.Pattern = "<[^>]*>" 
 RegEx.Global = True 
 RegEx.IgnoreCase = True 
 
        ' conserva la formattazione 
 strInput = Replace(strInput, "<br>", chr(10))
 strInput = Replace(strInput, "'", "")
 strInput = Replace(strInput, """", "")
 
 strInput = Replace(strInput, "é", "&eacute;")
 strInput = Replace(strInput, "è", "&egrave;")
 strInput = Replace(strInput, "à", "&agrave;")
 strInput = Replace(strInput, "ù", "&ugrave;")
 strInput = Replace(strInput, "ì", "&igrave;")
 strInput = Replace(strInput, "ò", "&ograve;")
 
 NoHTML = RegEx.Replace(strInput, "") 
 
End Function


Function NoLettAcc(strInput) 
  
 strInput = Replace(strInput, "é", "&eacute;")
 strInput = Replace(strInput, "è", "&egrave;")
 strInput = Replace(strInput, "à", "&agrave;")
 strInput = Replace(strInput, "ù", "&ugrave;")
 strInput = Replace(strInput, "ì", "&igrave;")
 strInput = Replace(strInput, "ò", "&ograve;")
 strInput = Replace(strInput, "€", "&#8364;")
 strInput = Replace(strInput, "'", "&#8217;")
 
 NoLettAcc = strInput 
 
End Function
%>

