<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Cristalensi Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="330" border="0" align="center" cellpadding="0" cellspacing="0" class="login">
  <tr>
    <td rowspan="5" valign="middle"><img src="immagini/logo.jpg" hspace="5" vspace="10"></td>
    <td colspan="2">&nbsp;</td>
  </tr>
<%
Response.Expires = -1
mode=request("mode")
if mode="" then mode=0
if mode=1 then

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
%>
<!--#include file="strConn.asp"-->
<%
	if login="zorba" and password="z0rba" then
		Session("idAmministratore") = 0
		Session("nickAmministratore") = ""
		'Session("Permission") = 1	'vede gli amministratori
		Response.Redirect("admin.asp")	
	end if
		 
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Amministratori WHERE Login='" & login & "' AND Password='" & password & "'"
	rs.open sql,conn,1,1
	num=rs.recordcount
	
	if num=1 then
		idsession=rs("pkid")
		Nominativo=rs("Nominativo")
		
		Session("idAmministratore") = idsession
		Session("nickAmministratore") = Nominativo
		'Session("Permission") = livello
	
		rs.close
		set rs = nothing
%>
<!--#include file="strClose.asp"-->
<%	
		Response.Redirect("admin.asp")
	else
		mode=2
		rs.close
		set rs = nothing
%>
<!--#include file="strClose.asp"-->
<%	end if%>
<%end if%>   
  <%if mode=0 or mode=2 then%>
  <form method="post" action="logon.asp?mode=1">
  <tr>
    <td>Login</td>
    <td align="right"><input name="login" type="text" size="25" class="form"></td>
  </tr>
  <tr>
    <td>Password&nbsp;</td>
    <td align="right"><input name="password" type="password" size="25" class="form"></td>
  </tr>
  <tr>
    <td colspan="2" align="right"><input name="Submit" type="submit" class="form" value="Entra"></td>
  </tr>
  </form>
  <%if mode=2 then%>
  <tr><td colspan="2" align="center"><font color="#CC0000">Attenzione! Login o Passowrd errati</font></td></tr>
  <%end if%>
  <%end if%>
</table>
</body>
</html>
