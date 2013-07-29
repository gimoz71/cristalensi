<%
On Error Resume Next

Set conn = Server.CreateObject("ADODB.Connection")
	'conn.open = "DRIVER={Microsoft Access Driver (*.mdb)};dbq=d:\inetpub\webs\cristalensiit\mdb-database\db_cristalensi.mdb"
	'conn.open = "DRIVER={Microsoft Access Driver (*.mdb)};dbq=d:\inetpub\webs\viadeimediciit\mdb-database\db_cristalensi.mdb"
	conn.open = "DRIVER={Microsoft Access Driver (*.mdb)};dbq="& Server.MapPath("/mdb-database/db_cristalensi.mdb")
	'conn.open = "DSN=cristalensi"
	
Set conn2 = Server.CreateObject("ADODB.Connection")
	'conn2.open = "DRIVER={Microsoft Access Driver (*.mdb)};dbq=d:\inetpub\webs\cristalensiit\mdb-database\db_cristalensi_privato.mdb"
	'conn2.open = "DRIVER={Microsoft Access Driver (*.mdb)};dbq=d:\inetpub\webs\viadeimediciit\mdb-database\db_cristalensi_privato.mdb"
	conn2.open = "DRIVER={Microsoft Access Driver (*.mdb)};dbq="& Server.MapPath("/mdb-database/db_cristalensi_privato.mdb")
		
If Err.Number <> 0 Then
	Response.Redirect("../aggiornamento.htm")
End IF

'percorso_pagine="D:\web\cristalensi\new\public\pagine"   'locale
percorso_pagine=Server.MapPath("/public/pagine/")  'online
%>

