<!--#include file="inc_strConn.asp"-->
<%
	Call Visualizzazione("",0,"pagamento_paypal_ok.asp")
	
	IdOrdine=request("item_number")
	if IdOrdine="" then IdOrdine=0
	
	Set ss = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where pkid="&idOrdine
	ss.Open sql, conn, 3, 3
	
	if ss.recordcount>0 then
		TotaleCarrello=ss("TotaleCarrello")
		CostoSpedizioneTotale=ss("CostoSpedizione")
		TipoTrasporto=ss("TipoTrasporto")
		DatiSpedizione=ss("DatiSpedizione")
		NoteCliente=ss("NoteCliente")
		
		FkPagamento=ss("FkPagamento")
		TipoPagamento=ss("TipoPagamento")
		CostoPagamento=ss("CostoPagamento")
		
		Nominativo=ss("Nominativo")
		Rag_Soc=ss("Rag_Soc")
		Cod_Fisc=ss("Cod_Fisc")
		PartitaIVA=ss("PartitaIVA")
		Indirizzo=ss("Indirizzo")
		Citta=ss("Citta")
		Provincia=ss("Provincia")
		CAP=ss("CAP")
		
		TotaleGenerale=ss("TotaleGenerale")
		
		DataAggiornamento=ss("DataAggiornamento")
		
		ss("stato")=4
		ss("DataAggiornamento")=now()
		ss("IpOrdine")=Request.ServerVariables("REMOTE_ADDR")
		ss.update	
	end if
	
	ss.close
	
	if FkPagamento=2 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Clienti where pkid="&idsession
		rs.Open sql, conn, 1, 1	
		
		nominativo_email=rs("nominativo")
		email=rs("email")
		
		rs.close
			
			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Cristalensi</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Grazie "&nominativo_email&" per aver scelto i nostri prodotti!<br>Questa &egrave; un email di conferma per il completamento dell'ordine n&deg; "&idordine&".<br> Il nostro staff avr&agrave; cura di spedirti la merce appena la banca avr&agrave; notificato il pagamento con Paypal.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Cristalensi</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = email
			Oggetto = "Conferma pagamento ordine n "&idordine&" con Paypal a Cristalensi.it"
			Testo = HTML1

'			Set eMail_cdo = CreateObject("CDO.Message")
'
'			eMail_cdo.From = Mittente
'			eMail_cdo.To = Destinatario
'			eMail_cdo.Subject = Oggetto
'
'			eMail_cdo.HTMLBody = Testo
'
'			eMail_cdo.Send()
'
'			Set eMail_cdo = Nothing
			
			'fine invio email
			
			'invio l'email all'amministratore
			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Cristalensi</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento da Paypal dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"
		
			Mittente = "info@cristalensi.it"
			Destinatario = "info@cristalensi.it"
			Oggetto = "Conferma pagamento ordine n "&idordine&" con Paypal a Cristalensi.it"
			Testo = HTML1

'			Set eMail_cdo = CreateObject("CDO.Message")
'
'			eMail_cdo.From = Mittente
'			eMail_cdo.To = Destinatario
'			eMail_cdo.Subject = Oggetto
'
'			eMail_cdo.HTMLBody = Testo
'
'			eMail_cdo.Send()
'
'			'invio al webmaster
'			
			Set eMail_cdo = Nothing
			
			Mittente = "info@cristalensi.it"
			Destinatario = "iurymazzoni@hotmail.com"
			Oggetto = "Conferma pagamento ordine n "&idordine&" con Paypal a Cristalensi.it"
			Testo = HTML1

'			Set eMail_cdo = CreateObject("CDO.Message")
'
'			eMail_cdo.From = Mittente
'			eMail_cdo.To = Destinatario
'			eMail_cdo.Subject = Oggetto
'
'			eMail_cdo.HTMLBody = Testo
'
'			eMail_cdo.Send()
'
'			Set eMail_cdo = Nothing
			
			'fine invio email
	end if
%>
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Cristalensi</title>
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="js/media-queries-ie.js"></script>
        <![endif]-->
        <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
        <script src="js/jquery.blueberry.js"></script>
        <script src="js/jquery.tipTip.js"></script>
        <link href="css/css.css" rel="stylesheet" type="text/css">
        <link href="css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="css/tipTip.css" rel="stylesheet" type="text/css">
        <style type="text/css">
            .clearfix:after {
                content: ".";
                display: block;
                height: 0;
                clear: both;
                visibility: hidden;
            }
        </style>
        <!--[if lt IE 8]>
            <link href="/css/tipTip_ie7.css" media="all" rel="stylesheet" type="text/css" />
        <![endif]-->
        <!--[if IE]>
            <style type="text/css">
                .clearfix {
                    zoom: 1;   /* triggers hasLayout */
                }   /* Only IE can see inside the conditional comment
                    and read this CSS rule. Don't ever use a normal HTML
                    comment inside the CC or it will close prematurely. */
            </style>
        <![endif]-->
    </head>
    <body>
        <div id="wrap">
            <!--#include file="inc_header.asp"-->

            <div id="main-content">
               
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <h3 style="font-size: 14px; display: inline; border: none;">Ordine n&deg; <%=idordine%> - Data <%=Left(DataAggiornamento, 10)%></h3>
                            <div class="carrello clearfix">
                                <p class="testo_grande">
                                	La procedura di pagamento con Paypal &egrave; stata completata correttamente.<br>
                                          <br>					    
                                      La merce verr&agrave; spedita al momento che la nostra banca ricever&agrave; il pagamento.<br>
                                      <br>
                                      Potrai seguire lo stato del tuo ordine direttamente dall'area clienti, comunque sar&agrave; cura del nostro staff informarti per email dell'invio dei prodotti ordinati.
                                      <br>
                                      Cordiali saluti, lo staff di Cristalensi
                                      <br>
                                      <br>
                                </p>
                                	
                                
                                <p class="area clearfix"><span class="colonna articolo">[Codice articolo] Nome prodotto</span><span class="colonna quantita">quantit&agrave;</span><span class="colonna prezzo_unitario">prezzo unitario</span><span class="colonna prezzo_totale">prezzo totale</span></p>
                                <div class="data">
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT PkId, FkOrdine, FkProdotto, PrezzoProdotto, Quantita, TotaleRiga, Titolo, CodiceArticolo, Colore FROM RigheOrdine WHERE FkOrdine="&idOrdine&""
	rs.Open sql, conn, 1, 1
	num_prodotti_carrello=rs.recordcount
	
%>                                    
									<%if rs.recordcount>0 then%>
                                        
                                        <%
                                        Do while not rs.EOF
                                        %>					
    
                                        <p class="riga"><span class="colonna articolo">[<%=rs("codicearticolo")%>]&nbsp;<%=rs("titolo")%><%if Len(rs("colore"))>0 then%>&nbsp;(<%=rs("colore")%>)<%end if%></span>
                                        <%
                                        quantita=rs("quantita")
                                        if quantita="" then quantita=1
                                        %>
                                        <span class="colonna quantita"><%=quantita%> pezzi </span><span class="colonna prezzo_unitario"><%=FormatNumber(rs("PrezzoProdotto"),2)%>&#8364;</span><span class="colonna prezzo_totale"><%=FormatNumber(rs("TotaleRiga"),2)%>&#8364;</span></p>
                                        <%
                                        rs.movenext
                                        loop
                                        %>
<%
	rs.close
%>                                        
                                    <%end if%>
                                </div>
                                
                                <p class="area clearfix"><span class="colonna descrizione">Modalit&agrave; di spedizione</span><span class="colonna prezzo_unitario">&nbsp;</span><span class="colonna prezzo_totale">Totale</span></p>
                                <div class="data">
                                    <p class="riga">
                                    <span class="colonna descrizione"><b><%=TipoTrasporto%></b></span>
                                    <span class="colonna prezzo_unitario">&nbsp;</span>
                                    <span class="colonna prezzo_totale"><%=FormatNumber(CostoSpedizioneTotale,2)%>&#8364;</span>
                                    </p>
                                    <p>&nbsp;</p>
                                    <h4>Riferimenti per l'indirizzo di spedizione</h4>
                                    <p><%=DatiSpedizione%></p>
                                    <p>&nbsp;</p>
                                    <h4>Eventuali annotazioni</h4>
                                    <p><%=NoteCliente%></p>
                                </div>
                                
                                <p class="area clearfix"><span class="colonna descrizione">Modalit&agrave; di pagamento - Descrizione</span><span class="colonna prezzo_unitario">&nbsp;</span><span class="colonna prezzo_totale">Totale</span></p>
                                <div class="data">
    
                                        <p class="riga">
                                        <span class="colonna descrizione"><b><%=TipoPagamento%></b></span>
                                        <span class="colonna prezzo_unitario">&nbsp;</span>
                                        <span class="colonna prezzo_totale"><%=FormatNumber(CostoPagamento,2)%>&#8364;</span>
                                        </p>
                                    <p>&nbsp;</p>
                                    <h4>Riferimenti per i dati di fatturazione:</h4>
                                    <div class="iscrizione clearfix">
                                    <div class="table">
                                    <div class="tr">
                                        <div class="td">
	                                        <%if Rag_Soc<>"" then%><%=Rag_Soc%>&nbsp;&nbsp;<%end if%><%if nominativo<>"" then%><%=nominativo%><%end if%>
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        	Codice fiscale: <%=Cod_Fisc%><%if PartitaIVA<>"" then%> - Partita IVA: <%=PartitaIVA%><%end if%>
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        	<%=indirizzo%>
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
	                                        <%=cap%> - <%=citta%> (<%=provincia%>)
                                        </div>
                                    </div>
                                </div>
                                </div> 
                                       
                                </div>
                                
                                  <h4 class="cart clearfix"><span class="total_price">Totale ordine:&nbsp;
                      <%if TotaleGenerale<>0 then%>
                      <%=FormatNumber(TotaleGenerale,2)%>
                      <%else%>
                      0,00
                      <%end if%>
                      &#8364;
                                  </span></h4>
                                    <form method="post" name="modulo" id="modulo" action="stampa_ordine.asp">
                                    <input type="hidden" name="idordine" id="idordine" value="<%=idordine%>">
                                    <input type="submit" name="stampa" value="stampa l'ordine" style="float:right;" class="button_link_red" />
                                
                                	</form>
                                
								
                            </div>
                        </div>
                    </div>
                </div>
                <!--#include file="inc_sx_prodotti.asp"-->
            </div>
        </div>
         <!--#include file="inc_footer.asp"-->
          <script src="js/init.js"></script>
    </body>
</html>
<!--#include file="inc_strClose.asp"-->