<!--#include file="inc_strConn.asp"-->
<%
	mode=request("mode")
	if mode="" then mode=0
	
	'se la session &eacute; gi&agrave; aperta sfrutto il pkid dell'ordine, altrimenti ne apro una
	IdOrdine=session("ordine_shop")
	if IdOrdine="" then IdOrdine=0
	if idOrdine=0 then response.redirect("carrello1.asp")
	
		'inserisco le eventuali note dal carrello1
	if fromURL="carrello1.asp" then
		Set os1 = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Ordini where PkId="&idOrdine
		os1.Open sql, conn, 3, 3
		os1("NoteCliente")=request("NoteCliente")
		os1.update
		os1.close
	end if
	if idsession=0 then response.Redirect("iscrizione.asp?prov=1")

	
	Call Visualizzazione("",0,"carrello2extra.asp")
	
	mode=request("mode")
	if mode="" then mode=0
	
	'inserisco il costo del trasporto. se nn ne &eacute; stato scelto uno, perch&eacute; sono appena entrato adesso in questa pagina, prendo il primo costo dal db
	
'	TipoTrasportoScelto=3

'	Set trasp_rs = Server.CreateObject("ADODB.Recordset")
'	if TipoTrasportoScelto=0 then
'		sql = "SELECT * FROM CostiTrasporto"
'	else
'		sql = "SELECT * FROM CostiTrasporto where PkId="&TipoTrasportoScelto
'	end if
'	trasp_rs.Open sql, conn, 1, 1
'	if trasp_rs.recordcount>0 then	
'		PkIdTrasportoScelto=trasp_rs("PkId")
'		NomeTrasportoScelto=trasp_rs("Nome")
'		CostoTrasportoScelto=trasp_rs("Costo")
'		TipoCostoTrasportoScelto=trasp_rs("TipoCosto")
'	end if
'	trasp_rs.close
	
	Destinazione=request("Destinazione")
	
	
	
	if Destinazione="" then
		Set cli_rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Clienti where pkid="&idsession
		cli_rs.Open sql, conn, 1, 1
		if cli_rs.recordcount>0 then
			Nominativo=cli_rs("Nome")&" "&cli_rs("Nominativo")
			Indirizzo=cli_rs("Indirizzo")
			CAP=cli_rs("CAP")
			Citta=cli_rs("Citta")
			Provincia=cli_rs("Provincia")
			nazionediversa=cli_rs("nazionediversa")
			Telefono=cli_rs("Telefono")
			
			Destinazione=Nominativo&" - "&Indirizzo&" - "&CAP&" "&Citta&" ("&Provincia&") "&nazionediversa&" - Telefono: "&Telefono&"" 
		end if
		cli_rs.close	
	end if
	
	Set os1 = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where PkId="&idOrdine
	os1.Open sql, conn, 3, 3
	
	TotaleCarrello=os1("TotaleCarrello")
	CostoSpedizione=os1("CostoSpedizione")
	
'	if TipoCostoTrasportoScelto=1 then
'		CostoSpedizione=CostoTrasportoScelto
'	end if
'	if TipoCostoTrasportoScelto=2 then
'		CostoSpedizione=(TotaleCarrello*CostoTrasportoScelto)/100
'	end if
'	if TipoCostoTrasportoScelto=3 or TotaleCarrello>=250 then
'		CostoSpedizione=0
'	end if
	
	os1("TipoTrasporto")="Corriere Int."
	os1("CostoSpedizione")=CostoSpedizione
	'TotaleGnerale_AG=TotaleCarrello+CostoSpedizione
	os1("TotaleGenerale")=TotaleCarrello+CostoSpedizione
	os1("FkCliente")=idsession
	stato_ordine=os1("stato")
	if stato_ordine="" then stato_ordine=0
	
	if mode=0 then
'		if italia_log=True then
'			os1("stato")=1
'		end if
'		if italia_log=False then
			if stato_ordine<3 then os1("stato")=12
'		end if
	else
		if mode=1 then os1("stato")=22
		os1("DatiSpedizione")=Destinazione
		NoteCliente=request("NoteCliente")
		os1("NoteCliente")=NoteCliente
	end if
	os1("DataAggiornamento")=now()
	os1("IpOrdine")=Request.ServerVariables("REMOTE_ADDR")
	os1.update
	
	os1.close
	
	if mode=1 then response.Redirect("carrello3.asp")
	if mode=2 then response.Redirect("calcolospedizione.asp")
%>
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Cristalensi - Carrello</title>
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="/js/media-queries-ie.js"></script>
        <![endif]-->
        <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
        <script src="/js/jquery.blueberry.js"></script>
        <script src="/js/jquery.tipTip.js"></script>
        <link href="/css/css.css" rel="stylesheet" type="text/css">
        <link href="/css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="/css/tipTip.css" rel="stylesheet" type="text/css">
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
        <script language="javascript">
		function Continua()
		{
				document.modulocarrello.method = "post";
				document.modulocarrello.action = "/carrello2extra.asp?mode=1";
				document.modulocarrello.submit();
		}
		</script>
		<script language="javascript">
		function CalcoloSpedizione()
		{
				document.modulocarrello.method = "post";
				document.modulocarrello.action = "/carrello2extra.asp?mode=2";
				document.modulocarrello.submit();
		}
		</script>
        <!--Codice Statistiche Google Analytics Iury Mazzoni ## NON CANCELLARE!! ## -->
		<script type="text/javascript">
        
          var _gaq = _gaq || [];
          _gaq.push(['_setAccount', 'UA-320952-2']);
          _gaq.push(['_trackPageview']);
        
          (function() {
            var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
            ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
            var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
          })();
        
        </script>
        <!--Codice Statistiche Google Analytics Iury Mazzoni ## NON CANCELLARE!! ## -->
    </head>
    <body>
        <div id="wrap">
            <!--#include file="inc_header.asp"-->

            <div id="main-content">
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT PkId, FkOrdine, FkProdotto, PrezzoProdotto, Quantita, TotaleRiga, Titolo, CodiceArticolo, Colore, Lampadina FROM RigheOrdine WHERE FkOrdine="&idOrdine&""
	rs.Open sql, conn, 1, 1
	num_prodotti_carrello=rs.recordcount
	
	Set ss = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where pkid="&idOrdine
	ss.Open sql, conn, 1, 1
	
	if ss.recordcount>0 then
		TotaleCarrello=ss("TotaleCarrello")
		CostoSpedizioneTotale=ss("CostoSpedizione")
		if CostoSpedizioneTotale="" or isnull(CostoSpedizioneTotale) then CostoSpedizioneTotale=0
		TotaleGenerale=ss("TotaleGenerale")
		NoteCliente=ss("NoteCliente")	
	end if
%>               
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <h3 style="font-size: 14px; display: inline; border: none;">Il tuo ordine: modalit&agrave; di spedizione/ritiro prodotti</h3>
                            <div class="carrello clearfix">
                                <form name="modulocarrello" id="modulocarrello">
                                <p class="area clearfix"><span class="colonna articolo">[Codice articolo] Nome prodotto</span><span class="colonna quantita">quantit&agrave;</span><span class="colonna prezzo_unitario">prezzo unitario</span><span class="colonna prezzo_totale">prezzo totale</span></p>
                                <div class="data">
                                    <%if rs.recordcount>0 then%>
                                        
                                        <%
                                        Do while not rs.EOF
                                        %>					
    
                                        <p class="riga"><span class="colonna articolo">[<%=rs("codicearticolo")%>]&nbsp;<strong><%=rs("titolo")%></strong><%if Len(rs("colore"))>0 or Len(rs("lampadina"))>0 then%><br /><%if Len(rs("colore"))>0 then%>&nbsp;Col.:&nbsp;<%=rs("colore")%><%end if%><%if Len(rs("lampadina"))>0 then%>&nbsp;-&nbsp;Lamp.:&nbsp;<%=rs("lampadina")%><%end if%><%end if%></span>
                                        <%
                                        quantita=rs("quantita")
                                        if quantita="" then quantita=1
                                        %>
                                        <span class="colonna quantita"><%=quantita%> pezzi </span><span class="colonna prezzo_unitario"><%=FormatNumber(rs("PrezzoProdotto"),2)%>&#8364;</span><span class="colonna prezzo_totale"><%=FormatNumber(rs("TotaleRiga"),2)%>&#8364;</span></p>
                                        <%
                                        rs.movenext
                                        loop
                                        %>
                                        
									<%else%>
                                    	<p class="riga">Il carrello &eacute; vuoto</p>
                                    <%end if%>
                                    <p>&nbsp;</p>
                                    <h3 style="font-size:12px;">Eventuali annotazioni</h3>
                                    <p>Potete usare questo spazio per inserie eventuali annotazioni o comunicazioni in relazione ai prodotti acquistati.</p>
                                    <textarea name="NoteCliente" cols="100" rows="2" id="NoteCliente" style="margin-left:20px;"><%=NoteCliente%></textarea> 
                                       
                                </div>

                                <p class="area clearfix"><span class="colonna descrizione">Modalit&agrave; di spedizione - Descrizione</span><span class="colonna prezzo_unitario">Costo</span><span class="colonna prezzo_totale">Totale</span></p>
                                <div class="data">
                                        <%if stato_ordine=22 then%>					
                                        <p class="riga">
                                        <span class="colonna descrizione"><b>Corriere internazionale</b></span>
                                        <span class="colonna prezzo_unitario"><%=FormatNumber(CostoSpedizioneTotale,2)%>&#8364;</span>
                                        <span class="colonna prezzo_totale"><%=FormatNumber(CostoSpedizioneTotale,2)%>&#8364;</span>
                                        </p>
                                        <%else%>
                                        <p class="riga">
                                        <span class="colonna descrizione"><b>Corriere internazionale</b><br />Per le spedizioni internazionali &eacute; necessario calcolare il peso dei prodotti ordinati per sapere il costo della spedizione.<br />
			      Per chi desidera continuare l'acquisto deve seguire la seguente procedura:<br />
			      -assicurarsi che i prodotti inseriti nel carrello siano quelli desiderati,<br />
			      -cliccare sul pulsante &quot;clicca qui per il calcolo del costo di spedizione&quot;.<br />
			      Nelle successive 24h (ma potrebbe passare anche solo 1h) ricever&agrave; un'email che le comunicher&agrave; il costo di spedizione e la possibilit&agrave; di continuare l'acquisto. <br />
			      Una volta ricevuta la comunicazione dall'amministratore, &egrave; necessario tornare sul sito internet, e dalla Home Page farsi riconoscere inserendo Login (Email) e Password inserite durante l'iscrizione. A quel punto, sempre nell'Area clienti, ci sar&agrave; un link con &quot;I miei ordini&quot;, cliccandoci andrete nell'elenco dei vostri ordini e potrete continuare l'acquito, trovandoci gi&agrave; calcolato il Costo di spedizione.</span>
                                        <span class="colonna prezzo_unitario">&nbsp;</span>
                                        <span class="colonna prezzo_totale">&nbsp;</span>
                                        </p>
                                        <%end if%>
                                    <p>&nbsp;</p>
                                    <h3 style="font-size:12px;">Riferimenti per l'indirizzo di spedizione</h3>
                                    <p>E' possibile  indicare anche un indirizzo diverso da quello indicato (i dati riportati sono gli stessi indicati al momento dell'iscrizione)</p>
                                    <textarea name="Destinazione" cols="100" rows="2" id="Destinazione" style="margin-left:20px;"><%=Destinazione%></textarea> 
                                       
                                </div>
                                
                                <%if ss.recordcount>0 then%>
                                  <h4 class="cart clearfix"><span class="total_price">Totale carrello: 
								  <%if ss("TotaleGenerale")<>0 then%>
								  <%=FormatNumber(ss("TotaleGenerale"),2)%>
                                  <%else%>
                                  0,00
                                  <%end if%>
                                  &#8364;&nbsp;
                                  </span></h4>
									<%if rs.recordcount>0 then%>
                                    
                                    <p><button type="button" name="indietro" onClick="location.href='/carrello1.asp'" style="float:left;" class="button_link">&laquo; passo precedente</button>
                                    <button type="button" name="continua" onClick="<%if stato_ordine=22 then%>Continua();<%else%>CalcoloSpedizione();<%end if%>" style="float:right;" class="button_link_red"><%if stato_ordine=22 then%>clicca qui per continuare l'acquisto<%else%>clicca qui per il calcolo del costo di spedizione<%end if%> &raquo;</button></p>
                                    <%end if%>
								<%end if%>
                                
                                </form>
                                
								<%
                                ss.close
                                rs.close
                                %>
                            </div>
                        </div>
                    </div>
                </div>
                <!--#include file="inc_sx_prodotti.asp"-->
            </div>
        </div>
         <!--#include file="inc_footer.asp"-->
    </body>
</html>
<!--#include file="inc_strClose.asp"-->