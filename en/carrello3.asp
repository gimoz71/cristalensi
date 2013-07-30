<!--#include file="inc_strConn.asp"-->
<%
	Call Visualizzazione("",0,"carrello3.asp")
	
	mode=request("mode")
	if mode="" then mode=0
	
	'se la session è già aperta sfrutto il pkid dell'ordine, altrimenti ne apro una
	IdOrdine=session("ordine_shop")
	if IdOrdine="" then IdOrdine=0
	if idOrdine=0 then response.redirect("carrello1.asp")
	
	if idsession=0 then response.Redirect("iscrizione.asp?prov=1")
	
	'inserisco il costo del pagamento. se nn ne è stato scelto uno, perchè sono appena entrato adesso in questa pagina, prendo il primo costo dal db
	
	TipoPagamentoScelto=request("TipoPagamentoScelto")
	if TipoPagamentoScelto="" then TipoPagamentoScelto=0

	Set trasp_rs = Server.CreateObject("ADODB.Recordset")
	if TipoPagamentoScelto=0 then
		sql = "SELECT * FROM CostiPagamento"
	else
		sql = "SELECT * FROM CostiPagamento where PkId="&TipoPagamentoScelto
	end if
	trasp_rs.Open sql, conn, 1, 1
	if trasp_rs.recordcount>0 then	
		PkIdPagamentoScelto=trasp_rs("PkId")
		NomePagamentoScelto=trasp_rs("Nome")
		CostoPagamentoScelto=trasp_rs("Costo")
		TipoCostoPagamentoScelto=trasp_rs("TipoCosto")
	end if
	trasp_rs.close
	
	Nominativo=request("Nominativo")
	Rag_Soc=request("Rag_Soc")
	
	if Nominativo="" and Rag_Soc="" then
		Set cli_rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Clienti where pkid="&idsession
		cli_rs.Open sql, conn, 1, 1
		if cli_rs.recordcount>0 then
			Nominativo=cli_rs("Nome")&" "&cli_rs("Nominativo")
			Rag_Soc=cli_rs("Rag_Soc")
			Cod_Fisc=cli_rs("Cod_Fisc")
			PartitaIVA=cli_rs("PartitaIVA")
			Indirizzo=cli_rs("Indirizzo")
			CAP=cli_rs("CAP")
			Citta=cli_rs("Citta")
			Provincia=cli_rs("Provincia")
		end if
		cli_rs.close	
	else
		Cod_Fisc=request("Cod_Fisc")
		PartitaIVA=request("PartitaIVA")
		Indirizzo=request("Indirizzo")
		CAP=request("CAP")
		Citta=request("Citta")
		Provincia=request("Provincia")
	end if
	
	Set os1 = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where PkId="&idOrdine
	os1.Open sql, conn, 3, 3
	
	TotaleCarrello=os1("TotaleCarrello")
	CostoSpedizione=os1("CostoSpedizione")
	
	if TipoCostoPagamentoScelto=1 then
		CostoPagamento=CostoPagamentoScelto
	end if
	if TipoCostoPagamentoScelto=2 then
		CostoPagamento=((TotaleCarrello+CostoSpedizione)*CostoPagamentoScelto)/100
	end if
	if TipoCostoPagamentoScelto=3 then
		CostoPagamento=0
	end if
	
	os1("FkPagamento")=PkIdPagamentoScelto
	os1("TipoPagamento")=NomePagamentoScelto
	os1("CostoPagamento")=CostoPagamento
	'TotaleGnerale_AG=TotaleCarrello+CostoSpedizione+CostoPagamento
	os1("TotaleGenerale")=TotaleCarrello+CostoSpedizione+CostoPagamento
	os1("FkCliente")=idsession
	
	if mode=0 then
		os1("stato")=2
		italia_log=session("italia_log")
		if italia_log="No" then os1("stato")=22
	else
		os1("stato")=3
		os1("Nominativo")=Nominativo
		os1("Rag_Soc")=Rag_Soc
		os1("Cod_Fisc")=Cod_Fisc
		os1("PartitaIVA")=PartitaIVA
		os1("Indirizzo")=Indirizzo
		os1("CAP")=CAP
		os1("Citta")=Citta
		os1("Provincia")=Provincia
	end if
	os1("DataAggiornamento")=now()
	os1("IpOrdine")=Request.ServerVariables("REMOTE_ADDR")
	os1.update
	
	os1.close
	
	if mode=1 then response.Redirect("ordine.asp")
%>
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Cristalensi</title>
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
		function Cambia()
		{
				document.modulocarrello.method = "post";
				document.modulocarrello.action = "carrello3.asp";
				document.modulocarrello.submit();
		}
		</script>
		<script language="javascript">
		function Continua()
		{
				document.modulocarrello.method = "post";
				document.modulocarrello.action = "carrello3.asp?mode=1";
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
	sql = "SELECT PkId, FkOrdine, FkProdotto, PrezzoProdotto, Quantita, TotaleRiga, Titolo, CodiceArticolo, Colore FROM RigheOrdine WHERE FkOrdine="&idOrdine&""
	rs.Open sql, conn, 1, 1
	num_prodotti_carrello=rs.recordcount
	
	Set ss = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where pkid="&idOrdine
	ss.Open sql, conn, 1, 1
	
	if ss.recordcount>0 then
		TotaleCarrello=ss("TotaleCarrello")
		CostoSpedizioneTotale=ss("CostoSpedizione")
		TipoTrasporto=ss("TipoTrasporto")
		DatiSpedizione=ss("DatiSpedizione")
		CostoPagamentoTotale=ss("CostoPagamento")
		TotaleGenerale=ss("TotaleGenerale")
		NoteCliente=ss("NoteCliente")	
	end if
%>               
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <h3 style="font-size: 14px; display: inline; border: none;">Il tuo ordine: modalit&agrave; di pagamento</h3>
                            <div class="carrello clearfix">
                                <form name="modulocarrello" id="modulocarrello">
                                <p class="area clearfix"><span class="colonna articolo">[Codice articolo] Nome prodotto</span><span class="colonna quantita">quantità</span><span class="colonna prezzo_unitario">prezzo unitario</span><span class="colonna prezzo_totale">prezzo totale</span></p>
                                <div class="data">
                                    <%if rs.recordcount>0 then%>
                                        
                                        <%
                                        Do while not rs.EOF
                                        %>					
    
                                        <p class="riga"><span class="colonna articolo">[<%=rs("codicearticolo")%>]&nbsp;<%=rs("titolo")%><%if Len(rs("colore"))>0 then%>&nbsp;(<%=rs("colore")%>)<%end if%></span>
                                        <%
                                        quantita=rs("quantita")
                                        if quantita="" then quantita=1
                                        %>
                                        <span class="colonna quantita"><%=quantita%> pezzi </span><span class="colonna prezzo_unitario"><%=FormatNumber(rs("PrezzoProdotto"),2)%>€</span><span class="colonna prezzo_totale"><%=FormatNumber(rs("TotaleRiga"),2)%>€</span></p>
                                        <%
                                        rs.movenext
                                        loop
                                        %>
                                        
									<%else%>
                                    	<p class="riga">Il carrello è vuoto</p>
                                    <%end if%>
                                </div>
                                
                                <p class="area clearfix"><span class="colonna descrizione">Modalit&agrave; di spedizione</span><span class="colonna prezzo_unitario">&nbsp;</span><span class="colonna prezzo_totale">Totale</span></p>
                                <div class="data">
                                    <p class="riga">
                                    <span class="colonna descrizione"><b><%=TipoTrasporto%></b></span>
                                    <span class="colonna prezzo_unitario">&nbsp;</span>
                                    <span class="colonna prezzo_totale"><%=FormatNumber(CostoSpedizioneTotale,2)%>€</span>
                                    </p>
                                    <p>&nbsp;</p>
                                    <h4>Riferimenti per l'indirizzo di spedizione</h4>
                                    <p><%=DatiSpedizione%></p>
                                    <p>&nbsp;</p>
                                    <h4>Colori misure e annotazioni</h4>
                                    <p><%=NoteCliente%></p>
                                </div>
                                
                                <%
								Set trasp_rs = Server.CreateObject("ADODB.Recordset")
								if session("italia_log")="No" then
									sql = "SELECT Top 2 * FROM CostiPagamento"
								else
									sql = "SELECT * FROM CostiPagamento"
								end if
								
								trasp_rs.Open sql, conn, 1, 1
								if trasp_rs.recordcount>0 then
								%>
                                <p class="area clearfix"><span class="colonna descrizione">Modalit&agrave; di pagamento - Descrizione</span><span class="colonna prezzo_unitario">Costo</span><span class="colonna prezzo_totale">Totale</span></p>
                                <div class="data">
                                        <%
										Do while not trasp_rs.EOF
										PkIdPagamento=trasp_rs("pkid")
										NomePagamento=trasp_rs("nome")
										DescrizionePagamento=trasp_rs("descrizione")
										CostoPagamento=trasp_rs("costo")
										
										TipoCosto=trasp_rs("TipoCosto")
										if TipoCosto="" then TipoCosto=3
                                        %>					
    
                                        <p class="riga">
                                        <span class="colonna descrizione"><input type="radio" name="TipoPagamentoScelto" id="TipoPagamentoScelto" value="<%=PkIdPagamento%>" <%if PkIdPagamento=PkIdPagamentoScelto then%> checked="checked"<%end if%> onClick="Cambia();">&nbsp;<b><%=NomePagamento%></b><br><%=NoLettAcc(DescrizionePagamento)%></span>
                                        <span class="colonna prezzo_unitario"><%=FormatNumber(CostoPagamento,2)%><%if TipoCosto=1 then%>€<%end if%><%if TipoCosto=2 then%>%<%end if%></span>
                                        <span class="colonna prezzo_totale"><%if PkIdPagamento=PkIdPagamentoScelto then%><%=FormatNumber(CostoPagamentoTotale,2)%>€<%else%>-<%end if%></span>
                                        </p>
                                        <%
                                        trasp_rs.movenext
                                        loop
                                        %>
                                    <p>&nbsp;</p>
                                    <h4>Riferimenti per i dati di fatturazione:</h4>
                                    <p>&egrave; possibile  indicare dati diversi da quelli indicati (i dati riportati sono gli stessi indicati al momento dell'iscrizione).<br>La fattura, alle aziende che espressamente la richiedono, è emessa per ordini superiori a 150€.</p>
                                    
                                    <div class="iscrizione clearfix">
                                    <div class="table">
                                    <div class="tr">
                                        <div class="td">
	                                        Nome e Cognome (*)<br />
                                            <input name="nominativo" type="text" class="form" id="nominativo"  size="30" maxlength="50" value="<%=nominativo%>" />
                                        </div>
                                        <div class="td">
                                        	Ragione sociale ( nel caso in cui si tratti di un'Azienda )<br />
                                            <input name="Rag_Soc" type="text" class="form" id="Rag_Soc"  size="30" maxlength="50" value="<%=Rag_Soc%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        Codice Fiscale<br />
                                            <input name="cod_fisc" type="text" class="form" id="cod_fisc"  size="20" maxlength="16" value="<%=cod_fisc%>" />
                                        </div>
                                        <div class="td">
                                        Partita IVA ( nel caso in cui si tratti di un'Azienda )<br />
                                            <input name="PartitaIVA" type="text" class="form" id="PartitaIVA"  size="20" maxlength="11" value="<%=PartitaIVA%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        	Indirizzo (*)<br />
                                            <input name="indirizzo" type="text" class="form" id="indirizzo"  size="30" maxlength="100" value="<%=indirizzo%>" />
                                        </div>
                                        <div class="td">
                                        	CAP<br />
                                            <input name="cap" type="text" class="form" id="cap"  size="7" maxlength="5" value="<%=cap%>" />
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
	                                        Citt&agrave; (*)<br />
                                            <input name="citta" type="text" class="form" id="citta"  size="30" maxlength="50" value="<%=citta%>" />
                                        </div>
                                        <div class="td">
	                                        Provincia<br />
                                            <input type="text" name="provincia" id="provincia" value="<%=provincia%>" size="3" maxlength="2" class="form" />
                                        </div>
                                    </div>
                                </div>
                                </div> 
                                       
                                </div>
                                <%end if%>
                                <%trasp_rs.close%>
                                
                                <%if ss.recordcount>0 then%>
                                  <h4 class="cart clearfix"><span class="total_price">Totale carrello: 
								  <%if ss("TotaleGenerale")<>0 then%>
								  <%=FormatNumber(ss("TotaleGenerale"),2)%>
                                  <%else%>
                                  0,00
                                  <%end if%>
                                  €&nbsp;
                                  </span></h4>
									<%if rs.recordcount>0 then%>
                                    
                                    <p><button type="button" name="indietro" onClick="location.href='carrello2.asp'" style="float:left;" class="button_link">&laquo; passo precedente</button>
                                    <button type="button" name="continua" onClick="Continua();" style="float:right;" class="button_link_red">CONCLUDI L'ACQUISTO &raquo;</button></p>
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