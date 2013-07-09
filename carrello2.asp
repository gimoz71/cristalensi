<!--#include file="inc_strConn.asp"-->
<%
	mode=request("mode")
	if mode="" then mode=0
	
	'se la session è già aperta sfrutto il pkid dell'ordine, altrimenti ne apro una
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
	
	italia_log=session("italia_log")
	if italia_log="No" then response.Redirect("carrello2extra.asp") 
	
	Call Visualizzazione("",0,"carrello2.asp")
	
	mode=request("mode")
	if mode="" then mode=0
	
	'inserisco il costo del trasporto. se nn ne è stato scelto uno, perchè sono appena entrato adesso in questa pagina, prendo il primo costo dal db
	
	TipoTrasportoScelto=request("TipoTrasportoScelto")
	if TipoTrasportoScelto="" then TipoTrasportoScelto=0

	Set trasp_rs = Server.CreateObject("ADODB.Recordset")
	if TipoTrasportoScelto=0 then
		sql = "SELECT * FROM CostiTrasporto"
	else
		sql = "SELECT * FROM CostiTrasporto where PkId="&TipoTrasportoScelto
	end if
	trasp_rs.Open sql, conn, 1, 1
	if trasp_rs.recordcount>0 then	
		PkIdTrasportoScelto=trasp_rs("PkId")
		NomeTrasportoScelto=trasp_rs("Nome")
		CostoTrasportoScelto=trasp_rs("Costo")
		TipoCostoTrasportoScelto=trasp_rs("TipoCosto")
	end if
	trasp_rs.close
	
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
			Telefono=cli_rs("Telefono")
			
			Destinazione=Nominativo&" - "&Indirizzo&" - "&CAP&" "&Citta&" ("&Provincia&") - Telefono: "&Telefono&"" 
		end if
		cli_rs.close	
	end if
	
	Set os1 = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where PkId="&idOrdine
	os1.Open sql, conn, 3, 3
	
	TotaleCarrello=os1("TotaleCarrello")
	
	if TipoCostoTrasportoScelto=1 then
		CostoSpedizione=CostoTrasportoScelto
	end if
	if TipoCostoTrasportoScelto=2 then
		CostoSpedizione=(TotaleCarrello*CostoTrasportoScelto)/100
	end if
	if TipoCostoTrasportoScelto=3 or TotaleCarrello>=250 then
		CostoSpedizione=0
	end if
	
	os1("TipoTrasporto")=NomeTrasportoScelto
	os1("CostoSpedizione")=CostoSpedizione
	'TotaleGnerale_AG=TotaleCarrello+CostoSpedizione
	os1("TotaleGenerale")=TotaleCarrello+CostoSpedizione
	os1("FkCliente")=idsession
	if mode=0 then
		os1("stato")=1
	else
		os1("stato")=2
		os1("DatiSpedizione")=Destinazione
		NoteCliente=request("NoteCliente")
		os1("NoteCliente")=NoteCliente
	end if
	os1("DataAggiornamento")=now()
	os1("IpOrdine")=Request.ServerVariables("REMOTE_ADDR")
	os1.update
	
	os1.close
	
	if mode=1 then response.Redirect("carrello3.asp")
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
        <script language="javascript">
		function Cambia()
		{
				document.modulocarrello.method = "post";
				document.modulocarrello.action = "carrello2.asp";
				document.modulocarrello.submit();
		}
		</script>
		<script language="javascript">
		function Continua()
		{
				document.modulocarrello.method = "post";
				document.modulocarrello.action = "carrello2.asp?mode=1";
				document.modulocarrello.submit();
		}
		</script>
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
                                <p class="area clearfix"><span class="colonna articolo">[Codice articolo] Nome prodotto</span><span class="colonna quantita">quantità</span><span class="colonna prezzo_unitario">prezzo unitario</span><span class="colonna prezzo_totale">prezzo totale</span></p>
                                <div class="data">
                                    <%if rs.recordcount>0 then%>
                                        
                                        <%
                                        Do while not rs.EOF
                                        %>					
    
                                        <p class="riga"><a href="scheda_prodotto.asp?id=<%=rs("FkProdotto")%>" class="colonna articolo">[<%=rs("codicearticolo")%>]&nbsp;<%=rs("titolo")%><%if Len(rs("colore"))>0 then%>&nbsp;(<%=rs("colore")%>)<%end if%></a>
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
                                    <p>&nbsp;</p>
                                    <h4>Colori misure e annotazioni</h4>
                                    <p>Scrivere in questo spazio il colore e la misura dei prodotti nel caso in cui fossero presenti più varianti degli stessi.<br>Oppure potete usare questo spazio per inserie qualche annotazione o comunicazione.</p>
                                    <textarea name="NoteCliente" cols="100" rows="2" id="NoteCliente" style="margin-left:20px;"><%=NoteCliente%></textarea> 
                                       
                                </div>
                                <%
								Set trasp_rs = Server.CreateObject("ADODB.Recordset")
								sql = "SELECT * FROM CostiTrasporto"
								trasp_rs.Open sql, conn, 1, 1
								%>
                                <p class="area clearfix"><span class="colonna descrizione">Modalit&agrave; di spedizione - Descrizione</span><span class="colonna prezzo_unitario">Costo</span><span class="colonna prezzo_totale">Totale</span></p>
                                <div class="data">
                                        <%
                                        if trasp_rs.recordcount>0 then
										Do while not trasp_rs.EOF
										PkIdSpedizione=trasp_rs("pkid")
										NomeSpedizione=trasp_rs("nome")
										DescrizioneSpedizione=trasp_rs("descrizione")
										CostoSpedizione=trasp_rs("costo")
										
										TipoCosto=trasp_rs("TipoCosto")
										if TipoCosto="" then TipoCosto=3
                                        %>					
    
                                        <p class="riga">
                                        <span class="colonna descrizione"><input type="radio" name="TipoTrasportoScelto" id="TipoTrasportoScelto" value="<%=PkIdSpedizione%>" <%if PkIdSpedizione=PkIdTrasportoScelto then%> checked="checked"<%end if%> onClick="Cambia();">&nbsp;<b><%=NomeSpedizione%></b><br><%=NoLettAcc(DescrizioneSpedizione)%></span>
                                        <span class="colonna prezzo_unitario"><%=FormatNumber(CostoSpedizione,2)%><%if TipoCosto=1 then%>€<%end if%><%if TipoCosto=2 then%>%<%end if%></span>
                                        <span class="colonna prezzo_totale"><%if PkIdSpedizione=PkIdTrasportoScelto then%><%=FormatNumber(CostoSpedizioneTotale,2)%>€<%else%>-<%end if%></span>
                                        </p>
                                        <%
                                        trasp_rs.movenext
                                        loop
										end if
                                        %>
                                    <p>&nbsp;</p>
                                    <h4>Riferimenti per l'indirizzo di spedizione</h4>
                                    <p>E' possibile  indicare anche un indirizzo diverso da quello indicato (i dati riportati sono gli stessi indicati al momento dell'iscrizione)</p>
                                    <textarea name="Destinazione" cols="100" rows="2" id="Destinazione" style="margin-left:20px;"><%=Destinazione%></textarea> 
                                       
                                </div>
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
                                    
                                    <input type="button" name="continua" value="&laquo; passo precedente" onClick="location.href='carrello1.asp'" style="float:left;" class="button_link" />
                                    <input type="button" name="continua" value="clicca qui per continuare l'acquisto &raquo;" onClick="Continua();" style="float:right;" class="button_link_red" />
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
          <script src="js/init.js"></script>
    </body>
</html>
<!--#include file="inc_strClose.asp"-->