<!--#include file="inc_strConn.asp"-->
<%
	mode=request("mode")
	if mode="" then mode=0
		
	if idsession=0 then response.Redirect("iscrizione.asp?prov=2")
		
	'inserisco il costo del trasporto. se nn ne è stato scelto uno, perchè sono appena entrato adesso in questa pagina, prendo il primo costo dal db
	
	Destinazione=request("Destinazione")
	
	if mode=1 then
		Set cli_rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Commenti_Clienti"
		cli_rs.Open sql, conn, 3, 3
		cli_rs.addnew
			cli_rs("Testo")=request("Testo")
			cli_rs("FkIscritto")=idsession
			cli_rs("Data")=now()
			cli_rs("Pubblicato")=False
			cli_rs("Risposta")=False
		cli_rs.update
		cli_rs.close	
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
                            <h3 style="font-size: 14px; display: inline; border: none;">Inserisci il tuo commento!</h3>
                            <div class="carrello clearfix">
                                 <%if mode=1 then%>
                                 	<p>
                                    Il tuo commento è stato inserito correttamente, adesso il nostro staff lo valuterà e se sarà approvato, ti verrà recapitata una notifica via email.<br />Grazie della tua collaborazione dallo staff di Cristalensi.<br /><br /><a href="commenti_elenco.asp" class="button_link_red" style="float:right">Elenco commenti</a>
                                    </p>
								 <%else%>   
                                    <form name="modulocarrello" id="modulocarrello" method="post" action="commenti_form.asp?mode=1">
                                    <p>Inserisci un commento su i prodotti acquistati, se ti sono piaciuti o no, oppure un commento sul sito internet o sull'azienda e lo staff.<br />Il commento non sarà pubblicato immediatamente ma sarà soggetto a un controllo da parte del nostro staff per evitare che vengano inseriti contenuti non leciti, offese e termini non pubblicabili.<br />Si prega di non inserire codice html, email, link e collegamenti ad altri siti internet: il commento non sarà pubblicato.<br />Per ogni commento saranno pubblicati anche il <strong>Nome</strong> e la <strong>Città</strong> inseriti al momento dell'iscrizione.</p>
                                    <textarea name="testo" cols="105" rows="5" id="testo"></textarea>
                                    <p>
                                    <input type="button" name="reset" value="&laquo; elenco commenti" class="button_link" style="float:left;" onClick="location.href='commenti_elenco.asp'">
                                    <input type="submit" name="continua" value="clicca qui per inserire il tuo commento &raquo;" class="button_link_red" style="float:right">
                                    </p>
                                    </form>
								<%end if%>
                            </div>
                        </div>
                    </div>
                </div>
                <!--#include file="inc_sx.asp"-->
            </div>
        </div>
         <!--#include file="inc_footer.asp"-->
          <script src="js/init.js"></script>
    </body>
</html>
<!--#include file="inc_strClose.asp"-->