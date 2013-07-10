<!--#include file="inc_strConn.asp"-->
<%
if idsession=0 then response.Redirect("iscrizione.asp?prov=2")

mode=request("mode")
if mode="" then mode=0

if mode=1 or mode=2 then 
	IdOrdine=request("IdOrdine")
	
	if IdOrdine>0 then
		session("ordine_shop")=IdOrdine
		if mode=1 then response.Redirect("carrello1.asp")
		if mode=2 then response.Redirect("carrello2extra.asp")
	end if
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
        <script language="JavaScript" type="text/JavaScript">
		<!--
		function MM_openBrWindow(theURL,winName,features) { //v2.0
		  window.open(theURL,winName,features);
		}
		//-->
		</script>
    </head>
    <body>
        <div id="wrap">
            <!--#include file="inc_header.asp"-->

            <div id="main-content">
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini WHERE FkCliente="&idsession&""
	rs.Open sql, conn, 1, 1
%>                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <h3 style="font-size: 14px; display: inline; border: none;">I tuoi ordini:</h3>
                            <div class="carrello clearfix">
                                <p class="area clearfix"><span class="colonna num_ordine">[codice ordine] - Data</span><span class="colonna totale_ordine">totale</span><span class="colonna stato">stato/informazioni</span><span class="colonna azioni">azioni</span></p>
                                <div class="data">
                                    <%
									if rs.recordcount>0 then
										Do while not rs.EOF
										
										InfoSpedizione=rs("InfoSpedizione")
										NoteCri=rs("NoteCri")
										stato=rs("Stato")
										
										if stato=0 then etichetta_stato="Carrello iniziato"
										if stato=1 then etichetta_stato="Carrello iniziato"
										if stato=2 then etichetta_stato="Spedizione scelta"
										if stato=12 then etichetta_stato="Spedizione scelta"
										if stato=3 then etichetta_stato="Pagamento da scegliere"
										if stato=22 then etichetta_stato="Pagamento da scegliere"
										
										if stato=4 then etichetta_stato="Pagato con PayPal"
										if stato=5 then etichetta_stato="Pagamento annullato"
										if stato=6 then etichetta_stato="In fase di pagamento"
										if stato=7 then etichetta_stato="Ordine in lavorazione"
										if stato=8 then
											etichetta_stato="Prodotti spediti"
											if InfoSpedizione<>"" then etichetta_stato=etichetta_stato&"<br>"&InfoSpedizione
											if Left(NoteCri,4)="http" then etichetta_stato=etichetta_stato&"<br><a href="""&NoteCri&""" target=_blank>LINK</a>"
										end if
										%>	
    
                                        <p class="riga">
                                        <span class="colonna num_ordine">[<%=rs("PkId")%>]&nbsp;-&nbsp;<%=rs("DataAggiornamento")%></span>
                                        <span class="colonna totale_ordine">
                                          <%
										  TotaleGenerale=rs("TotaleGenerale")
										  if TotaleGenerale="" or Isnull(TotaleGenerale) then TotaleGenerale=0
										  %>
										  <%=FormatNumber(TotaleGenerale,2)%>€
                                        </span>
                                        <span class="colonna stato"><%=etichetta_stato%></span>
                                        <span class="colonna azioni">
                                          <button type="button" name="visualizza" class="button_link" onClick="MM_openBrWindow('stampa_ordine.asp?idordine=<%=rs("PkId")%>&mode=0','','width=760,height=400,scrollbars=yes')">visualizza</button>
										  <%if stato=0 or stato=1 or stato=2 or stato=3 or stato=6 then%>
                                           &nbsp;<button type="button" name="modifica" class="button_link" onClick="document.location.href='ordini_elenco.asp?IdOrdine=<%=rs("PkId")%>&mode=1';">continua l'ordine</button>
                                          <%else%>
                                            <%if stato=12 or stato=22 then%>
                                            <a href="ordini_elenco.asp?IdOrdine=<%=rs("PkId")%>&mode=2"><b>[<%=rs("PkId")%>]&nbsp;-&nbsp;<%=rs("DataAggiornamento")%></b></a>
                                            &nbsp;<button type="button" name="modifica" class="button_link" onClick="document.location.href='ordini_elenco.asp?IdOrdine=<%=rs("PkId")%>&mode=2';">continua l'ordine</button>
                                            <%end if%>
                                          <%end if%>
                                            &nbsp;<button type="button" name="stampa" class="button_link" onClick="MM_openBrWindow('stampa_ordine.asp?idordine=<%=rs("PkId")%>&mode=1','','width=760,height=900,scrollbars=yes')">stampa</button>
                                        </span>
                                        </p>
                                        <%
                                        rs.movenext
                                        loop
                                        %>
                                        
									<%else%>
                                    	<p class="riga">Non sono presenti ordini</p>
                                    <%end if%>    
                                </div>
								<%
                                rs.close
                                %>
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