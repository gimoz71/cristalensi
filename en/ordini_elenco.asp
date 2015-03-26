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
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Cristalensi - Orders</title>
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="/js/media-queries-ie.js"></script>
        <![endif]-->
				<link href="/css/css.css" rel="stylesheet" type="text/css">
        <link href="/css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="/css/tipTip.css" rel="stylesheet" type="text/css">
        <script src="http://code.jquery.com/jquery-1.11.2.min.js"></script>
        <script src="/js/jquery.blueberry-min.js"></script>
        <script src="/js/jquery.tipTip-min.js"></script>
        <style type="text/css">
            .clearfix:after {
                content: ".";
                display: block;
                height: 0;
                clear: both;
                visibility: hidden;
            }
        </style>
        <!--[if lt IE 9]>
            <style>
                #menu, #language {
                    display: block !important;

                }
                #language li {
                    display: inline-block !important;
                    float: left !important;
                    text-align: center !important;
                    padding: 6px 17px !important;
                    height: auto !important;

                }
                #menu li {
                    display: inline-block !important;
                    float: left !important;
                    text-align: center !important;
                    padding: 11px 17px !important;
                    height: auto !important;

                }
                ul.slides {height: 170px !important}
                .button_link {
                    background: #999 !important;
                }
                .button_link_red {
                    background: #c00 !important;
                }
            </style>
        <![endif]-->
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
	sql = "SELECT * FROM Ordini WHERE FkCliente="&idsession&""
	rs.Open sql, conn, 1, 1
%>                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <h3 style="font-size: 14px; display: inline; border: none;">Your orders:</h3>
                            <div class="carrello clearfix">
                                <p class="area clearfix"><span class="colonna num_ordine">[n&deg; order] - Date</span><span class="colonna totale_ordine">total</span><span class="colonna stato">informations</span><span class="colonna azioni">actions</span></p>
                                <div class="data">
                                    <%
									if rs.recordcount>0 then
										Do while not rs.EOF

										InfoSpedizione=rs("InfoSpedizione")
										NoteCri=rs("NoteCri")
										stato=rs("Stato")

										if stato=0 then etichetta_stato="Cart started"
										if stato=1 then etichetta_stato="Cart started"
										if stato=2 then etichetta_stato="Shipping chosen"
										if stato=12 then etichetta_stato="Shipping chosen"
										if stato=3 then etichetta_stato="Payment to choose"
										if stato=22 then etichetta_stato="Payment to choose"

										if stato=4 then etichetta_stato="Paid with PayPal"
										if stato=5 then etichetta_stato="Payment Canceled"
										if stato=6 then etichetta_stato="When payment"
										if stato=7 then etichetta_stato="Order in process"
										if stato=8 then
											etichetta_stato="Products shipped"
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
										  <%=FormatNumber(TotaleGenerale,2)%>&#8364;
                                        </span>
                                        <span class="colonna stato"><%=etichetta_stato%></span>
                                        <span class="colonna azioni">
                                          <button type="button" name="visualizza" class="button_link" onClick="MM_openBrWindow('stampa_ordine.asp?idordine=<%=rs("PkId")%>&mode=0','','width=760,height=400,scrollbars=yes')">display</button>
										  <%if stato=0 or stato=1 or stato=2 or stato=3 or stato=6 then%>
                                           &nbsp;<button type="button" name="modifica" class="button_link" onClick="document.location.href='ordini_elenco.asp?IdOrdine=<%=rs("PkId")%>&mode=1';">complete the order</button>
                                          <%else%>
                                            <%if stato=12 or stato=22 then%>
                                            <a href="ordini_elenco.asp?IdOrdine=<%=rs("PkId")%>&mode=2"><b>[<%=rs("PkId")%>]&nbsp;-&nbsp;<%=rs("DataAggiornamento")%></b></a>
                                            &nbsp;<button type="button" name="modifica" class="button_link" onClick="document.location.href='ordini_elenco.asp?IdOrdine=<%=rs("PkId")%>&mode=2';">complete the order</button>
                                            <%end if%>
                                          <%end if%>
                                            &nbsp;<button type="button" name="stampa" class="button_link" onClick="MM_openBrWindow('stampa_ordine.asp?idordine=<%=rs("PkId")%>&mode=1','','width=760,height=900,scrollbars=yes')">print</button>
                                        </span>
                                        </p>
                                        <%
                                        rs.movenext
                                        loop
                                        %>

									<%else%>
                                    	<p class="riga">There are no orders</p>
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
    </body>
</html>
<!--#include file="inc_strClose.asp"-->
