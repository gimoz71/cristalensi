<!--#include file="inc_strConn.asp"-->
<%
Call Visualizzazione("Produttori","0","produttori.asp")
%>
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Elenco produttori articoli illuminazione CRISTALENSI Negozio lampadari, piantane, plafoniere, lampade esterne, ventilatori, lampade per bambini, lampade per il bagno</title>
		<meta name="description" content="Elenco produttori articoli illuminazione, elenco di imprese illuminazione, catalogo dei produttori di lampadari, piantane, plafoniere, lampade esterne, ventilatori, prodotti per bambini">
		<meta name="keywords" content="Produttori articoli illuminazione, imprese illuminazione, produttori di lampadari, piantane, plafoniere, lampade esterne, ventilatori, prodotti per bambini">
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="/js/media-queries-ie.js"></script>
        <![endif]-->
        <link href="/css/css.css" rel="stylesheet" type="text/css">
        <link href="/css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="/css/tipTip.css" rel="stylesheet" type="text/css">
        
        <link href="/css/cookies-enabler.css" rel="stylesheet" type="text/css">
        
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
		<script type="text/plain" class="ce-script">
        
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
                
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <h1>Elenco produttori articoli per illuminazione</h1>
                            <p>
                                <i>Questo &eacute; l'elenco delle imprese produttrici di articoli per illuminazione che riforniscono il nostro negozio.<br>
					    Scegliendo un produttore vedrete l'elenco dei suoi prodotti, da l&igrave; potete accedere alla scheda del prodotto e acquistarlo. Se cercate un articolo specifico di un produttore, ma non lo trovate nel suo elenco, contattate il nostro staff per avere informazioni e un preventivo: nel catalogo sul sito internet non sono presenti tutti i prodotti, &eacute; stata fatta una selezione dai singoli cataloghi dei produttori.</i>.
                            </p>
                            <%
							Set prod_rs = Server.CreateObject("ADODB.Recordset")
							sql = "SELECT * FROM Produttori ORDER BY Prodotti DESC, Titolo ASC"
							prod_rs.open sql,conn, 1, 1
							if prod_rs.recordcount>0 then
								
							%>
                            <ul class="produttori clearfix">
                                <%
								Do while not prod_rs.EOF
								
								id=prod_rs("PkId")
								titolo=prod_rs("titolo")
								descrizione=prod_rs("descrizione")
								file_img=prod_rs("logo")
								link=prod_rs("prodotti")
								
								url="/prodotti.asp?FkProduttore="&id
								%>
                                <%if link=1 then%>
                                <li>
                                    <a href="<%=url%>" title="Elenco prodotti di <%=titolo%>">
                                        <%if file_img<>"" then%>
                                        <img src="/public/<%=file_img%>" style="margin-bottom: 10px; width:120px; height:90px;" alt="<%=titolo%>" title="<%=titolo%>" />
                                        <%else%>
                                        <img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" style="margin-bottom: 10px" alt="logo del produttore <%=titolo%> non disponibile" />
                                        <%end if%>
                                        <div class="clear"></div>
                                        <span class="button_link"><%=titolo%></span>
                                    </a>
                                </li>
                                <%else%>
                                <li>
                                    <a href="#" onClick="MM_openBrWindow('/richiesta_informazioni_produttore.asp?produttore=<%=titolo%>&amp;id=<%=id%>','','scrollbars=yes,width=650,height=650')" title="Richiesta informazioni del produttore <%=titolo%>">
                                        <%if file_img<>"" then%>
                                        <img src="/public/<%=file_img%>" style="margin-bottom: 10px; width:120px; height:90px;" alt="<%=titolo%>" title="<%=titolo%>" />
                                        <%else%>
                                        <img src="/public/logo_cristalensi_piccolo.jpg" style="margin-bottom: 10px" alt="logo del produttore <%=titolo%> non disponibile" />
                                        <%end if%>
                                        <div class="clear"></div>
                                        <span class="button_link"><%=titolo%></span>
                                    </a>
                                </li>
                                <%end if%>
                                <%
								prod_rs.movenext
								loop	
								%>
                            </ul>
                            <%else%>
                                <p><br /><br /><br />Nessun produttore presente</p>
							<%	
							end if
							prod_rs.close
							%>
                            
                        </div>
                    </div>
                </div>
                <!--#include file="inc_sx_prodotti.asp"-->
            </div>
        </div>
         <!--#include file="inc_footer.asp"-->
    </body>
</html>
