<!--#include file="../../inc_strConn.asp"-->
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Vendita lampade per esterni illuminazione esterna lampade esterne CRISTALENSI vendita online e diretta</title>
		<meta name="description" content="Stai cercando lampade per illuminare il giardino, il parco, la piscina o semplicemente l'esterno della tua abitazione? Cristalensi ha un catalogo molto ampio: scegli tra lampade a soffitto e lampadari, plafoniere, applique, lampade da terra e applique, vendita ">
		<meta name="keywords" content="Vendita lampade esterne applique plafoniere faretti da incasso illuminazione giardini, lampade esterne applique plafoniere faretti da incasso illuminazione giardini Stai cercando lampade per illuminare il giardino, il parco, la piscina o semplicemente l'esterno della tua abitazione? Cristalensi ha un catalogo molto ampio: scegli tra lampade a soffitto e lampadari, plafoniere, applique, lampade da terra e applique ">

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
            <!--#include file="../../inc_header.asp"-->

            <div id="main-content">
                
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <a href="/ricerca_avanzata_modulo.asp" class="right button_link_red" style="font-size:10px; font-weight:bold; margin-top:3px;">RICERCA AVANZATA &raquo;</a>
                            <h1>Vendita lampade per esterni: moderne e classiche</h1>
                            <p>
                                <i>Stai cercando prodotti per l’illuminazione esterna della tua casa? Cristalensi ha un ampio catalogo di prodotti diviso in base alla loro destinazione e ha in vendita lampade per esterni, sia moderne che classiche: scegli tra lampade esterne da terra e pali, lampade esterne a soffitto e plafoniere per esterni, lampade esterne a parete, spot e applique esterne, lampade esterne da incasso, faretti da incasso per esterni oppure lampade led adatte per l'esterno. Tutti prodotti per esterni, per l’illuminazione di giardini, parchi, piscine, loggiati, terrazze con raffinato design dalle più importanti marche e produttori.</i>
                            </p>
                            <%
							Set prod_rs = Server.CreateObject("ADODB.Recordset")
							sql = "SELECT * FROM Categorie2 WHERE FkCategoria1=11 ORDER BY Posizione"
							prod_rs.open sql,conn, 1, 1
							if prod_rs.recordcount>0 then
							%>
                            <ul class="galleria clearfix">
                                <%
								Do while not prod_rs.EOF
								
								id=prod_rs("PkId")
								url="/prodotti.asp?cat="&id
								'url="prodotti.asp?cat="&id
								%>
                                <li>
                                    <a href="<%=url%>" title="<%=titolo_cat%><%=" - "&prod_rs("titolo")%>">
                                        <%
										'file_img="../"&prod_rs("logo")
										file_img="/public/"&prod_rs("logo")
										if file_img<>"" then
										%>
                                        <img src="<%=file_img%>" width="160" height="120" style="margin-bottom: 10px" alt="<%=titolo_cat%><%=" - "&prod_rs("titolo")%>" title="<%=titolo_cat%><%=" - "&prod_rs("titolo")%>" />
                                        <%else%>
                                        <img src="/public/logo_cristalensi_piccolo.jpg" width="160" height="120" style="margin-bottom: 10px" alt="Immagine della categoria <%=titolo_cat%><%=" - "&prod_rs("titolo")%> non disponibile" />
                                        <%end if%>
                                        <span class="button_link"><%=prod_rs("titolo")%></span>
                                    </a>
                                </li>
                                <%
								prod_rs.movenext
								loop	
								%>
                            </ul>
							<%	
							end if
							prod_rs.close
							%>
                            <p></p>
                            <%
							Set prod_rs = Server.CreateObject("ADODB.Recordset")
							sql = "SELECT * FROM Categorie2 WHERE FkCategoria1=12 ORDER BY Posizione"
							prod_rs.open sql,conn, 1, 1
							if prod_rs.recordcount>0 then
							%>
                            <ul class="galleria clearfix">
                                <%
								Do while not prod_rs.EOF
								
								id=prod_rs("PkId")
								url="/prodotti.asp?cat="&id
								'url="prodotti.asp?cat="&id
								%>
                                <li>
                                    <a href="<%=url%>" title="<%=titolo_cat%><%=" - "&prod_rs("titolo")%>">
                                        <%
										'file_img="../"&prod_rs("logo")
										file_img="/public/"&prod_rs("logo")
										if file_img<>"" then
										%>
                                        <img src="<%=file_img%>" width="160" height="120" style="margin-bottom: 10px" alt="<%=titolo_cat%><%=" - "&prod_rs("titolo")%>" title="<%=titolo_cat%><%=" - "&prod_rs("titolo")%>" />
                                        <%else%>
                                        <img src="/public/logo_cristalensi_piccolo.jpg" width="160" height="120" style="margin-bottom: 10px" alt="Immagine della categoria <%=titolo_cat%><%=" - "&prod_rs("titolo")%> non disponibile" />
                                        <%end if%>
                                        <span class="button_link"><%=prod_rs("titolo")%></span>
                                    </a>
                                </li>
                                <%
								prod_rs.movenext
								loop	
								%>
                            </ul>
							<%	
							end if
							prod_rs.close
							%>
                            
                        </div>
                    </div>
                </div>
                <!--#include file="../../inc_sx_prodotti.asp"-->
            </div>
        </div>
         <!--#include file="../../inc_footer.asp"-->
    </body>
</html>
<!--#include file="../../inc_strClose.asp"-->