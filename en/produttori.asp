<!--#include file="inc_strConn.asp"-->
<%
Call Visualizzazione("Produttori","0","produttori.asp")
%>
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>list of italian producers of lighting products lamps lights producers' catalogs</title>
		<meta name="description" content="This is a list of italian producers of lighting products serving our shop, a selection of lamps and lights was made by individual italian producers' catalogs">
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
                
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <h1>List of italian producers of lighting products</h1>
                            <p>
                                <i>This is a list of italian producers of lighting products serving our shop.
Choosing a manufacturer will see a list of its products, from there you can access the product page and buy it. If you are looking for a specific article of a italian producer, but can not find it in its list, please contact our staff for information and an estimate: in the catalog on the website are not present all the products, a selection was made by individual italian producers' catalogs.</i>.
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
								
								url="/en/prodotti.asp?FkProduttore="&id
								%>
                                <%if link=1 then%>
                                <li>
                                    <a href="<%=url%>" title="Elenco prodotti di <%=titolo%>">
                                        <%if file_img<>"" then%>
                                        <img src="/public/<%=file_img%>" style="margin-bottom: 10px; width:120px; height:90px;" alt="<%=titolo%>" title="<%=titolo%>" />
                                        <%else%>
                                        <img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" style="margin-bottom: 10px" alt="logo of producer <%=titolo%> unavailable" />
                                        <%end if%>
                                        <div class="clear"></div>
                                        <span class="button_link"><%=titolo%></span>
                                    </a>
                                </li>
                                <%else%>
                                <li>
                                    <a href="#" onClick="MM_openBrWindow('richiesta_informazioni_produttore.asp?produttore=<%=titolo%>&amp;id=<%=id%>','','scrollbars=yes,width=650,height=650')" title="Request information about the producers <%=titolo%>">
                                        <%if file_img<>"" then%>
                                        <img src="/public/<%=file_img%>" style="margin-bottom: 10px; width:120px; height:90px;" alt="<%=titolo%>" title="<%=titolo%>" />
                                        <%else%>
                                        <img src="/public/logo_cristalensi_piccolo.jpg" style="margin-bottom: 10px" alt="logo of producer <%=titolo%> unavailable" />
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
                                <p><br /><br /><br />No producers</p>
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
