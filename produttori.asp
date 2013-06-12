<!--#include file="inc_strConn.asp"-->
<%
Call Visualizzazione("Produttori","0","produttori.asp")
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
        <link href="css/css.css" rel="stylesheet" type="text/css">
        <link href="css/blueberry.css" rel="stylesheet" type="text/css">
        <style type="text/css">
            .clearfix:after {
                content: ".";
                display: block;
                height: 0;
                clear: both;
                visibility: hidden;
            }
        </style>
        <script>
            $(window).load(function() {
                    $('.blueberry').blueberry({
                        pager: false
                    });
            });
        </script>
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
                
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <h3>Elenco produttori articoli per illuminazione</h3>
                            <p>
                                <i>Questo è l'elenco delle imprese produttrici di articoli per illuminazione che riforniscono il nostro negozio.<br>
					    Scegliendo un produttore vedrete l'elenco dei suoi prodotti, da lì potete accedere alla scheda del prodotto e acquistarlo. Se cercate un articolo specifico di un produttore, ma non lo trovate nel suo elenco, contattate il nostro staff per avere informazioni e un preventivo: nel catalogo sul sito internet non sono presenti tutti i prodotti, è stata fatta una selezione dai singoli cataloghi dei produttori.</i>.
                            </p>
                            <%
							Set prod_rs = Server.CreateObject("ADODB.Recordset")
							sql = "SELECT * FROM Produttori ORDER BY Prodotti DESC, Titolo ASC"
							prod_rs.open sql,conn, 1, 1
							if prod_rs.recordcount>0 then
								
							%>
                            <ul class="galleria clearfix">
                                <%
								Do while not prod_rs.EOF
								
								id=prod_rs("PkId")
								titolo=prod_rs("titolo")
								descrizione=prod_rs("descrizione")
								file_img=prod_rs("logo")
								link=prod_rs("prodotti")
								
								url="prodotti.asp?FkProduttore="&id
								%>
                                <%if link=1 then%>
                                <li>
                                    <a href="<%=url%>" title="Elenco prodotti di <%=titolo%>">
                                        <%if file_img<>"" then%>
                                        <img src="public/<%=file_img%>" width="120" height="90" style="margin-bottom: 10px" alt="<%=titolo%>" title="<%=titolo%>" />
                                        <%else%>
                                        <img src="public/logo_cristalensi_piccolo.jpg" width="120" height="90" style="margin-bottom: 10px" alt="logo del produttore <%=titolo%> non disponibile" />
                                        <%end if%>
                                        <span class="button_link"><%=titolo%></span>
                                    </a>
                                </li>
                                <%else%>
                                <li>
                                    <a href="#" onClick="MM_openBrWindow('richiesta_informazioni_produttore.asp?produttore=<%=titolo%>&amp;id=<%=id%>','','scrollbars=yes,width=650,height=650')" title="Richiesta informazioni del produttore <%=titolo%>">
                                        <%if file_img<>"" then%>
                                        <img src="public/<%=file_img%>" width="120" height="90" style="margin-bottom: 10px" alt="<%=titolo%>" title="<%=titolo%>" />
                                        <%else%>
                                        <img src="public/logo_cristalensi_piccolo.jpg" width="120" height="90" style="margin-bottom: 10px" alt="logo del produttore <%=titolo%> non disponibile" />
                                        <%end if%>
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
