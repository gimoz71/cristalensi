<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_clsImageSize.asp"-->
<%
Function CleanStr(sTesto)
	If Len(sTesto)>0 Then
		sTesto = Replace(sTesto,"'","")
		stesto = replace(sTesto, "*", "")
		stesto = replace(sTesto, "%", "")
		stesto = replace(sTesto, "=", "")
		stesto = replace(sTesto, "&", "")
	End If
	CleanStr=sTesto
End Function

titolo=CleanStr(request("titolo"))
'titolo=Replace(titolo, "'", "")
'titolo=Replace(titolo, "&", "")
'titolo=Replace(titolo, "=", "")

cat=request("cat")
if cat="" then cat=0

FkProduttore=request("FkProduttore")
if FkProduttore="" then FkProduttore=0

prezzo_da=CleanStr(request("prezzo_da"))
if prezzo_da="" then prezzo_da=0

prezzo_a=CleanStr(request("prezzo_a"))
if prezzo_a="" then prezzo_a=0
%>
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>advanced search of lights online store lamps italian style Products gallery</title>
		<meta name="description" content="advanced search of lights Products sales Products gallery lamps store online lights italian style">
		<meta name="keywords" content="advanced search of lights online store lamps italian style, Modern lamps, Classic lamps, Rustic lamps, Tiffany lamps, Murano lamps, Crystal lamps, Lamps for kids, Bathroom lights, Spotlights and tracks, Ceiling fans, Outside modern lights, Outside classic lights, Light bulbs and Drivers, LED Lights, Ultramodern lamps">
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

                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <div class="spacer">
                            </div>
                            <%
								p=request("p")
								if p="" then p=1

								order=request("order")
								if order="" then order=1
								'if FkProduttore>0 and order=1 then order=5

								if order=1 then ordine="Titolo_en ASC"
								if order=2 then ordine="Titolo_en DESC"
								if order=3 then ordine="PrezzoProdotto ASC, PrezzoListino ASC"
								if order=4 then ordine="PrezzoProdotto DESC, PrezzoListino DESC"

								Set prod_rs = Server.CreateObject("ADODB.Recordset")
								'if cat>0 then sql = "SELECT * FROM Prodotti WHERE (FkCategoria2="&cat&" and (Offerta=0 or Offerta=2)) ORDER BY "&ordine&""
								'if FkProduttore>0 then sql = "SELECT * FROM Prodotti WHERE (FkProduttore="&FkProduttore&" and (Offerta=0 or Offerta=2)) ORDER BY "&ordine&""
								sql = "SELECT * FROM Prodotti WHERE "
								if prezzo_da>0 or prezzo_a>0 then
									sql = sql + "((PrezzoProdotto>="&prezzo_da&" AND PrezzoProdotto<="&prezzo_a&" AND PrezzoProdotto>0) OR (PrezzoProdotto=0 AND PrezzoListino>="&prezzo_da&" AND PrezzoListino<="&prezzo_a&")) "
								else
									sql = sql + "PrezzoProdotto>=0 "
								end if
								if cat>0 then
									sql = sql + "AND FkCategoria2="&cat&" "
								end if
								if FkProduttore>0 then
									sql = sql + "AND FkProduttore="&FkProduttore&" "
								end if
								if titolo<>"" then
									sql = sql + "AND (CodiceArticolo LIKE '%"&titolo&"%' OR Titolo_en LIKE '%"&titolo&"%') "
								end if
								sql = sql + "AND Offerta<10 "
								sql = sql + "ORDER BY "&ordine&""
								prod_rs.open sql,conn, 1, 1

								if prod_rs.recordcount>0 then
							%>
                                <h3>Result of advanced search: <%=prod_rs.recordcount%> prodotti trovati</h3>

                                <p class="area"> <strong>Arrange by price:</strong> <a href="ricerca_avanzata_elenco.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=3"><img src="/images/01_new<%if order=3 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" hspace="3" border="0" align="top" alt="arrange the products by price from lowest to highest" title="arrange the products by price from lowest to highest" /></a>&nbsp;<a href="ricerca_avanzata_elenco.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=4"><img src="/images/10_new<%if order=4 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" border="0" align="top" alt="arrange the products by price from highest to lowest" title="arrange the products by price from highest to lowest" /></a>

                                  &nbsp;-&nbsp;<strong>Arrange by name:</strong> <a href="ricerca_avanzata_elenco.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=1"><img src="/images/az_new<%if order=1 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" hspace="3" border="0" align="top" alt="arrange the products by name from A to Z" title="arrange the products by name from A to Z" /></a>&nbsp;<a href="ricerca_avanzata_elenco.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=2"><img src="/images/za_new<%if order=2 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" border="0" align="top" alt="arrange the products by name from Z to A" title="arrange the products by name from Z to A" /></a>



                                <ul class="prodotti clearfix">
                                <%

								prod_rs.PageSize = 50
								if prod_rs.recordcount > 0 then
									prod_rs.AbSolutePage = p
									maxPage = prod_rs.PageCount
								End if

								Do while not prod_rs.EOF and rowCount < prod_rs.PageSize
								RowCount = RowCount + 1

									id=prod_rs("pkid")
									titolo_prodotto=prod_rs("titolo_en")
									NomePagina=prod_rs("NomePagina_en")
									if Len(NomePagina)>0 then
										NomePagina="/public/pagine/"&NomePagina
										'NomePagina="/public/pagine/scheda_prodotto.asp?id="&id
									else
										NomePagina="#"
									end if
									codicearticolo=prod_rs("codicearticolo")
									descrizione_prodotto=NoHTML(prod_rs("descrizione_en"))
									allegato_prodotto=prod_rs("allegato")
									prezzoarticolo=prod_rs("PrezzoProdotto")
									prezzolistino=prod_rs("PrezzoListino")

									fkproduttore_pr=prod_rs("fkproduttore")
									FkCategoria2=prod_rs("FkCategoria2")

									if fkproduttore_pr>0 then
										Set pr_rs = Server.CreateObject("ADODB.Recordset")
										sql = "SELECT * FROM Produttori WHERE PkId="&fkproduttore_pr&""
										pr_rs.open sql,conn, 1, 1
										if pr_rs.recordcount>0 then
											produttore=pr_rs("titolo")
										end if
										pr_rs.close
									end if

									if FkCategoria2>0 then
										Set cat_rs = Server.CreateObject("ADODB.Recordset")
										sql = "SELECT Categorie1.PkId as Cat_Principale, Categorie1.Titolo_en as Titolo1, Categorie2.PkId, Categorie2.Titolo_en as Titolo2, Categorie2.Descrizione_en as Descrizione2 "
										sql = sql + "FROM Categorie1 INNER JOIN Categorie2 ON Categorie1.PkId = Categorie2.FkCategoria1 "
										sql = sql + "WHERE Categorie2.PkId="&FkCategoria2
										cat_rs.open sql,conn, 1, 1
										if cat_rs.recordcount>0 then
											'cat_principale=cat_rs("Cat_Principale")
											'titolo_cat=cat_rs("titolo1")&" "&cat_rs("titolo2")
											titolo_cat=cat_rs("titolo2")
										end if
										cat_rs.close
									end if
								%>

                                    <li class="clearfix">
                                        <div class="thumb">
										<%
										'immagine
										Set img_rs = Server.CreateObject("ADODB.Recordset")
										sql = "SELECT * FROM Immagini WHERE Record="&id&" AND Tabella='Prodotti' Order by PkId ASC"
										img_rs.open sql,conn, 1, 1
										if img_rs.recordcount>0 then
											tot_img=img_rs.recordcount
											titolo_img=img_rs("titolo")
											file_img=img_rs("file")

											if file_img<>"" then

											'calcolo misure immagini
											Set objImageSize = New ImageSize
											With objImageSize
											  .ImageFile = server.mappath("/public/"&file_img&"")
											  '.ImageFile = path_img&file_img

											  If .IsImage Then
												W=.ImageWidth
												H=.ImageHeight
												'response.Write("w:"&w&"h:"&h)
											  Else
												'Response.Write "Name: " & .ImageName & "<br>"
												'Response.Write "it isn't an image"
											  End If

											End With
											Set objImageSize = Nothing
										%>

                                        	<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/<%=file_img%>" alt="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>" style="width:<%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>px; height:<%if H<=120 then%><%=H%><%else%>120<%end if%>px;" border="0"></a>
										<%else%>
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="no image for the product <%=titolo_prodotto%>"></a>
										<%
                                            end if
                                        else
                                            tot_img=0
                                            titolo_img=""
                                            file_img=""
                                        %>
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="no image for the product <%=titolo_prodotto%>"></a>
										<%
                                        end if
                                        img_rs.close
                                        %>
                                        <%if tot_img>0 then%><span style="float:right;">[<%if tot_img=1 then%>1 Image<%else%><%=tot_img%> Images<%end if%>]</span><%end if%>
                                        </div>
                                        <div class="data">
                                            <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%> - <%=titolo_cat%>"><strong><%=titolo_prodotto%></strong><%if codicearticolo<>"" then%>&nbsp;[<%=codicearticolo%>]<%end if%></a> <%if fkproduttore_pr>0 then%><span class="produttore">Producer: <a href="prodotti.asp?FkProduttore=<%=fkproduttore_pr%>" title="List of products from the same producers: <%=produttore%>"><strong><%=produttore%></strong></a></span><%end if%>
                                            <p><%=Left(descrizione_prodotto,150)%><%if Len(descrizione_prodotto)>150 then%>...<%end if%><%if FkCategoria2>0 then%></p><p><i>Category:</i> <a href="prodotti.asp?cat=<%=FkCategoria2%>" title="List of products from the same category: <%=titolo_cat%>" style="font-size:9px;"><%=titolo_cat%></a><%end if%></p>
                                            <a href="<%=NomePagina%>" title="Product description&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="button_link scheda-link"><span>Product description</span></a>

                                            <%if prezzoarticolo=0 then%>
                                            <p class="cart clearfix"><span class="price">List price: <span><%=prezzolistino%>&#8364;</span></span>&nbsp;&nbsp;<a href="#" onClick="MM_openBrWindow('richiesta_informazioni.asp?codice=<%=codicearticolo%>&titolo=<%=titolo_prodotto%>&amp;produttore=<%=produttore%>&amp;id=<%=id%>','','width=650,height=650,scrollbars=yes')" class="cart-link button_link_red">Cristalensi price? Click here to have an estimate from our staff</a></p>
                                            <%else%>
                                            <p class="cart clearfix"><%if prezzolistino<>0 then%><span class="price">List price: <span><%=prezzolistino%>&#8364;</span></span><%end if%>&nbsp;&nbsp;<%if prezzoarticolo<>"" then%><span class="cristalprice">Cristalensi price: <%=prezzoarticolo%>&#8364;&nbsp;&nbsp;<small><i>IVA/VAT included</i></small></span><%end if%><a href="<%=NomePagina%>" title="Place in the shopping basket&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="cart-link button_link_red"><span>Add to cart</span></a></p>
                                            <%end if%>
                                        </div>
                                    </li>
                                 <%
									prod_rs.movenext
									loop
								%>
                                  <!--paginazione-->
								  <%if prod_rs.recordcount>50 then%>
                                  <li>
                                  Pag. <strong><%=p%></strong> of <%=prod_rs.PageCount%> - Go to Page:
                                  <%if p > 2 then%>&nbsp;[<a href="ricerca_avanzata_elenco.asp?p=1&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">first page</a>]<%end if%>
                                  <% if p >= 5 then %>
                                  [<a href="ricerca_avanzata_elenco.asp?p=<%=p-4%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">&lt;&lt; 5 prev</a>]
                                  <% end if %>
                                  <% if p > 1 then %>
                                  [<a href="ricerca_avanzata_elenco.asp?p=<%=p-1%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">&lt; prev</a>]
                                  <% end if %>
                                  <% for page = p+1 to p+4 %>
                                  <%if not page>maxPage then%><a href="ricerca_avanzata_elenco.asp?p=<%=Page%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>"><%=page%></a>&nbsp;<%end if%>
                                  <% if page >= prod_rs.PageCount then
                                     page = p+4
                                     'exit for
                                    end if
                                  %>
                                  <% next %>
                                  <% 'if page-1 <= prod_rs.PageCount then %>
                                  <% if cInt(p) < maxPage then %>
                                  [<a href="ricerca_avanzata_elenco.asp?p=<%=p+1%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">next &gt;</a>]
                                  <% end if %>
                                  <% if prod_rs.PageCount-page > 3 then %>
                                  [<a href="ricerca_avanzata_elenco.asp?p=<%=p+5%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">5 next &gt;&gt;</a>]
                                  <% end if%>
                                  <%if maxPage>5 and cInt(p)<>prod_rs.PageCount then%>[<a href="ricerca_avanzata_elenco.asp?p=<%=prod_rs.PageCount%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">the last page</a>]<%end if%>
                                    </li>
                                  <%end if%>
                                  <!--fine paginazione-->

                                </ul>
                                <%else%>
                                	<h3>Result of advanced search: <%=prod_rs.recordcount%> products</h3>
                                    <p>Currently no products are displayed on the website for the search. Cristalensi has contact with many companies of lighting products, so if you know exactly one item and want to get a price quote, fill out the form indicating the name of the product and possibly some details,
will contact you as soon as possible: our staff will be at your disposal for any clarification.<br />
	      <br /><a href="#" onClick="MM_openBrWindow('richiesta_informazioni_produttore.asp?produttore=<%=titolo_produttore%>&amp;id=<%=FkProduttore%>','','scrollbars=yes,width=650,height=650')">Click here to open the form to request information and estimate</a></p>
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
<!--#include file="inc_strClose.asp"-->
