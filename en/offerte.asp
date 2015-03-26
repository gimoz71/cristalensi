<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_clsImageSize.asp"-->
<%
	order=request("order")
	if order="" then order=3
	if order=1 then ordine="Titolo_en ASC"
	if order=2 then ordine="Titolo_en DESC"
	if order=3 then ordine="prezzoprodotto ASC"
	if order=4 then ordine="prezzoprodotto DESC"
%>
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>lighting products on offer low price lights shopping</title>
		<meta name="description" content="In Cristalensi store you find italian lamps on offer, Cristalensi is an ecommerce about lights, online store for italian lighting products and discounted products">
		<meta name="keywords" content="store online shop ecommerce lights lamps lighting products discounted articles on offer italian style">
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
                            <h1>Lights offer</h1>
                            <p><em>In this page you will find all of our offers on lighting: they are the catalog products at fantastic prices. Every product has its own detailed card, to see it, all you have to do is click on the name or on the photo of the article.<br>
					    Instead, to consult the entire catalog click here on <a href="prodotti.asp" title="Products gallery">[Products]</a> (or on the same voice in the menu above)  but you can also choose a category or a producer from the menu on the left.</em>
                          </p>
                          <SCRIPT LANGUAGE=javascript>
                                    <!--
                                        function invia_account() {
                                            document.getElementById("form_prodotti").submit();
                                        }
                                    // End -->
                                    </SCRIPT>
                             		<div class="half_panel left_p">
                                    <form method="post" action="prodotti.asp" name="form_prodotti" id="form_prodotti">
                                      <p>
                                      Haven't found what you're looking for? Choose a category:
                                        <%
                                        Set cs=Server.CreateObject("ADODB.Recordset")
                                        sql = "SELECT Categorie1.PkId as PkId_1, Categorie1.Titolo_en as Titolo_1, Categorie2.PkId as PkId_2, Categorie2.Titolo_en as Titolo_2 "
                                        sql = sql + "FROM Categorie1 INNER JOIN Categorie2 ON Categorie1.PkId = Categorie2.Fkcategoria1 "
                                        'sql = sql + "WHERE Categorie2.FkCategoria1 = "&cat_principale&" "
                                        sql = sql + "ORDER BY Categorie1.Titolo_en ASC, Categorie2.Titolo_en ASC"
                                        cs.Open sql, conn, 1, 1
                                        %>
                                        <select name="Cat" id="Cat" class="form" onChange="invia_account()" style="margin-top:10px;">
                                            <option title="Choose a category" value="0">Choose a category</option>
											<%
                                            if cs.recordcount>0 then
                                            Do While Not cs.EOF
                                            %>
                                            <option title="<%=cs("Titolo_2")%>" value=<%=cs("pkid_2")%>><%=cs("Titolo_2")%></option>
                                            <%
                                            cs.movenext
                                            loop
                                            end if
                                            %>
                                         </select>
                                         <%cs.close%>
                                        </p>
                                       </form>
                                       </div>
                                       <div class="half_panel right_p">
                                        <p>Or, for a more detailed search, use the<br>
                                        <span><a href="ricerca_avanzata_modulo.asp" class="button_link_red" style="margin-top:7px;">ADVANCED SEARCH</a></span>
                                        </p>
                                        </div>
                                    <div class="clear"></div>

                            <p class="area"> <strong>Arrange by price:</strong>
                            <a href="offerte.asp?order=3"><img src="/images/01_new<%if order=3 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" hspace="3" border="0" align="top" alt="arrange the products by price from lowest to highest" title="arrange the products by price from lowest to highest" /></a>
                            <a href="offerte.asp?order=4"><img src="/images/10_new<%if order=4 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" border="0" align="top" alt="arrange the products by price from highest to lowest" title="arrange the products by price from highest to lowest" /></a>
                            &nbsp;-&nbsp;
                            <strong>Arrange by name:</strong>
                            <a href="offerte.asp?order=1"><img src="/images/az_new<%if order=1 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px"  hspace="3" border="0" align="top" alt="arrange the products by name from A to Z" title="arrange the products by name from A to Z" /></a>&nbsp;
                            <a href="offerte.asp?order=2"><img src="/images/za_new<%if order=2 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px"  border="0" align="top" alt="arrange the products by name from Z to A" title="arrange the products by name from Z to A" /></a></p>

                          <ul class="prodotti clearfix">

                            <%
                            Set prod_rs = Server.CreateObject("ADODB.Recordset")
                            sql = "SELECT * FROM Prodotti WHERE Offerta=1 or Offerta=2 ORDER BY "&ordine&""
                            prod_rs.open sql,conn, 1, 1
                            if prod_rs.recordcount>0 then

                                    Do while not prod_rs.EOF

                                            id=prod_rs("pkid")
                                            titolo_prodotto=prod_rs("titolo_en")

                                            NomePagina=prod_rs("NomePagina_en")
                                            if NomePagina="" then NomePagina="#"
                                            if NomePagina<>"#" then NomePagina="/public/pagine/"&NomePagina
                                            'if NomePagina<>"#" then NomePagina="/public/pagine/scheda_prodotto_en.asp?pkid="&id

                                            codicearticolo=prod_rs("codicearticolo")
                                            descrizione_prodotto=prod_rs("descrizione_en")
                                            allegato_prodotto=prod_rs("allegato")
                                            prezzoarticolo=prod_rs("prezzoprodotto")
                                            prezzolistino=prod_rs("prezzolistino")
                                            fkproduttore=prod_rs("fkproduttore")
                                            if fkproduttore="" then fkproduttore=0

                                            if fkproduttore>0 then
                                                    Set pr_rs = Server.CreateObject("ADODB.Recordset")
                                                    sql = "SELECT * FROM Produttori WHERE PkId="&fkproduttore&""
                                                    pr_rs.open sql,conn, 1, 1
                                                    if pr_rs.recordcount>0 then
                                                            produttore=pr_rs("titolo")
                                                    end if
                                                    pr_rs.close
                                            end if

                                            FkCategoria2 = prod_rs("FkCategoria2")
                                            if FkCategoria2="" then FkCategoria2=0

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
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/<%=file_img%>" alt="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>" style="max-width: <%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>px; height: <%if H<=120 then%><%=H%><%else%>120<%end if%>px;" border="0"></a>
										<%else%>
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/logo_cristalensi_piccolo.jpg" alt="no image for the product <%=titolo_prodotto%>"></a>
                                    <%
                                            end if
                                    else
                                            tot_img=0
                                            titolo_img=""
                                            file_img=""
                                    %>
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/logo_cristalensi_piccolo.jpg" alt="no image for the product <%=titolo_prodotto%>"></a>
                                    <%
                                    end if
                                    img_rs.close
                                    %>
                                    <%if tot_img>0 then%><span>[<%if tot_img=1 then%>1 Image<%else%><%=tot_img%> Images<%end if%>]</span><%end if%>
                                    </div>
                                    <div class="data">
                                        <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%> - <%=titolo_cat%>"><strong><%=titolo_prodotto%></strong><%if codicearticolo<>"" then%>&nbsp;[<%=codicearticolo%>]<%end if%></a> <%if fkproduttore>0 then%><span class="produttore">Producers: <a href="prodotti.asp?FkProduttore=<%=fkproduttore%>" title="List of products of producers: <%=produttore%>"><strong><%=produttore%></strong></a></span><%end if%>
                                            <p><%=Left(descrizione_prodotto,150)%><%if Len(descrizione_prodotto)>150 then%>...<%end if%><%if FkCategoria2>0 then%><br /><i>Category:</i> <a href="prodotti.asp?cat=<%=FkCategoria2%>" title="List of products from the same category: <%=titolo_cat%>" style="font-size:9px;"><%=titolo_cat%></a><%end if%></p>
                                            <a href="<%=NomePagina%>" title="Product description&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="button_link scheda-link"><span>Product description</span></a>
                                            <%if prezzoarticolo=0 then%>
                                            <p class="cart clearfix"><span class="price">List price: <span><%=prezzolistino%>&#8364;</span></span>&nbsp;&nbsp;<span class="cristalprice"><a href="#" onClick="MM_openBrWindow('richiesta_informazioni.asp?codice=<%=codicearticolo%>&titolo=<%=titolo_prodotto%>&amp;produttore=<%=produttore%>&amp;id=<%=id%>','','width=650,height=650,scrollbars=yes')">Cristalensi price? click here to have an estimate from our staff</a></span></p>
                                            <%else%>
                                            <p class="cart clearfix"><%if prezzolistino<>0 then%><span class="price">List price: <span><%=prezzolistino%>&#8364;</span></span><%end if%>&nbsp;&nbsp;<%if prezzoarticolo<>"" then%><span class="cristalprice">Cristalensi price: <%=prezzoarticolo%>&#8364;&nbsp;&nbsp;<small><i>IVA/VAT included</i></small></span><%end if%><a href="<%=NomePagina%>" title="Place in the shopping basket <%=titolo_prodotto%> <%=codicearticolo%>" class="cart-link button_link_red"><span>Add to cart</span></a></p>
                                            <%end if%>
                                    </div>
                                </li>
                                <%
									prod_rs.movenext
									loop
								end if
								prod_rs.close
								%>
                          </ul>
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
