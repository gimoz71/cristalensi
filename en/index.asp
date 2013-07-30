<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_clsImageSize.asp"-->
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>CRISTALENSI lights online store lamps italian lighting products sales</title>
		<meta name="description" content="In Cristalensi you find italian lamps for sales, Cristalensi is an ecommerce about lights, online store for italian lighting products and discounted products">
		<meta name="keywords" content="lamps for sales, ecommerce about lights, online store for italian lighting products, Modern lamps, LED Lights, Classic lamps, Rustic lamps, Tiffany, Murano lights, Crystal lamps, lights for children, Spotlights and tracks, Ceiling fans, Outside modern lights, Outside classic lamps , Light bulbs and Drivers">
        <meta name="language" content="en">
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
    <!--plugin facebook-->
    <div id="fb-root"></div>
	<script>(function(d, s, id) {
      var js, fjs = d.getElementsByTagName(s)[0];
      if (d.getElementById(id)) return;
      js = d.createElement(s); js.id = id;
      js.src = "//connect.facebook.net/it_IT/all.js#xfbml=1";
      fjs.parentNode.insertBefore(js, fjs);
    }(document, 'script', 'facebook-jssdk'));</script>    
    <!--fine plugin facebook-->
        <div id="wrap">
            <!--#include file="inc_header.asp"-->
            <div id="main-content">
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <a href="chi_siamo.asp" title="Our Showroom in Italy"><img class="negozio" src="/images/negozio.jpg" alt="CRISTALENSI lights online store"></a>
                            <img class="anni" src="/images/50anni_eng.jpg" alt="Cristalensi 50 years">
                            <h3>Cristalensi, light as idea</h3>
                            <p class="incipit">At just a click away, a vast and refined assortment of italian lighting products for inside and out... Take a look around our <strong>on-line Store</strong> of lights or visit our <a href="chisiamo.asp" title="Our Showroom of lamps and lights"><strong>Showroom of lamps</strong></a>, we can satisfy all your style requirements be they classical or modern. 
                            </p>
                            <!--facebook-->
                            <div class="half_panel social_box left_p_fb">
                            <div class="fb-like-box" data-href="https://www.facebook.com/pages/Cristalensi-vendita-lampade-per-interni-ed-esterni/144109972402284?ref=hl" data-show-faces="true" data-stream="false" data-show-border="false" data-header="true" data-height="230"></div>
                            </div>
                            <!--dicono di noi-->
                            <div class="half_panel social_box right_p">
                            <h4 class="area-commenti">About us...<a href="commenti_elenco.asp" style="float: right; padding: 1px 10px;" class="button_link_red" title="Comments about lighthings shop online and reviews of italian lighting products">ALL COMMENTS AND REVIEWS &raquo;</a></h4>
                            <%
							Set com_rs = Server.CreateObject("ADODB.Recordset")
							sql = "SELECT TOP 5 * FROM Commenti_Clienti WHERE Pubblicato=True ORDER BY PkId DESC"
							com_rs.open sql,conn, 1, 1
				
							if com_rs.recordcount>0 then
								Do While not com_rs.EOF
							%>
							<p><%=Left(NoHTML(com_rs("Testo")), 120)%>...</p>
							<%
								com_rs.movenext
								loop
							end if
							com_rs.close
							%>
                            </div>
                            <div class="spacer"></div>
                            <!--prodotti in offerta-->
                            <h4 class="area clearfix"><span class="title">PRODUCTS ON OFFER: Don't miss the opportunity to buy italian style products!!!</span><a href="offerte.asp" class="right button_link_red" title="Prodotti illuminazone in offerta">ALL THE OFFERS &raquo;</a></h4>
                            <%
							'random prodotti in offerta
							Set prod_rs = Server.CreateObject("ADODB.Recordset")
							sql = "SELECT pkid,codicearticolo,titolo_en,prezzoprodotto,prezzolistino,nomepagina_en FROM Prodotti WHERE Offerta=1 or Offerta=2 ORDER BY Titolo_en ASC"
							prod_rs.open sql,conn, 1, 1
							
							Randomize()
							constnum = 4
				
							if prod_rs.recordcount>0 then
								IF NOT prod_rs.EOF THEN
								rndArray = prod_rs.GetRows()
								prod_rs.Close
							%>
                            <ul class="listino clearfix">
							<%	
								Lenarray =  UBOUND( rndArray, 2 ) + 1
								skip =  Lenarray  / constnum 
								IF Lenarray <= constnum THEN skip = 1
								FOR i = 0 TO Lenarray - 1 STEP skip
									numero = RND * ( skip - 1 )
									id = rndArray( 0, i + numero )
									codicearticolo = rndArray( 1, i + numero )
									titolo_prodotto = rndArray( 2, i + numero )
									prezzoarticolo = rndArray( 3, i + numero )
									prezzolistino = rndArray( 4, i+ numero )
									
									NomePagina = rndArray( 5, i+ numero )
									if Len(NomePagina)>0 then
										NomePagina="/public/pagine/"&NomePagina
										'NomePagina="/public/pagine/scheda_prodotto_en.asp?id="&id
									else
										NomePagina="#"
									end if
									
									
									'recupero l'immagine
									Set img_rs = Server.CreateObject("ADODB.Recordset")
									sql = "SELECT * FROM Immagini WHERE Record="&id&" AND Tabella='Prodotti'"
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
										  End If 
										  
										End With
										Set objImageSize = Nothing
							%>
                            	<li>
                                    <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>"><img src="/public/<%=file_img%>" alt="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>" style="width:<%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>px; height:<%if H<=120 then%><%=H%><%else%>120<%end if%>px;" border="0"><span class="nome-articolo"><%=titolo_prodotto%><%if codicearticolo<>"" then%><br>[<%=codicearticolo%>]<%end if%></span></a>
										<%else%>
                                    <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>"><img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="Product image not available"><span class="nome-articolo"><%=titolo_prodotto%><%if codicearticolo<>"" then%><br>[<%=codicearticolo%>]<%end if%></span></a>	
                                    <%
                                        end if
                                    else
                                        tot_img=0
                                        titolo_img=""
                                        file_img=""
                                    %>
                                    <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>"><img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="Product image not available"><span class="nome-articolo"><%=titolo_prodotto%><%if codicearticolo<>"" then%><br>[<%=codicearticolo%>]<%end if%></span></a>
                                    <%	
                                    end if
                                    img_rs.close
                                    %>
                                    <%if prezzolistino<>"" then%><p class="price">List price: <span><%=prezzolistino%>€</span></p><%end if%>
                                    <%if prezzoarticolo<>"" then%><p class="cristalprice">Cristalensi price: <%=prezzoarticolo%>€</p><%end if%>
                                    <a class="scheda" href="<%=NomePagina%>" title="Description <%=titolo_prodotto%>"><span class="button_link">Product description</span></a>
                                </li>
                                <%
									NEXT
									end if
								%>
                            </ul>
                            <%
							else
								prod_rs.close
							end if
							%>
                            <!--elenco categorie-->
                            <h4 class="area clearfix"><span class="title">CATALOG PRODUCTS</span><a href="ricerca_avanzata_modulo.asp" class="right button_link_red" title="Advanced search lighting products">ADVANCED SEARCH &raquo;</a></h4>
                            </p>
                            <ul class="catalogo clearfix">
                            <%
							'elenco categorie
							Set prod_rs = Server.CreateObject("ADODB.Recordset")
							sql = "SELECT * FROM Categorie1 ORDER BY Posizione"
							prod_rs.open sql,conn, 1, 1
							if prod_rs.recordcount>0 then
								conta=0
								Do while not prod_rs.EOF
								
								cat=prod_rs("PkId")
								titolo_cat=prod_rs("Titolo_en")
								nomepagina_categorie=prod_rs("NomePagina_en")
								if nomepagina_categorie="" then nomepagina_categorie="#"
								if nomepagina_categorie<>"#" then nomepagina_categorie="/public/pagine/"&nomepagina_categorie
								'if nomepagina_categorie<>"#" then nomepagina_categorie="/public/pagine/categorie_en.asp?pkid="&cat
							%>    
                                <li>
                                    <%
									file_img=""
									Set cat_rs = Server.CreateObject("ADODB.Recordset")
									sql = "SELECT * FROM Categorie2 WHERE FkCategoria1="&cat&" AND Logo<>'' ORDER BY Posizione"
									cat_rs.open sql,conn, 1, 1
									if cat_rs.recordcount>0 then
									file_img=cat_rs("logo")
									end if
									cat_rs.close
									
									if file_img<>"" then
									%>
									<a href="<%=nomepagina_categorie%>" title="Products list <%=titolo_cat%>"><img src="/public/<%=file_img%>" width="160" height="120" vspace="2" border="0" alt="<%=titolo_cat%>"><span class="button_link"><%=titolo_cat%></span></a>
										<%else%>
									<a href="<%=nomepagina_categorie%>" title="Products list <%=titolo_cat%>"><img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="image of the category <%=titolo_cat%> not available"><span class="button_link"><%=titolo_cat%></span></a>	
									<%	
										end if
									%>
                                </li>
                            <%
								prod_rs.movenext
								loop	
							end if
							prod_rs.close
							%>
                            
                            </ul>
                            <!--elenco produttori: select con js-->

                            <h4 class="area clearfix"><span class="title">PRODUCERS</span><a href="produttori.asp" class="right button_link_red" title="complete list of producers of lamps">COMPLETE LIST OF PRODUCERS &raquo;</a></h4>
                            <p>If you know the brand of the product you can select below or going to the complete list of producers.
                            </p>
                            <%
							Set cs=Server.CreateObject("ADODB.Recordset")
							sql = "Select * From Produttori order by titolo ASC"
							cs.Open sql, conn, 1, 1
							if cs.recordcount>0 then
							%>
							<SCRIPT LANGUAGE=javascript>
							<!--
								function invia_produttore() {
									document.getElementById("form_produttori").submit();
								}
							// End -->
							</SCRIPT>
							<form method="post" name="form_produttori" id="form_produttori" action="prodotti.asp">
                            <select name="FkProduttore" id="FkProduttore" class="form" onChange="invia_produttore()">
                            <option value="0">Select a brand</option>
                            <%
                            Do While Not cs.EOF
                            %>
                            <option value="<%=cs("pkid")%>"><%=cs("titolo")%></option>
                            <%
                            cs.movenext
                            loop
                            %>
                            </select>
                            </form>
							<%end if%>
							<%cs.close%>
                            <!--fine elenco produttori-->
                            <p>&nbsp;</p>
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