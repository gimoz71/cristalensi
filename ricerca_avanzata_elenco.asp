<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_clsImageSize.asp"-->
<%
titolo=request("titolo")

cat=request("cat")				  
if cat="" then cat=0

FkProduttore=request("FkProduttore")				  
if FkProduttore="" then FkProduttore=0

prezzo_da=request("prezzo_da")
if prezzo_da="" then prezzo_da=0

prezzo_a=request("prezzo_a")
if prezzo_a="" then prezzo_a=0
%>
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Ricerca articoli illumninazione lampade per esterni lampade per interni</title>
		<meta name="description" content="Fai una ricerca nel catalogo di Cristalensi, showroom vicino Firenze, vende lampade e lampadari on line, prodotti per illuminazione da interno, illuminazione da esterno, lampadari, piantane, plafoniere, lampade da esterno, ventilatori, lampade per bambini e lampade per il bagno, prodotti in molti stili dal moderno al classico.">
<meta name="keywords" content="ricerca prodotti illuminazione, ricerca nel catalogo, vendita lampadari on line, prodotti illuminazione da interni, prodotti illuminazione da esterni, lampade da interno, lampade da esterno, piantane, plafoniere, ventilatori, lampade per bambini, lampade per il bagno, lampade moderne, lampade classiche, lampade rustiche, lampade tiffany, lampade in cristallo, lampade murano, faretti, lampade da incasso, lampade a led, lampade a risparmio energetico, lampade economiche, lampadari economici">
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
                
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <div class="slogan">
                                <h3>Eccezionale sconto!!! Nessun costo di spedizione per ordini superiori a 250€</h3>
                                <p>Per ordini inferiori a 250€ il costo di spedizione è di 10€.<br> Condizioni valide solo per le spedizioni in tutta Italia, isole comprese.</p>
                            </div>
                            <%
								p=request("p")
								if p="" then p=1
								
								order=request("order")
								if order="" then order=1
								'if FkProduttore>0 and order=1 then order=5
								
								if order=1 then ordine="Titolo ASC"
								if order=2 then ordine="Titolo DESC"
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
									sql = sql + "AND (CodiceArticolo LIKE '%"&titolo&"%' OR Titolo LIKE '%"&titolo&"%') "
								end if
								sql = sql + "AND Offerta<10 "
								sql = sql + "ORDER BY "&ordine&""
								prod_rs.open sql,conn, 1, 1
								
								if prod_rs.recordcount>0 then
							%>
                                <h3>Risultato Ricerca avanzata: <%=prod_rs.recordcount%> prodotti trovati</h3>
                                                                
                                <p class="area"> <strong>Ordinamento per prezzo:</strong> <a href="ricerca_avanzata_elenco.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=3"><img src="images/01_new<%if order=3 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" hspace="3" border="0" align="top" alt="ordina i prodotti per prezzo dal pi&ugrave; basso al pi&ugrave; alto" title="ordina i prodotti per prezzo dal pi&ugrave; basso al pi&ugrave; alto" /></a>&nbsp;<a href="ricerca_avanzata_elenco.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=4"><img src="images/10_new<%if order=4 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" border="0" align="top" alt="ordina i prodotti per prezzo dal pi&ugrave; alto al pi&ugrave; basso" title="ordina i prodotti per prezzo dal pi&ugrave; alto al pi&ugrave; basso" /></a>
                              
                                  &nbsp;-&nbsp;<strong>Ordinamento per nome:</strong> <a href="ricerca_avanzata_elenco.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=1"><img src="images/az_new<%if order=1 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" hspace="3" border="0" align="top" alt="ordina i prodotti per titolo dalla A alla Z" title="ordina i prodotti per titolo dalla A alla Z" /></a>&nbsp;<a href="ricerca_avanzata_elenco.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=2"><img src="images/za_new<%if order=2 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" border="0" align="top" alt="ordina i prodotti per titolo dalla Z alla A" title="ordina i prodotti per titolo dalla Z alla A" /></a>
                              
                              
                              
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
									titolo_prodotto=prod_rs("titolo")
									NomePagina=prod_rs("NomePagina")
									if Len(NomePagina)>0 then
										'NomePagina="/public/pagine/"&NomePagina
										NomePagina="scheda_prodotto.asp?id="&id
									else
										NomePagina="#"
									end if
									codicearticolo=prod_rs("codicearticolo")
									descrizione_prodotto=NoHTML(prod_rs("descrizione"))
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
										sql = "SELECT Categorie1.PkId as Cat_Principale, Categorie1.Titolo as Titolo1, Categorie2.PkId, Categorie2.Titolo as Titolo2, Categorie2.Descrizione as Descrizione2 "
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
											file_img="logo_cristalensi_piccolo.jpg"
											if file_img<>"" then
											
											'calcolo misure immagini
											Set objImageSize = New ImageSize
											With objImageSize
											  '.ImageFile = server.mappath("public/"&file_img&"")
											  .ImageFile = path_img&file_img
											  
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
                                        
                                        	<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="public/<%=file_img%>" alt="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>" width="<%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>" height="<%if H<=120 then%><%=H%><%else%>120<%end if%>" border="0"></a>
										<%else%>
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"></a>	
										<%
                                            end if
                                        else
                                            tot_img=0
                                            titolo_img=""
                                            file_img=""
                                        %>
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"></a>
										<%	
                                        end if
                                        img_rs.close
                                        %>
                                        </div>
                                        <div class="data">
                                            <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%> - <%=titolo_cat%>"><strong><%=titolo_prodotto%></strong><%if codicearticolo<>"" then%>&nbsp;[<%=codicearticolo%>]<%end if%></a> <%if fkproduttore_pr>0 then%><span class="produttore">Produttore: <a href="prodotti.asp?FkProduttore=<%=fkproduttore_pr%>" title="Elenco prodotti dello stesso produttore: <%=produttore%>"><strong><%=produttore%></strong></a></span><%end if%>
                                            <p><%=Left(descrizione_prodotto,150)%><%if Len(descrizione_prodotto)>150 then%>...<%end if%><%if FkCategoria2>0 then%></p><p><i>Il prodotto lo trovi nella categoria:</i> <a href="prodotti.asp?cat=<%=FkCategoria2%>" title="Elenco prodotti della stessa categoria: <%=titolo_cat%>"><%=titolo_cat%></a><%end if%></p>
                                            <a href="<%=NomePagina%>" title="Scheda del prodotto&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="button_link scheda-link"><span>Scheda prodotto</span></a>
											<%if tot_img>0 then%><span style="float:right;">[<%if tot_img=1 then%>1 Immagine<%else%><%=tot_img%> Immagini<%end if%>]</span><%end if%>
                                            <%if prezzoarticolo=0 then%>
                                            <p class="cart clearfix"><span class="price">Prezzo listino: <span><%=prezzolistino%>€</span></span>&nbsp;&nbsp;<a href="#" onClick="MM_openBrWindow('richiesta_informazioni.asp?codice=<%=codicearticolo%>&titolo=<%=titolo_prodotto%>&amp;produttore=<%=produttore%>&amp;id=<%=id%>','','width=650,height=650,scrollbars=yes')" class="cart-link button_link_red">Prezzo Cristalensi? clicca qui per un preventivo dal nostro staff</a></p>
                                            <%else%>
                                            <p class="cart clearfix"><%if prezzolistino<>0 then%><span class="price">Prezzo listino: <span><%=prezzolistino%>€</span></span><%end if%>&nbsp;&nbsp;<%if prezzoarticolo<>"" then%><span class="cristalprice">Prezzo Cristalensi: <%=prezzoarticolo%>€&nbsp;&nbsp;<small><i>Iva compresa</i></small></span><%end if%><a href="<%=NomePagina%>" title="Inserisci&nbsp;nel&nbsp;carrello&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="cart-link button_link_red"><span>Inserisci nel carrello</span></a></p>
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
                                  Pag. <strong><%=p%></strong> di <%=prod_rs.PageCount%> - Vai alla Pagina: 
                                  <%if p > 2 then%>&nbsp;[<a href="ricerca_avanzata_elenco.asp?p=1&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">prima pag.</a>]<%end if%>
                                  <% if p >= 5 then %>
                                  [<a href="ricerca_avanzata_elenco.asp?p=<%=p-4%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">&lt;&lt; 5 prec</a>] 
                                  <% end if %>
                                  <% if p > 1 then %>
                                  [<a href="ricerca_avanzata_elenco.asp?p=<%=p-1%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">&lt; prec</a>] 
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
                                  [<a href="ricerca_avanzata_elenco.asp?p=<%=p+1%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">succ &gt;</a>] 
                                  <% end if %>
                                  <% if prod_rs.PageCount-page > 3 then %>
                                  [<a href="ricerca_avanzata_elenco.asp?p=<%=p+5%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">5 succ &gt;&gt;</a>] 
                                  <% end if%>
                                  <%if maxPage>5 and cInt(p)<>prod_rs.PageCount then%>[<a href="ricerca_avanzata_elenco.asp?p=<%=prod_rs.PageCount%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&titolo=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=<%=order%>">ultima pag.</a>]<%end if%>	
                                    </li>
                                  <%end if%>
                                  <!--fine paginazione-->
                                     
                                </ul>
                                <%else%>
                                	<h3>Risultato Ricerca avanzata: <%=prod_rs.recordcount%> prodotti trovati</h3>
                                    <p>Al momento non sono esposti sul sito internet prodotti per la ricerca effettuata. Cristalensi &egrave; contatto con moltissime aziende di prodotti per illuminazione, quindi, se conosci esattamente un articolo e vuoi avere un preventivo di prezzo, riempi il modulo indicandoci il nome del prodotto ed eventualmente alcuni dettagli, 
		  verrai contattato il prima possibile: il nostro staff sar&agrave; a Tua disposizione per qualsiasi chiarimento.<br />
	      <br /><a href="#" onClick="MM_openBrWindow('richiesta_informazioni_produttore.asp?produttore=<%=titolo_produttore%>&amp;id=<%=FkProduttore%>','','scrollbars=yes,width=650,height=650')">Clicca qui per aprire il modulo per la richiesta informazioni e preventivi</a></p>
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
        <script src="js/init.js"></script>
    </body>
</html>
<!--#include file="inc_strClose.asp"-->