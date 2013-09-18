<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_clsImageSize.asp"-->
<%
cat=request("cat")				  
if cat="" then cat=0

FkProduttore=request("FkProduttore")				  
if FkProduttore="" then FkProduttore=0


if cat>0 then
	Set cat_rs = Server.CreateObject("ADODB.Recordset")
	'sql = "SELECT * FROM Categorie2 WHERE PKId="&cat
	sql = "SELECT Categorie1.PkId as Cat_Principale, Categorie1.Titolo as Titolo1, Categorie1.Testo1 as Testo1, Categorie2.PkId, Categorie2.Titolo as Titolo2, Categorie2.Descrizione as Descrizione2, Categorie2.Testo1 as Titolo1Cat2 "
	sql = sql + "FROM Categorie1 INNER JOIN Categorie2 ON Categorie1.PkId = Categorie2.FkCategoria1 "
	sql = sql + "WHERE Categorie2.PkId="&cat
	cat_rs.open sql,conn, 1, 1
	if cat_rs.recordcount>0 then
		cat_principale=cat_rs("Cat_Principale")
		'titolo_cat=cat_rs("titolo1")&" "&cat_rs("titolo2")
		titolo_cat=cat_rs("titolo2")
		title_cat=titolo_cat&" "&cat_rs("testo1")
		descrizione_cat=cat_rs("descrizione2")
		nuovo_title_cat=cat_rs("Titolo1Cat2")
	end if
	cat_rs.close
	
	Call Visualizzazione("Categorie2",Cat,"prodotti.asp")
elseif FkProduttore>0 then

	Set az_rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Produttori WHERE PKId="&FkProduttore
	az_rs.open sql,conn, 1, 1
	if az_rs.recordcount>0 then
		titolo_produttore=az_rs("titolo")
	end if
	az_rs.close
	
	Call Visualizzazione("Produttori",FkProduttore,"prodotti.asp")
else

	Call Visualizzazione("",0,"prodotti.asp")
end if
%>
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title><%if cat>0 then%><%=nuovo_title_cat%><%end if%><%if FkProduttore>0 then%><%=titolo_produttore%> catalogo prodotti illuminazione vendita online Cristalensi<%end if%><%if cat=0 and FkProduttore=0 then%>catalogo articoli illuminazione da interno illuminazione da esterno vendita online Cristalesni<%end if%></title>
		<meta name="description" content="<%if cat>0 then%><%=NoHTML(descrizione_cat)%><%end if%><%if FkProduttore>0 then%>Catalogo prodotti in vendita di <%=titolo_produttore%>, vendita online prodotti illuminazione su Cristalensi<%end if%><%if cat=0 and FkProduttore=0 then%>catalogo prodotti illuminazione da interno, illuminazione da esterno, vendita online su Cristalesni<%end if%>">
		<meta name="keywords" content="<%if cat>0 then%><%=nuovo_title_cat%><%end if%><%if FkProduttore>0 then%><%=titolo_produttore%> catalogo prodotti illuminazione vendita online Cristalensi<%end if%><%if cat=0 and FkProduttore=0 then%>catalogo prodotti illuminazione da interno illuminazione da esterno vendita online Cristalesni<%end if%>">
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
                            <div class="slogan">
                                <h3>Eccezionale sconto!!! Nessun costo di spedizione per ordini superiori a 250€</h3>
                                <p>Per ordini inferiori a 250&#8364; il costo di spedizione &eacute; di 10&#8364;.<br> Condizioni valide solo per le spedizioni in tutta Italia, isole comprese.</p>
                            </div>
                            <%
							'elenco prodotti di una categoria o di un produttore
							if cat>0 or FkProduttore>0 then%>
                                <h1>Catalogo prodotti: <%if cat>0 then%><%=titolo_cat%><%end if%><%if FkProduttore>0 then%><%=titolo_produttore%><%end if%></h1>
                                <%if descrizione_cat<>"" then%>
                                <p>
                                    <i><%=descrizione_cat%></i>
                                </p>
                                <%else%>
                                <p>&nbsp;</p>
                                
								<%end if%>
                                
								<%if cat>0 then%>
									
									<SCRIPT LANGUAGE=javascript>
                                    <!--
                                        function invia_account() {
                                            document.getElementById("form_prodotti").submit();
                                        }
                                    // End -->
                                    </SCRIPT>
                             		
                                    <div class="half_panel left_p">
                                        <form method="post" action="/prodotti.asp" name="form_prodotti" id="form_prodotti">
                                        <p>
                                          Non hai trovato il prodotto che cercavi?<br />scegli un'altra categoria:
                                          <%
                                          Set cs=Server.CreateObject("ADODB.Recordset")
                                          sql = "SELECT Categorie1.PkId as PkId_1, Categorie1.Titolo as Titolo_1, Categorie2.PkId as PkId_2, Categorie2.Titolo as Titolo_2 "
                                          sql = sql + "FROM Categorie1 INNER JOIN Categorie2 ON Categorie1.PkId = Categorie2.Fkcategoria1 "
                                          sql = sql + "WHERE Categorie2.FkCategoria1 = "&cat_principale&" "
                                          sql = sql + "ORDER BY Categorie1.Titolo ASC, Categorie2.Titolo ASC"
                                          cs.Open sql, conn, 1, 1
                                          %>
                                          <select name="Cat" id="Cat" class="form" onChange="invia_account()" style="margin-top:10px;">
                                              <%
                                              if cs.recordcount>0 then
                                              Do While Not cs.EOF
                                              %>
                                              <option title="<%=cs("Titolo_2")%>" value=<%=cs("pkid_2")%> <%if cInt(cat)=cInt(cs("pkid_2")) then%> selected<%end if%>><%=cs("Titolo_2")%></option>
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
                                        <p>Oppure, per una ricerca maggiormente dettagliata, usa la<br>
                                        <span><a href="/ricerca_avanzata_modulo.asp" class="button_link_red" style="margin-top:7px;">RICERCA AVANZATA</a></span>
                                        </p>
                                    </div>
                                    <div class="clear"></div>
                              	<%end if%>
                                <%
									p=request("p")
									if p="" then p=1
									
									order=request("order")
									if order="" then order=1
									if FkProduttore>0 and order=1 then order=5
								%>
                                
                                <p class="area"> <strong>Ordinamento per prezzo:</strong> <a href="prodotti.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=3"><img src="/images/01_new<%if order=3 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" hspace="3" border="0" align="top" alt="ordina i prodotti per prezzo dal pi&ugrave; basso al pi&ugrave; alto" title="ordina i prodotti per prezzo dal pi&ugrave; basso al pi&ugrave; alto" /></a>&nbsp;<a href="prodotti.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=4"><img src="/images/10_new<%if order=4 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" border="0" align="top" alt="ordina i prodotti per prezzo dal pi&ugrave; alto al pi&ugrave; basso" title="ordina i prodotti per prezzo dal pi&ugrave; alto al pi&ugrave; basso" /></a>
                              <%if FkProduttore>0 then%>
                                  &nbsp;-&nbsp;<strong>Ordinamento per codice:</strong> <a href="prodotti.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=5"><img src="/images/az_new<%if order=5 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" hspace="3" border="0" align="top" alt="ordina i prodotti per codice articolo dalla A alla Z" title="ordina i prodotti per codice articolo dalla A alla Z" /></a>&nbsp;<a href="prodotti.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=6"><img src="/images/za_new<%if order=6 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" border="0" align="top" alt="ordina i prodotti per codice articolo dalla Z alla A" title="ordina i prodotti per codice articolo dalla Z alla A" /></a>
                              <%else%>
                                  &nbsp;-&nbsp;<strong>Ordinamento per nome:</strong> <a href="prodotti.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=1"><img src="/images/az_new<%if order=1 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" hspace="3" border="0" align="top" alt="ordina i prodotti per titolo dalla A alla Z" title="ordina i prodotti per titolo dalla A alla Z" /></a>&nbsp;<a href="prodotti.asp?cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=2"><img src="/images/za_new<%if order=2 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" border="0" align="top" alt="ordina i prodotti per titolo dalla Z alla A" title="ordina i prodotti per titolo dalla Z alla A" /></a>
                              <%end if%>
                              
                              <%
									if order=1 then ordine="Titolo ASC"
									if order=2 then ordine="Titolo DESC"
									if order=3 then ordine="PrezzoProdotto ASC, PrezzoListino ASC"
									if order=4 then ordine="PrezzoProdotto DESC, PrezzoListino DESC"
									if order=5 then ordine="CodiceArticolo ASC"
									if order=6 then ordine="CodiceArticolo DESC"
								
									Set prod_rs = Server.CreateObject("ADODB.Recordset")
									if cat>0 then sql = "SELECT * FROM Prodotti WHERE (FkCategoria2="&cat&" and (Offerta=0 or Offerta=2)) ORDER BY "&ordine&""
									if FkProduttore>0 then sql = "SELECT * FROM Prodotti WHERE (FkProduttore="&FkProduttore&" and (Offerta=0 or Offerta=2)) ORDER BY "&ordine&""
									prod_rs.open sql,conn, 1, 1
									if prod_rs.recordcount>0 then
								
									prod_rs.PageSize = 10
									if prod_rs.recordcount > 0 then 
										prod_rs.AbSolutePage = p 
										maxPage = prod_rs.PageCount 
									End if
								%>
                                <ul class="prodotti clearfix">
                                <% 
										
										Do while not prod_rs.EOF and rowCount < prod_rs.PageSize
										RowCount = RowCount + 1
										
											id=prod_rs("pkid")
											titolo_prodotto=prod_rs("titolo")
											NomePagina=prod_rs("NomePagina")
											if Len(NomePagina)>0 then
												NomePagina="/public/pagine/"&NomePagina
												'NomePagina="/public/pagine/scheda_prodotto.asp?id="&id
											else
												NomePagina="#"
											end if
											codicearticolo=prod_rs("codicearticolo")
											descrizione_prodotto=NoHTML(prod_rs("descrizione"))
											allegato_prodotto=prod_rs("allegato")
											prezzoarticolo=prod_rs("PrezzoProdotto")
											prezzolistino=prod_rs("PrezzoListino")
											
											if fkproduttore=0 then
												fkproduttore_pr=prod_rs("fkproduttore")
												if fkproduttore_pr="" then fkproduttore_pr=0
												
												if fkproduttore_pr>0 then
													Set pr_rs = Server.CreateObject("ADODB.Recordset")
													sql = "SELECT * FROM Produttori WHERE PkId="&fkproduttore_pr&""
													pr_rs.open sql,conn, 1, 1
													if pr_rs.recordcount>0 then
														produttore=pr_rs("titolo")
													end if
													pr_rs.close
												end if
											end if
											
											if cat=0 then
												FkCategoria2=prod_rs("FkCategoria2")
												if FkCategoria2="" then FkCategoria2=0
												
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
                                        
                                        	<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/<%=file_img%>" alt="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>" style="max-width:<%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>px; height:<%if H<=120 then%><%=H%><%else%>120<%end if%>px;" border="0"></a>
										<%else%>
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"></a>	
										<%
                                            end if
                                        else
                                            tot_img=0
                                            titolo_img=""
                                            file_img=""
                                        %>
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"></a>
										<%	
                                        end if
                                        img_rs.close
                                        %>
                                        <%if tot_img>0 then%><span>[<%if tot_img=1 then%>1 Immagine<%else%><%=tot_img%> Immagini<%end if%>]</span><%end if%>
                                        </div>
                                        
                                        <div class="data">
                                            <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%> - <%=titolo_cat%>"><strong><%=titolo_prodotto%></strong><%if codicearticolo<>"" then%>&nbsp;[<%=codicearticolo%>]<%end if%></a> <%if fkproduttore_pr>0 then%><span class="produttore">Produttore: <a href="prodotti.asp?FkProduttore=<%=fkproduttore_pr%>" title="Elenco prodotti dello stesso produttore: <%=produttore%>"><strong><%=produttore%></strong></a></span><%end if%>
                                            <p><%=Left(descrizione_prodotto,150)%><%if Len(descrizione_prodotto)>150 then%>...<%end if%><%if FkCategoria2>0 then%></p><p><i>Categoria:</i> <a href="prodotti.asp?cat=<%=FkCategoria2%>" title="Elenco prodotti della stessa categoria: <%=titolo_cat%>" style="font-size:9px;"><%=titolo_cat%></a><%end if%></p>
                                            <a href="<%=NomePagina%>" title="Scheda del prodotto&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="button_link scheda-link"><span>Scheda prodotto</span></a>
											
                                            <%if prezzoarticolo=0 then%>
                                            <p class="cart clearfix"><span class="price">Prezzo listino: <span><%=prezzolistino%>&#8364;</span></span>&nbsp;&nbsp;<a href="#" onClick="MM_openBrWindow('richiesta_informazioni.asp?codice=<%=codicearticolo%>&titolo=<%=titolo_prodotto%>&amp;produttore=<%=produttore%>&amp;id=<%=id%>','','width=650,height=650,scrollbars=yes')" class="cart-link button_link_red">Prezzo Cristalensi? clicca qui per un preventivo dal nostro staff</a></p>
                                            <%else%>
                                            <p class="cart clearfix"><%if prezzolistino<>0 then%><span class="price">Prezzo listino: <span><%=prezzolistino%>&#8364;</span></span><%end if%>&nbsp;&nbsp;<%if prezzoarticolo<>"" then%><span class="cristalprice">Prezzo Cristalensi: <%=prezzoarticolo%>&#8364;&nbsp;&nbsp;<small><i>Iva compresa</i></small></span><%end if%><a href="<%=NomePagina%>" title="Inserisci&nbsp;nel&nbsp;carrello&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="cart-link button_link_red"><span>Inserisci nel carrello</span></a></p>
                                            <%end if%>
                                        </div>
                                        
                                    </li>
                                 <%
									prod_rs.movenext
									loop	
								%>
                                  <!--paginazione-->
								  <%if prod_rs.recordcount>10 then%>
                                  <li>
                                  Pag. <strong><%=p%></strong> di <%=prod_rs.PageCount%> - Vai alla Pagina: 
                                  <%if p > 2 then%>&nbsp;[<a href="/prodotti.asp?p=1&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=<%=order%>">prima pag.</a>]<%end if%>
                                  <% if p >= 5 then %>
                                  [<a href="/prodotti.asp?p=<%=p-4%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=<%=order%>">&lt;&lt; 5 prec</a>] 
                                  <% end if %>
                                  <% if p > 1 then %>
                                  [<a href="/prodotti.asp?p=<%=p-1%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=<%=order%>">&lt; prec</a>] 
                                  <% end if %>
                                  <% for page = p+1 to p+4 %>
                                  <%if not page>maxPage then%><a href="/prodotti.asp?p=<%=Page%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=<%=order%>"><%=page%></a>&nbsp;<%end if%>
                                  <% if page >= prod_rs.PageCount then
                                     page = p+4
                                     'exit for
                                    end if
                                  %>
                                  <% next %>
                                  <% 'if page-1 <= prod_rs.PageCount then %>
                                  <% if cInt(p) < maxPage then %>
                                  [<a href="/prodotti.asp?p=<%=p+1%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=<%=order%>">succ &gt;</a>] 
                                  <% end if %>
                                  <% if prod_rs.PageCount-page > 3 then %>
                                  [<a href="/prodotti.asp?p=<%=p+5%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=<%=order%>">5 succ &gt;&gt;</a>] 
                                  <% end if%>
                                  <%if maxPage>5 and cInt(p)<>prod_rs.PageCount then%>[<a href="/prodotti.asp?p=<%=prod_rs.PageCount%>&cat=<%=cat%>&FkProduttore=<%=FkProduttore%>&order=<%=order%>">ultima pag.</a>]<%end if%>	
                                    </li>
                                  <%end if%>
                                  <!--fine paginazione-->
                                     
                                </ul>
                                <%else%>
                                	<%if cat>0 then%>
                                    <p>Nessun prodotto per la categoria selezionata</p>
                                   	<%end if%>
									<%if FkProduttore>0 then%>
                                    <p>Al momento non sono esposti sul sito internet prodotti di <%=titolo_Produttore%>, ma abbiamo comunque a disposizione il loro catalogo e vendiamo i loro prodotti nel nostro negozio. <br />Quindi, se conosci un articolo di questo produttore e vuoi avere un preventivo di prezzo riempi il modulo indicandoci il prodotto, 
		  verrai contattato il prima possibile: il nostro staff sar&agrave; a Tua disposizione per qualsiasi chiarimento.<br />
	      <br /><a href="#" onClick="MM_openBrWindow('/richiesta_informazioni_produttore.asp?produttore=<%=titolo_produttore%>&amp;id=<%=FkProduttore%>','','scrollbars=yes,width=650,height=650')">Clicca qui per aprire il modulo per la richiesta informazioni e preventivi</a></p>
          							<%end if%>
                               <%	
								end if
								prod_rs.close
							   %>
                             <%
							 else
							 'prodotti in vetrina
							 %>
                             	<h1>Articoli illuminazione in primo piano</h1>
                                <p>
                                    <i>
                                    Questa &eacute; una breve selezione di articoli di illuminazione che rappresentano la nostra galleria.<br />
					    			Per consultare tutto il catalogo ed accedere ai singoli prodotti, potete scegliere una categoria sulla sinistra.<br />
									Ogni articolo ha una propria scheda dettagliata, per accederci &eacute; sufficiente cliccare o sul nome, o sulla fotografia oppure su "Scheda prodotto".
                                    </i>
                                </p>
                                
                                <%
								Set prod_rs = Server.CreateObject("ADODB.Recordset")
								sql = "SELECT * FROM Prodotti WHERE (PrimoPiano=True And (Offerta=0 or Offerta=2)) ORDER BY PrezzoProdotto ASC"
								prod_rs.open sql,conn, 1, 1
								
								Randomize()
								constnum = 5
								
								if prod_rs.recordcount>0 then
								%>
                                <ul class="prodotti clearfix">
                                <%
								IF NOT prod_rs.EOF THEN
									rndArray = prod_rs.GetRows()
									prod_rs.Close
							
									Lenarray =  UBOUND( rndArray, 2 ) + 1
									skip =  Lenarray  / constnum 
									IF Lenarray <= constnum THEN skip = 1
									FOR i = 0 TO Lenarray - 1 STEP skip
											numero = RND * ( skip - 1 )
											id = rndArray( 0, i + numero )
											codicearticolo = rndArray( 1, i + numero )
											titolo_prodotto = rndArray( 4, i + numero )
											descrizione_prodotto = NoHTML(rndArray( 5, i + numero ))
											allegato_prodotto = rndArray( 6, i + numero )
											prezzoarticolo = rndArray( 7, i + numero )
											prezzolistino = rndArray( 8, i+ numero )
											
											fkproduttore = rndArray( 11, i+ numero )
											if fkproduttore="" then fkproduttore=0
											
											NomePagina = rndArray( 13, i+ numero )
											if NomePagina="" then NomePagina="#"
											if NomePagina<>"#" then NomePagina="/public/pagine/"&NomePagina
										
											if fkproduttore>0 then
												Set pr_rs = Server.CreateObject("ADODB.Recordset")
												sql = "SELECT * FROM Produttori WHERE PkId="&fkproduttore&""
												pr_rs.open sql,conn, 1, 1
												if pr_rs.recordcount>0 then
													produttore=pr_rs("titolo")
												end if
												pr_rs.close
											end if
											
											FkCategoria2 = rndArray( 12, i+ numero )
											if FkCategoria2="" then FkCategoria2=0
											
											if FkCategoria2>0 then
												Set cat_rs = Server.CreateObject("ADODB.Recordset")
												sql = "SELECT Categorie1.PkId as Cat_Principale, Categorie1.Titolo as Titolo1, Categorie2.PkId, Categorie2.Titolo as Titolo2, Categorie2.Descrizione as Descrizione2 "
												sql = sql + "FROM Categorie1 INNER JOIN Categorie2 ON Categorie1.PkId = Categorie2.FkCategoria1 "
												sql = sql + "WHERE Categorie2.PkId="&FkCategoria2
												cat_rs.open sql,conn, 1, 1
												if cat_rs.recordcount>0 then
													'cat_principale=cat_rs("Cat_Principale")
													titolo_cat=cat_rs("titolo1")&" - "&cat_rs("titolo2")
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
											'file_img="logo_cristalensi_piccolo.jpg"
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
                                        
                                        	<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/<%=file_img%>" alt="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>" style="width: <%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>px; height: <%if H<=120 then%><%=H%><%else%>120<%end if%>px;" border="0"></a>
										<%else%>
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"></a>	
										<%
                                            end if
                                        else
                                            tot_img=0
                                            titolo_img=""
                                            file_img=""
                                        %>
                                    		<a href="<%=NomePagina%>" style="display: block;" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="/public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"></a>
										<%	
                                        end if
                                        img_rs.close
                                        %>
                                        <%if tot_img>0 then%><span>[<%if tot_img=1 then%>1 Immagine<%else%><%=tot_img%> Immagini<%end if%>]</span><%end if%>
                                        </div>
                                        <div class="data">
                                            <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%> - <%=titolo_cat%>"><strong><%=titolo_prodotto%></strong><%if codicearticolo<>"" then%>&nbsp;[<%=codicearticolo%>]<%end if%></a> <%if fkproduttore>0 then%><span class="produttore">Produttore: <a href="/prodotti.asp?FkProduttore=<%=fkproduttore%>" title="Elenco prodotti dello stesso produttore: <%=produttore%>"><strong><%=produttore%></strong></a></span><%end if%>
                                            <p><%=Left(descrizione_prodotto,150)%><%if Len(descrizione_prodotto)>150 then%>...<%end if%><%if FkCategoria2>0 then%></p><p><i>Categoria:</i> <a href="prodotti.asp?cat=<%=FkCategoria2%>" title="Elenco prodotti della stessa categoria: <%=titolo_cat%>" style="font-size:9px;"><%=titolo_cat%></a><%end if%></p>
                                            <a href="<%=NomePagina%>" title="Scheda del prodotto&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="button_link scheda-link"><span>Scheda prodotto</span></a>
                                            <%if prezzoarticolo=0 then%>
                                            <p class="cart clearfix"><span class="price">Prezzo listino: <span><%=prezzolistino%>&#8364;</span></span>&nbsp;&nbsp;<a href="#" onClick="MM_openBrWindow('/richiesta_informazioni.asp?codice=<%=codicearticolo%>&titolo=<%=titolo_prodotto%>&amp;produttore=<%=produttore%>&amp;id=<%=id%>','','width=650,height=650,scrollbars=yes')" class="cart-link button_link_red">Prezzo Cristalensi? clicca qui per un preventivo dal nostro staff</a></p>
                                            <%else%>
                                            <p class="cart clearfix"><%if prezzolistino<>0 then%><span class="price">Prezzo listino: <span><%=prezzolistino%>&#8364;</span></span><%end if%>&nbsp;&nbsp;<%if prezzoarticolo<>"" then%><span class="cristalprice">Prezzo Cristalensi: <%=prezzoarticolo%>&#8364;&nbsp;&nbsp;<small><i>Iva compresa</i></small></span><%end if%><a href="<%=NomePagina%>" title="Inserisci&nbsp;nel&nbsp;carrello&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="cart-link button_link_red"><span>Inserisci nel carrello</span></a></p>
                                            <%end if%>
                                        </div>
                                    </li>
                                    <% 
											NEXT
											end if
										else
											prod_rs.close
										end if
										
									%>
                                </ul>
                             <%end if%>
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