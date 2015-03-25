<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_clsImageSize.asp"-->
<%
cat=request("cat")				  
if cat="" then cat=0

		titolo_cat=""
		title_cat=""
		descrizione_cat=""
%>
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title><%=titolo_cat%></title>
		<meta name="description" content="<%=descrizione_cat%> catalogo prodotti illuminazione da interno, illuminazione da esterno, vendita online su Cristalesni">
		<meta name="keywords" content="<%=title_cat%> catalogo prodotti illuminazione vendita online Cristalensi">
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="js/media-queries-ie.js"></script>
        <![endif]-->
        <<link href="/css/css.css" rel="stylesheet" type="text/css">
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
            <link href="css/tipTip_ie7.css" media="all" rel="stylesheet" type="text/css" />
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
                                <h3>Eccezionale sconto!!! Nessun costo di spedizione per ordini superiori a 250&#8364;</h3>
                                <p>Per ordini inferiori a 250&#8364; il costo di spedizione &eacute; di 10&#8364;.<br> Condizioni valide solo per le spedizioni in tutta Italia, isole comprese.</p>
                            </div>
                                <h1>Catalogo prodotti: <%=titolo_cat%></h1>
                                <p>
                                    <i><%=descrizione_cat%></i>
                                </p>
                                <p>&nbsp;</p>
                                
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
                                          Non hai trovato il prodotto che cercavi?<br />scegli una categoria:
                                          <%
                                          Set cs=Server.CreateObject("ADODB.Recordset")
                                          sql = "SELECT Categorie1.PkId as PkId_1, Categorie1.Titolo as Titolo_1, Categorie2.PkId as PkId_2, Categorie2.Titolo as Titolo_2 "
                                          sql = sql + "FROM Categorie1 INNER JOIN Categorie2 ON Categorie1.PkId = Categorie2.Fkcategoria1 "
                                          'sql = sql + "WHERE Categorie2.FkCategoria1 = "&cat&" "
                                          sql = sql + "ORDER BY Categorie1.Titolo ASC, Categorie2.Titolo ASC"
                                          cs.Open sql, conn, 1, 1
                                          %>
                                          <select name="Cat" id="Cat" class="form" onChange="invia_account()" style="margin-top:10px;">
                                              <option title="Scegli una categoria" value="0">Scegli una categoria</option>
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
                                        <p>Oppure, per una ricerca maggiormente dettagliata, usa la<br>
                                        <span><a href="/ricerca_avanzata_modulo.asp" class="button_link_red" style="margin-top:7px;">RICERCA AVANZATA</a></span>
                                        </p>
                                    </div>
                                    <div class="clear"></div>
                                    <br /><br />
                              <%
								
									Set prod_rs = Server.CreateObject("ADODB.Recordset")
									if cat>0 then sql = "SELECT * FROM Prodotti WHERE (FkCategoria2="&cat&" and (Offerta=0 or Offerta=2)) ORDER BY PkId DESC"
									prod_rs.open sql,conn, 1, 1
									if prod_rs.recordcount>0 then
								
									prod_rs.PageSize = 30
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
                                            <p class="cart clearfix"><%if prezzolistino<>0 then%><span class="price">Prezzo listino: <span><%=prezzolistino%>&#8364;</span></span><%end if%>&nbsp;&nbsp;<%if prezzoarticolo<>"" then%><span class="cristalprice">Prezzo Cristalensi: <%=prezzoarticolo%>&#8364;</span>&nbsp;&nbsp;<small><i>Iva compresa</i></small><%end if%><a href="<%=NomePagina%>" title="Inserisci&nbsp;nel&nbsp;carrello&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="cart-link button_link_red"><span>Inserisci nel carrello</span></a></p>
                                            <%end if%>
                                        </div>
                                        
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
                <!--#include file="inc_sx_prodotti.asp"-->
            </div>
        </div>
        <!--#include file="inc_footer.asp"-->
    </body>
</html>
<!--#include file="inc_strClose.asp"-->