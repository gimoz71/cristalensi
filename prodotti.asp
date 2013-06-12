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
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title><%if cat>0 then%><%=nuovo_title_cat%><%end if%><%if FkProduttore>0 then%><%=titolo_produttore%> catalogo prodotti illuminazione vendita online Cristalensi<%end if%><%if cat=0 and FkProduttore=0 then%>catalogo prodotti illuminazione da interno illuminazione da esterno vendita online Cristalesni<%end if%></title>
		<meta name="description" content="<%if cat>0 then%><%=NoHTML(descrizione_cat)%><%end if%><%if FkProduttore>0 then%>Catalogo prodotti in vendita di <%=titolo_produttore%>, vendita online prodotti illuminazione su Cristalensi<%end if%><%if cat=0 and FkProduttore=0 then%>catalogo prodotti illuminazione da interno, illuminazione da esterno, vendita online su Cristalesni<%end if%>">
		<meta name="keywords" content="<%if cat>0 then%><%=nuovo_title_cat%><%end if%><%if FkProduttore>0 then%><%=titolo_produttore%> catalogo prodotti illuminazione vendita online Cristalensi<%end if%><%if cat=0 and FkProduttore=0 then%>catalogo prodotti illuminazione da interno illuminazione da esterno vendita online Cristalesni<%end if%>">
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
                            <div class="slogan">
                                <h3>Eccezionale sconto!!! Nessun costo di spedizione per ordini superiori a 250€</h3>
                                <p>Per ordini inferiori a 250€ il costo di spedizione è di 10€.<br> Condizioni valide solo per le spedizioni in tutta Italia, isole comprese.</p>
                            </div>
                            <%
							'elenco prodotti di una categoria o di un produttore
							if cat>0 or FkProduttore>0 then%>
                                <h3>Catalogo prodotti: <%if cat>0 then%><%=titolo_cat%><%end if%><%if FkProduttore>0 then%><%=titolo_produttore%><%end if%></h3>
                                <%if descrizione_cat<>"" then%>
                                <p>
                                    <i><%=descrizione_cat%></i>.
                                </p>
                                <%end if%>
                                <ul class="prodotti clearfix">
                                    <li class="clearfix">
                                        <div class="thumb">
                                        <a href="#">
                                            <img src="images/example.jpg">
                                        </a>
                                        </div>
                                        <div class="data">
                                            <a href="scheda.html"><strong>Applique di vetro murano</strong></a> <span class="produttore">produttore: <a href="#"><strong>Illumnando</strong></a></span>
                                            <p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vel mauris vitae urna ornare convallis a eu elit. Nulla varius lobortis lorem a molestie...</p>
                                            <a class="button_link" href="scheda.html">Scheda prodotto</a>
                                            <p class="cart clearfix"><span class="price">Prezzo listino: <span>155€</span></span> <span class="cristalprice">Prezzo listino: 155€</span><a href="#" class="cart-link button_link"><span>Inserisci nel carrello</span></a></p>
                                        </div>
                                    </li>
                                    
                                </ul>
                             <%
							 else
							 'prodotti in vetrina
							 %>
                             	<h3>Prodotti in vetrina</h3>
                                <p>
                                    <i>
                                    Questa è una breve selezione di prodotti che rappresentano la nostra galleria.<br />
					    			Per consultare tutto il catalogo ed accedere ai singoli prodotti, potete scegliere una categoria sulla sinistra.<br />
									Ogni prodotto ha una propria scheda dettagliata, per accederci è sufficiente cliccare sul nome o sulla foto del prodotto.
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
                                        
                                        	<a href="<%=NomePagina%>" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="public/<%=file_img%>" alt="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>" width="<%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>" height="<%if H<=120 then%><%=H%><%else%>120<%end if%>" border="0"></a>
										<%else%>
                                    		<a href="<%=NomePagina%>" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"></a>	
										<%
                                            end if
                                        else
                                            tot_img=0
                                            titolo_img=""
                                            file_img=""
                                        %>
                                    		<a href="<%=NomePagina%>" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><img src="public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"></a>
										<%	
                                        end if
                                        img_rs.close
                                        %>
                                        </div>
                                        <div class="data">
                                            <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%> - <%=titolo_cat%>"><strong><%=titolo_prodotto%></strong><%if codicearticolo<>"" then%>&nbsp;[<%=codicearticolo%>]<%end if%></a> <%if fkproduttore>0 then%><span class="produttore">Produttore: <a href="prodotti.asp?FkProduttore=<%=fkproduttore%>" title="Elenco prodotti dello stesso produttore: <%=produttore%>"><strong><%=produttore%></strong></a></span><%end if%>
                                            <p><%=Left(descrizione_prodotto,150)%><%if Len(descrizione_prodotto)>150 then%>...<%end if%><%if FkCategoria2>0 then%><br /><i>Il prodotto lo trovi nella categoria:</i> <a href="prodotti.asp?cat=<%=FkCategoria2%>" title="Elenco prodotti della stessa categoria: <%=titolo_cat%>"><%=titolo_cat%></a><%end if%></p>
                                            <a href="<%=NomePagina%>" title="Scheda del prodotto&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="button_link">Scheda prodotto</a>
											<%if tot_img>0 then%><span style="float:right;">[<%if tot_img=1 then%>1 Immagine<%else%><%=tot_img%> Immagini<%end if%>]</span><%end if%>
                                            <%if prezzoarticolo=0 then%>
                                            <p class="cart clearfix"><span class="price">Prezzo listino: <span><%=prezzolistino%>€</span></span>&nbsp;&nbsp;&nbsp;<span class="cristalprice"><a href="#" onClick="MM_openBrWindow('richiesta_informazioni.asp?codice=<%=codicearticolo%>&titolo=<%=titolo_prodotto%>&amp;produttore=<%=produttore%>&amp;id=<%=id%>','','width=650,height=650,scrollbars=yes')">Prezzo Cristalensi? clicca qui per un preventivo dal nostro staff</a></span></p>
                                            <%else%>
                                            <p class="cart clearfix"><%if prezzolistino<>0 then%><span class="price">Prezzo listino: <span><%=prezzolistino%>€</span></span><%end if%>&nbsp;&nbsp;&nbsp;<%if prezzoarticolo<>"" then%><span class="cristalprice">Prezzo Cristalensi: <%=prezzoarticolo%>€</span><%end if%>&nbsp;&nbsp;&nbsp;<i>Iva compresa</i><a href="<%=NomePagina%>" title="Inserisci&nbsp;nel&nbsp;carrello&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="cart-link button_link"><span>Inserisci nel carrello</span></a></p>
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
        <script src="js/init.js"></script>
    </body>
</html>
