<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_clsImageSize.asp"-->
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
    </head>
    <body>
        <div id="wrap">
            <!--#include file="inc_header.asp"-->
            <div id="main-content">
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <img class="negozio" src="images/negozio.jpg">
                            <img class="anni" src="images/50anni_new.jpg">
                            <h3>Cristalensi, la luce come idea</h3>
                            <p class="incipit">A portata di click una vasta e raffinata gamma di prodotti per illuminazione da interno ed esterno per arredare la vostra casa, il giardino, il tuo locale... Naviga nel Negozio on-line oppure visita il nostro Showroom, soddisferemo tutte le tue esigenze sia classiche che moderne.
                            </p>
                            <div class="social">
                                <img class="facebook" src="images/facebook.png">
                                <p style="padding-top:10px; line-height: 160%;">Seguici su FACEBOOK collegandoti alla pagina ufficiale di Cristalensi, troverai le ultime novità e bellissime fotografie di arredamento da condividere con i tuoi amici, inoltre troverai altri appassionati di illuminazione per parlare dei nostri articoli</p>
                            </div>
                            <div class="social_panel facebook_p">Facebook place</div>
                            <div class="social_panel twitter_p">Twitter place</div>
                            <div class="slogan">
                                <h3>Eccezionale sconto!!! Nessun costo di spedizione per ordini superiori a 250€</h3>
                                <p>Per ordini inferiori a 250€ il costo di spedizione è di 10€.<br> Condizioni valide solo per le spedizioni in tutta Italia, isole comprese.</p>
                            </div>
                            <!--prodotti in offerta-->
                            <h4 class="area">OFFERTE: non perdere l'occasione!<a href="offerte.asp" style="float: right; padding: 1px 10px;" class="button_link_red">TUTTI I PRODOTTI IN OFFERTA &raquo;</a></h4>
                            <%
							'random prodotti in offerta
							Set prod_rs = Server.CreateObject("ADODB.Recordset")
							sql = "SELECT pkid,codicearticolo,titolo,prezzoprodotto,prezzolistino,nomepagina,offerta FROM Prodotti WHERE Offerta=1 OR Offerta=2 ORDER BY Titolo ASC"
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
									if NomePagina="" then NomePagina="#"
									'if NomePagina<>"#" then NomePagina="public/pagine/"&NomePagina
									if NomePagina<>"#" then NomePagina="scheda_prodotto.asp?pkid="&id
									
									
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
										   '.ImageFile = server.mappath("public/"&file_img&"")
										  .ImageFile = path_img&file_img
										  
										  If .IsImage Then
											W=.ImageWidth
											H=.ImageHeight
										  End If 
										  
										End With
										Set objImageSize = Nothing
							%>
                            	<li>
                                    <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>"><img src="public/<%=file_img%>" alt="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>" width="<%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>" height="<%if H<=120 then%><%=H%><%else%>120<%end if%>" border="0"><%=titolo_prodotto%><%if codicearticolo<>"" then%>&nbsp;[<%=codicearticolo%>]<%end if%></a>
										<%else%>
                                    <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>"><img src="public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"><%=titolo_prodotto%><%if codicearticolo<>"" then%>&nbsp;[<%=codicearticolo%>]<%end if%></a>	
                                    <%
                                        end if
                                    else
                                        tot_img=0
                                        titolo_img=""
                                        file_img=""
                                    %>
                                    <a href="<%=NomePagina%>" title="<%=titolo_prodotto%>"><img src="public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"><%=titolo_prodotto%><%if codicearticolo<>"" then%>&nbsp;[<%=codicearticolo%>]<%end if%></a>
                                    <%	
                                    end if
                                    img_rs.close
                                    %>
                                    <%if prezzolistino<>"" then%><p class="price">Prezzo listino: <span><%=prezzolistino%>€</span></p><%end if%>
                                    <%if prezzoarticolo<>"" then%><p class="cristalprice">Prezzo Cristalensi: <%=prezzoarticolo%>€</p><%end if%>
                                    <a class="scheda" href="<%=NomePagina%>" title="Scheda del prodotto <%=titolo_prodotto%>">Scheda prodotto</a>
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
                            <h4 class="area">CATALOGO PRODOTTI <a href="#" style="float: right; padding: 1px 10px;" class="button_link_red">RICERCA AVANZATA &raquo;</a></h4>
                            <!--<p>Ricerca il prodotto desiderato usando la divisione in categorie oppure la <button>RICERCA AVANZATA</button>-->
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
								titolo_cat=prod_rs("Titolo")
								nomepagina_categorie=prod_rs("NomePagina")
								if nomepagina_categorie="" then nomepagina_categorie="#"
								'if nomepagina_categorie<>"#" then nomepagina_categorie="public/pagine/"&nomepagina_categorie
								if nomepagina_categorie<>"#" then nomepagina_categorie="categorie.asp?pkid="&cat
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
									<a href="<%=nomepagina_categorie%>" title="Elenco articoli <%=titolo_cat%>"><img src="public/<%=file_img%>" width="160" height="120" vspace="2" border="0" alt="<%=titolo_cat%>"><%=titolo_cat%></a>
										<%else%>
									<a href="<%=nomepagina_categorie%>" title="Elenco articoli <%=titolo_cat%>"><img src="immagini/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine della categoria <%=titolo_cat%> non disponibile"><%=titolo_cat%></a>	
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
                            <h4 class="area">PRODUTTORI<a href="produttori.asp" style="float: right;" title="Elenco completo dei produttori di articoli per illuminazione">ELENCO COMPLETO PRODUTTORI &raquo;</a></h4>
                            <p>Se conosci la marca del prodotto la puoi selezionare qui sotto oppure andando all'elenco completo dei produttori.
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
                            <option value="0">Seleziona un produttore</option>
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
        <script src="js/init.js"></script>
    </body>
</html>
<!--#include file="inc_strClose.asp"-->