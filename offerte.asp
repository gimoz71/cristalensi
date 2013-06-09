<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_clsImageSize.asp"-->
<%
	order=request("order")
	if order="" then order=3
	if order=1 then ordine="Titolo ASC"
	if order=2 then ordine="Titolo DESC"
	if order=3 then ordine="prezzoprodotto ASC"
	if order=4 then ordine="prezzoprodotto DESC"	
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
                                <p>Per ordini inferiori a 250&#8364; il costo di spedizione &egrave; di 10&#8364;.<br /> Condizioni valide solo per le spedizioni in tutta Italia, isole comprese.</p>
                            </div>
                            <h4>Prodotti in offerta</h4>
                            <p><em>In questa pagina trovate tutte le offerte di prodotti per illuminazione: sono gli articoli del catalogo con prezzi fantastici. Ogni prodotto ha una propria scheda dettagliata, per accederci &egrave; sufficiente cliccare sul nome o sulla foto dell'articolo.<br />
					    Invece, per consultare tutto il catalogo potete cliccare qui su <a href="prodotti.asp" title="Catalogo prodotti per illuminazione">[Prodotti]</a> (oppure sullla stessa voce del men&ugrave; in alto) ma potete anche scegliere una categoria o un produttore dal men&ugrave; sulla sinistra.</em>
                          </p>
                            <p class="area"> <strong>Ordinamento per prezzo:</strong> 
                            <a href="offerte.asp?order=3"><img src="images/01_new<%if order=3 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" hspace="3" border="0" align="top" alt="ordina i prodotti per prezzo dal pi&ugrave; basso al pi&ugrave; alto" title="ordina i prodotti per prezzo dal pi&ugrave; basso al pi&ugrave; alto" /></a>
                            <a href="offerte.asp?order=4"><img src="images/10_new<%if order=4 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px" border="0" align="top" alt="ordina i prodotti per prezzo dal pi&ugrave; alto al pi&ugrave; basso" title="ordina i prodotti per prezzo dal pi&ugrave; alto al pi&ugrave; basso" /></a>
                            &nbsp;-&nbsp;
                            <strong>Ordinamento per nome:</strong>
                            <a href="offerte.asp?order=1"><img src="images/az_new<%if order=1 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px"  hspace="3" border="0" align="top" alt="ordina i prodotti per titolo dalla A alla Z" title="ordina i prodotti per titolo dalla A alla Z" /></a>&nbsp;
                            <a href="offerte.asp?order=2"><a href="offerte.asp?order=2"><img src="images/za_new<%if order=2 then%>_sott<%end if%>.gif" style="float: none;width: 22px; height: 15px"  border="0" align="top" alt="ordina i prodotti per titolo dalla Z alla A" title="ordina i prodotti per titolo dalla Z alla A" /></a></a></p>
                          <p>Le nostre offerte in vetrina a prezzi scontati. Consulta tutte le offerte nell'apposita sezione "Prodotti in offerta"<br>Non perdere l'occasione!!!</p>
                         
                          <ul class="prodotti clearfix">
                                
                            <%
                            Set prod_rs = Server.CreateObject("ADODB.Recordset")
                            sql = "SELECT * FROM Prodotti WHERE Offerta=1 or Offerta=2 ORDER BY "&ordine&""
                            prod_rs.open sql,conn, 1, 1
                            if prod_rs.recordcount>0 then

                                    Do while not prod_rs.EOF

                                            id=prod_rs("pkid")
                                            titolo_prodotto=prod_rs("titolo")

                                            NomePagina_prodotto=prod_rs("NomePagina")			
                                            if NomePagina_prodotto="" then NomePagina_prodotto="#"
                                            'if NomePagina_prodotto<>"#" then NomePagina_prodotto="public/pagine/"&NomePagina_prodotto
                                            if NomePagina_prodotto<>"#" then NomePagina_prodotto="scheda_prodotto.asp?pkid="&id

                                            codicearticolo=prod_rs("codicearticolo")
                                            descrizione_prodotto=prod_rs("descrizione")
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
                                    <a href="<%=NomePagina_prodotto%>" title="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>"><img src="public/<%=file_img%>" alt="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>" width="<%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>" height="<%if H<=120 then%><%=H%><%else%>120<%end if%>" hspace="2" vspace="2" border="0"></a>
                                            <%else%>
                                    <a href="<%=NomePagina_prodotto%>" title="<%=titolo_prodotto%>"><img src="public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"></a>	
                                    <%
                                            end if
                                    else
                                            tot_img=0
                                            titolo_img=""
                                            file_img=""
                                    %>
                                    <a href="<%=NomePagina_prodotto%>" title="<%=titolo_prodotto%>"><img src="public/logo_cristalensi_piccolo.jpg" width="120" height="90" vspace="2" border="0" alt="immagine del prodotto <%=titolo_prodotto%> non disponibile"></a>
                                    <%	
                                    end if
                                    img_rs.close
                                    %>
                                    </div>
                                    <div class="data">
                                        <a href="<%=NomePagina_prodotto%>" title="<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>"><%=titolo_prodotto%><%if codicearticolo<>"" then%>&nbsp;[<%=codicearticolo%>]<%end if%></a> <%if fkproduttore>0 then%><span class="produttore">Produttore: <a href="prodotti.asp?FkProduttore=<%=fkproduttore%>" title="Elenco prodotti <%=produttore%>"><%=produttore%></a></span><%end if%>
                                        <p><%=Left(descrizione_prodotto,100)%><%if Len(descrizione_prodotto)>100 then%>...<%end if%><%if FkCategoria2>0 then%>&nbsp;&nbsp;Il prodotto lo trovi nella categoria: <a href="prodotti.asp?cat=<%=FkCategoria2%>" title="Elenco <%=titolo_cat%>"><%=titolo_cat%></a><%end if%></p>
                                        <a href="<%=NomePagina_prodotto%>" title="<%=titolo_prodotto%>&nbsp;<%=titolo_cat%>">Scheda del prodotto</a><%if tot_img>0 then%>[<img src="images/img.jpg" border="0" width="18" height="18" hspace="3" align="absmiddle" alt="Sono presenti altre immagini"><%if tot_img=1 then%>1 Immagine<%else%><%=tot_img%> Immagini<%end if%>]<%end if%>
                                        <%if allegato_prodotto<>"" then%>
                                        <img src="images/file.jpg" border="0" width="18" height="18" hspace="3" align="absmiddle" alt="E' presente un allegato">Allegato
                                        <%end if%>
                                        <%if prezzoarticolo=0 then%>
                                        	<p class="cart clearfix"><span class="price">Prezzo listino: <span><%=prezzolistino%>&#8364;</span></span> <a href="#" onClick="MM_openBrWindow('richiesta_informazioni.asp?codice=<%=codicearticolo%>&titolo=<%=titolo_prodotto%>&amp;produttore=<%=produttore%>&amp;id=<%=id%>','','width=650,height=650,scrollbars=yes')" class="cart-link">Prezzo Cristalensi? clicca qui per un preventivo dal nostro staff</a></p>
                                        <%else%>
                                        	<p class="cart clearfix"><%if prezzolistino<>0 then%><span class="price">Prezzo listino: <span><%=prezzolistino%>&#8364;</span></span><%end if%> <%if prezzoarticolo<>"" then%><span class="cristalprice">Prezzo listino: <%=prezzoarticolo%>&#8364;</span><%end if%><a href="<%=NomePagina_prodotto%>" title="Inserisci nel carrello" class="cart-link">Inserisci nel carrello</a></p>
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