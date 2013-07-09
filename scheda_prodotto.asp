<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_clsImageSize.asp"-->
<%
id=request("id")
if id="" then id=0
if id=0 then response.Redirect("prodotti.asp")

if id>0 then
	Set prod_rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Prodotti WHERE PKId="&id
	prod_rs.open sql,conn, 3, 3
	if prod_rs.recordcount>0 then
		CodiceArticolo=prod_rs("CodiceArticolo")
		'FkCat_Prod=prod_rs("FkCat_Prod")
		Titolo_prodotto=prod_rs("Titolo")
		Descrizione_prodotto=prod_rs("Descrizione")
		allegato_prodotto=prod_rs("Allegato")
		PrezzoArticolo=prod_rs("PrezzoProdotto")
		PrezzoListino=prod_rs("PrezzoListino")
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
				descrizione_cat=cat_rs("Descrizione2")
			end if
			cat_rs.close
		end if
		
		'aggiorno il contatore
		visualizzazioni=prod_rs("visualizzazioni")
		if visualizzazioni="" or IsNull(visualizzazioni) then visualizzazioni=0
		prod_rs("visualizzazioni")=visualizzazioni+1
		prod_rs.update
	end if
	prod_rs.close
	
	
	'Call Visualizzazione("Prodotti",id,"scheda_prodotto.asp")
end if
%>
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title><%=Titolo_prodotto%> - <%=titolo_cat%> - <%=codicearticolo%></title>
		<meta name="description" content="Cristalensi vende <%=titolo_cat%>: <%=Titolo_prodotto%> - <%=codicearticolo%>">
		<meta name="keywords" content="<%=Titolo_prodotto%>, <%=Titolo_prodotto%> <%=titolo_cat%>, <%=Titolo_prodotto%> <%=codicearticolo%>">
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="js/media-queries-ie.js"></script>
        <![endif]-->
        <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
        <script src="js/jquery.blueberry.js"></script>
        <script src="js/jquery.tipTip.js"></script>
        <script src="js/jquery.fancybox.js"></script>
        <link href="css/css.css" rel="stylesheet" type="text/css">
        <link href="css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="css/jquery.fancybox.css" rel="stylesheet" type="text/css">
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
    	<SCRIPT language="JavaScript">
            $(document).ready(function() {
                $('.fancybox').fancybox();
            });
		function verifica_1() {
				
			quantita=document.newsform2.quantita.value;
			num_colori=document.newsform2.num_colori.value;
			colore=document.newsform2.colore.value;
		
			if (quantita=="0"){
				alert("La quantita\' deve essere maggiore di 0");
				return false;
			}
			
			if (num_colori>1 && colore==""){
				alert("Deve essere scelto un colore");
				return false;
			}
			
			else
				
				document.newsform2.method = "post";
				//document.newsform2.action = "../../carrello1.asp";
				document.newsform2.action = "carrello1.asp";
				document.newsform2.submit();
		}
		</SCRIPT>
		<SCRIPT language="JavaScript">
		function verifica_2() {
				
			quantita=document.newsform2.quantita.value;
			num_colori=document.newsform2.num_colori.value;
			colore=document.newsform2.colore.value;
		
			if (quantita=="0"){
				alert("La quantita\' deve essere maggiore di 0");
				return false;
			}
			
			if (num_colori>1 && colore==""){
				alert("Deve essere scelto un colore");
				return false;
			}
			
			else
				
				document.newsform2.method = "post";
				//document.newsform2.action = "../../carrello1.asp";
				document.newsform2.action = "carrello1.asp";
				//document.newsform2.submit();
		}
		</SCRIPT>
    
    </head>
    <body>
    <!--facebook-->
    <div id="fb-root"></div>
	<script>(function(d, s, id) {
      var js, fjs = d.getElementsByTagName(s)[0];
      if (d.getElementById(id)) return;
      js = d.createElement(s); js.id = id;
      js.src = "//connect.facebook.net/it_IT/all.js#xfbml=1";
      fjs.parentNode.insertBefore(js, fjs);
    }(document, 'script', 'facebook-jssdk'));</script>
    <!--facebook-->
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

                            <ul class="scheda-prodotto clearfix">
                                <li class="clearfix">
                                    <a href="javascript:history.back()" class="cart-link button_link">Torna indietro</a>
                                    <h3><%=Titolo_prodotto%> - <%=codicearticolo%></h3>
                                    <p class="area clearfix"><%if codicearticolo<>"" then%>Codice articolo <strong>[<%=codicearticolo%>]</strong><%end if%><%if fkproduttore>0 then%><span class="produttore">produttore: <a href="/prodotti.asp?FkProduttore=<%=fkproduttore%>" title="Elenco prodotti dello stesso produttore: <%=produttore%>"><strong><%=produttore%></strong></a></span><%end if%></p>
                                    <div class="data">
                                        <%if prezzoarticolo=0 then%>
                                           <p class="cart-panel clearfix"s tyle="float: right; width: 30%;  text-align: center;"><span class="price">Prezzo listino: <span><%=prezzolistino%>€</span></span>&nbsp;&nbsp;<span class="cristalprice"><a href="#" onClick="MM_openBrWindow('../../richiesta_informazioni.asp?codice=<%=codicearticolo%>&titolo=<%=titolo_prodotto%>&amp;produttore=<%=produttore%>&amp;id=<%=id%>','','width=650,height=650,scrollbars=yes')" class="cart-link">Vuoi sapere il prezzo Cristalensi? clicca qui per avere un preventivo dal nostro staff</a></span>
                                       <%else%>
                                           <p class="cart-panel clearfix" style="float: right; width: 30%;  text-align: center;"><%if prezzolistino<>0 then%><span class="price">Prezzo listino: <span><%=prezzolistino%>€</span></span><%end if%><br><%if prezzoarticolo<>"" then%><span class="cristalprice">Prezzo Cristalensi: <%=prezzoarticolo%>€</span><%end if%><br><i>Iva compresa</i>
                                       <%end if%>
                                        <p><%=descrizione_prodotto%></p>
                                        <%if FkCategoria2>0 then%>
                                            <p> Il prodotto lo trovi nella categoria: <a href="/prodotti.asp?cat=<%=FkCategoria2%>" title="Elenco prodotti della stessa categoria: <%=titolo_cat%>"><%=titolo_cat%></a></p>
					<%end if%>
                                        <%if allegato_prodotto<>"" then%>
                                            <p><a href="/public/<%=allegato_prodotto%>" target="_blank"><img src="/images/file.jpg" border="0" width="18" height="18" hspace="3" align="absmiddle" alt="E' presente un allegato">Allegato</a></p>
                                        <%end if%>
                                        <%if prezzoarticolo=0 then%>
                                            <p class="cart clearfix"><span class="price">Prezzo listino: <span><%=prezzolistino%>€</span></span>&nbsp;&nbsp;<span class="cristalprice"><a href="#" onClick="MM_openBrWindow('../../richiesta_informazioni.asp?codice=<%=codicearticolo%>&titolo=<%=titolo_prodotto%>&amp;produttore=<%=produttore%>&amp;id=<%=id%>','','width=650,height=650,scrollbars=yes')" class="cart-link">Vuoi sapere il prezzo Cristalensi? clicca qui per avere un preventivo dal nostro staff</a></span>
                                        <%else%>
                                            <form name="newsform2" id="newsform2" onSubmit="return verifica_2();">
                                                <input type="hidden" name="id" id="id" value="<%=id%>">
                                                <%
                                                Set col_rs = Server.CreateObject("ADODB.Recordset")
                                                sql = "SELECT [Prodotto-Colore].FkProdotto, Colori.Titolo FROM [Prodotto-Colore] INNER JOIN Colori ON [Prodotto-Colore].FkColore = Colori.PkId WHERE ((([Prodotto-Colore].FkProdotto)="&id&")) ORDER BY Colori.Titolo ASC"
                                                col_rs.open sql,conn, 1, 1
                                                if col_rs.recordcount>1 then
                                                %>
                                                    <input type="hidden" name="num_colori" id="num_colori" value="<%=col_rs.recordcount%>">
                                                <%else%>
                                                    <input type="hidden" name="num_colori" id="num_colori" value="1">
                                                    <input type="hidden" name="colore" id="colore" value="*****">
                                                <%end if%>
                                                <p class="cart clearfix">

                                                    <%if col_rs.recordcount>1 then%>
                                                        <select name="colore" id="colore" style="width:auto; float:left; margin-top:7px;">
                                                        <option value="">Scegli il colore</option>
                                                        <%
                                                        Do While Not col_rs.EOF
                                                        %>
                                                            <option value="<%=col_rs("Titolo")%>"><%=col_rs("Titolo")%></option>
                                                        <%
                                                        col_rs.movenext
                                                        loop
                                                        %>
                                                        </select>
                                                    <%
                                                    end if
                                                    col_rs.close
                                                    %>
                                                    <a href="#" onClick="return verifica_1();" id="invia_qta_2" rel="nofollow" title="Inserisci&nbsp;nel&nbsp;carrello&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="cart-link button_link_red"><span>Inserisci nel carrello</span></a><span style="float:right; padding-top:7px;"><input type="text" name="quantita" id="quantita" value="0" size="2" style="width:20px; text-align:right; margin-left:5px;">&nbsp;pezzi&nbsp;&nbsp;</span>
                                                </p>
                                            </form>	
                                        <%end if%>
                                    </div>
                                    <%
                                    Set img_rs = Server.CreateObject("ADODB.Recordset")
                                    sql = "SELECT * FROM Immagini WHERE Record="&id&" AND Tabella='Prodotti' Order by PkId ASC"
                                    img_rs.open sql,conn, 1, 1
                                    if img_rs.recordcount>0 then

                                            Do while not img_rs.EOF
                                            titolo_img=img_rs("titolo")
                                            file_img=img_rs("file")
                                            zoom=img_rs("zoom")

                                            if zoom<>"" then
                                                    'percorso_img=server.mappath("public/"&zoom&"")
													percorso_img="public/"&zoom
                                                    'percorso_img=path_img&zoom
                                                    'percorso_img="../"&zoom
                                            else
                                                    'percorso_img=server.mappath("public/"&file_img&"")
													percorso_img="public/"&file_img
                                                    'percorso_img=path_img&file_img
                                                    'percorso_img="../"&file_img
                                            end if
                                            'calcolo misure immagini
                                            Set objImageSize = New ImageSize
                                            With objImageSize
                                              '.ImageFile = server.mappath("public/"&file_img&"")
                                              '.ImageFile = path_img&file_img
                                              .ImageFile = "public/"&file_img

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


                                    <div class="thumb">
                                        <a href="<%=percorso_img%>" class="fancybox" rel="gallery" title="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>">
					<img class="img-border" src="public/<%=file_img%>" width="<%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>" height="<%if H<=120 then%><%=H%><%else%>120<%end if%>" alt="<%if titolo_img<>"" then%><%=titolo_img%>&nbsp;<%=titolo_cat%><%else%><%=titolo_prodotto%>&nbsp;<%=titolo_cat%><%end if%>" title="<%if titolo_img<>"" then%><%=titolo_img%>&nbsp;<%=titolo_cat%><%else%><%=titolo_prodotto%>&nbsp;<%=titolo_cat%><%end if%>" /></a>
                                    </div>
                                    <%
                                    img_rs.movenext
                                    loop
                                    end if
                                    img_rs.close
                                    %>
                                </li>
                                <li class="clearfix">
                                    <img class="facebook" src="images/facebook2.png">
                                    <p class="fb-slogan">Se questo articolo ti piace, condividilo con i tuoi amici su FACEBOOK</p>
                                </li>
                                
                            </ul>
                            
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