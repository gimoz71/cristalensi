<!--#include file="../../en/inc_strConn.asp"-->
<!--#include file="../../en/inc_clsImageSize.asp"-->
<%
'id=request("id")
if id="" then id=0
if id=0 then response.Redirect("prodotti.asp")

if id>0 then
	Set prod_rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Prodotti WHERE PKId="&id
	prod_rs.open sql,conn, 3, 3
	if prod_rs.recordcount>0 then
		CodiceArticolo=prod_rs("CodiceArticolo")
		'FkCat_Prod=prod_rs("FkCat_Prod")
		Titolo_prodotto=prod_rs("Titolo_en")
		Descrizione_prodotto=prod_rs("Descrizione_en")
		allegato_prodotto=prod_rs("Allegato")
		PrezzoArticolo=prod_rs("PrezzoProdotto")
		PrezzoListino=prod_rs("PrezzoListino")
		fkproduttore=prod_rs("fkproduttore")
		if fkproduttore="" then fkproduttore=0
		
		offerta=prod_rs("offerta")
		if offerta="" then offerta=0
		
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
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title><%=Titolo_prodotto%> - <%=titolo_cat%> CRISTALENSI Product description</title>
		<meta name="description" content="Product description <%=Titolo_prodotto%> Cristalensi shop online <%=TogliTAG(descrizione_cat)%>">
		<meta name="keywords" content="Product description <%=Titolo_prodotto%> Cristalensi shop online <%=kw%>">
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="/js/media-queries-ie.js"></script>
        <![endif]-->
        <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
        <script src="/js/jquery.blueberry.js"></script>
        <script src="/js/jquery.tipTip.js"></script>
        <script src="/js/jquery.fancybox.js"></script>
        <link href="/css/css.css" rel="stylesheet" type="text/css">
        <link href="/css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="/css/jquery.fancybox.css" rel="stylesheet" type="text/css">
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
    	<SCRIPT language="JavaScript">
            $(document).ready(function() {
                $('.fancybox').fancybox();
            });
		function verifica_1() {
				
			quantita=document.newsform2.quantita.value;
			num_colori=document.newsform2.num_colori.value;
			colore=document.newsform2.colore.value;
			num_lampadine=document.newsform2.num_lampadine.value;
			lampadina=document.newsform2.lampadina.value;
		
			if (quantita=="0"){
				alert("The quantity must be greater than 0");
				return false;
			}
			
			if (num_colori>1 && colore==""){
				alert("You have to choose a color");
				return false;
			}
			
			if (num_lampadine>1 && lampadina==""){
				alert("You have to choose a light");
				return false;
			}
			
			else
				
				document.newsform2.method = "post";
				//document.newsform2.action = "../../carrello1.asp";
				document.newsform2.action = "/en/carrello1.asp";
				document.newsform2.submit();
		}
		</SCRIPT>
		<SCRIPT language="JavaScript">
		function verifica_2() {
				
			quantita=document.newsform2.quantita.value;
			num_colori=document.newsform2.num_colori.value;
			colore=document.newsform2.colore.value;
			num_lampadine=document.newsform2.num_lampadine.value;
			lampadina=document.newsform2.lampadina.value;
		
			if (quantita=="0"){
				alert("The quantity must be greater than 0");
				return false;
			}
			
			if (num_colori>1 && colore==""){
				alert("You have to choose a color");
				return false;
			}
			
			if (num_lampadine>1 && lampadina==""){
				alert("You have to choose a light");
				return false;
			}
			
			else
				
				document.newsform2.method = "post";
				//document.newsform2.action = "../../carrello1.asp";
				document.newsform2.action = "/en/carrello1.asp";
				//document.newsform2.submit();
		}
		</SCRIPT>
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
            <!--#include file="../../en/inc_header.asp"-->

            <div id="main-content">
                
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <div class="spacer">
                            </div>

                            <ul class="scheda-prodotto clearfix">
                                <li class="clearfix">
                                    <a href="javascript:history.back()" class="cart-link button_link">Go back</a>
                                    <h1><%=Titolo_prodotto%> - <%=codicearticolo%></h1>
                                    <p class="area clearfix"><%if codicearticolo<>"" then%>Article code <strong>[<%=codicearticolo%>]</strong><%end if%><%if fkproduttore>0 then%><span class="produttore">producers: <a href="/en/prodotti.asp?FkProduttore=<%=fkproduttore%>" title="List of product of the same producer: <%=produttore%>"><strong><%=produttore%></strong></a></span><%end if%></p>
                                    <div class="data">
                                        <%if prezzoarticolo=0 then%>
                                           <p class="cart-panel clearfix" style="float: right; width: 30%;  text-align: center;"><br /><span class="price">List price: <span><%=prezzolistino%>€</span></span><br /><br />
                                       <%else%>
                                           <p class="cart-panel clearfix" style="float: right; width: 30%;  text-align: center;"><%if prezzolistino<>0 then%><span class="price">List price: <span><%=prezzolistino%>&#8364;</span></span><%end if%><br><%if prezzoarticolo<>"" then%><span class="cristalprice">Cristalensi price: <%=prezzoarticolo%>&#8364;</span><%end if%><br><i>VAT included</i>
                                       <%end if%>
                                        <p><%=descrizione_prodotto%></p>
                                        <%if FkCategoria2>0 then%>
                                            <p> You find the product in the category: <a href="/en/prodotti.asp?cat=<%=FkCategoria2%>" title="List of products from the same category: <%=titolo_cat%>"><%=titolo_cat%></a></p>
					<%end if%>
                                        <%if allegato_prodotto<>"" then%>
                                            <p><a href="/public/<%=allegato_prodotto%>" target="_blank"><img src="/images/file.jpg" border="0" width="18" height="18" hspace="3" align="absmiddle" alt="Attached file">Attached file</a></p>
                                        <%end if%>
                                        <%if prezzoarticolo=0 then%>
                                            <p class="cart clearfix"><a href="#" onClick="MM_openBrWindow('/en/richiesta_informazioni.asp?codice=<%=codicearticolo%>&titolo=<%=titolo_prodotto%>&amp;produttore=<%=produttore%>&amp;id=<%=id%>','','width=650,height=650,scrollbars=yes')" class="cart-link button_link_red">Do you want to know the Cristalensi price? Click here to have an estimate from our staff</a>
                                        <%else%>
                                        	<%if offerta=10 then%>
											<p class="cart clearfix"><span class="cristalprice" style="float:right;">THE PRODUCT IS NOT AVAILABLE, CONTACT US!&nbsp;&nbsp;</span></p>
											<%else%>

                                            <form name="newsform2" id="newsform2" onSubmit="return verifica_2();">
                                                <input type="hidden" name="id" id="id" value="<%=id%>">
                                                <%
                                                Set col_rs = Server.CreateObject("ADODB.Recordset")
                                                sql = "SELECT [Prodotto-Colore].FkProdotto, Colori.Titolo_en FROM [Prodotto-Colore] INNER JOIN Colori ON [Prodotto-Colore].FkColore = Colori.PkId WHERE ((([Prodotto-Colore].FkProdotto)="&id&")) ORDER BY Colori.Titolo_en ASC"
                                                col_rs.open sql,conn, 1, 1
                                                if col_rs.recordcount>1 then
                                                %>
                                                    <input type="hidden" name="num_colori" id="num_colori" value="<%=col_rs.recordcount%>">
                                                <%else%>
                                                    <input type="hidden" name="num_colori" id="num_colori" value="1">
                                                    <input type="hidden" name="colore" id="colore" value="*****">
                                                <%end if%>
                                                
                                                <%
                                                Set lam_rs = Server.CreateObject("ADODB.Recordset")
                                                sql = "SELECT [Prodotto-Lampadina].FkProdotto, Lampadine.Titolo_en FROM [Prodotto-Lampadina] INNER JOIN Lampadine ON [Prodotto-Lampadina].FkLampadina = Lampadine.PkId WHERE ((([Prodotto-Lampadina].FkProdotto)="&id&")) ORDER BY Lampadine.Titolo_en ASC"
                                                lam_rs.open sql,conn, 1, 1
                                                if lam_rs.recordcount>1 then
                                                %>
                                                    <input type="hidden" name="num_lampadine" id="num_lampadine" value="<%=lam_rs.recordcount%>">
                                                <%else%>
                                                    <input type="hidden" name="num_lampadine" id="num_lampadine" value="1">
                                                    <input type="hidden" name="lampadina" id="lampadina" value="*****">
                                                <%end if%>
                                                
                                                <p class="cart clearfix">

                                                    <%if col_rs.recordcount>1 then%>
                                                        <select name="colore" id="colore" style="width:auto; float:left; margin-top:7px; margin-right:10px;">
                                                        <option value="">Choose the color and/or the finish </option>
                                                        <%
                                                        Do While Not col_rs.EOF
                                                        %>
                                                            <option value="<%=col_rs("Titolo_en")%>"><%=col_rs("Titolo_en")%></option>
                                                        <%
                                                        col_rs.movenext
                                                        loop
                                                        %>
                                                        </select>
                                                    <%
                                                    end if
                                                    col_rs.close
                                                    %>
                                                    
                                                    <%if lam_rs.recordcount>1 then%>
                                                        &nbsp;&nbsp;
                                                        <select name="lampadina" id="lampadina" style="width:auto; float:left; margin-top:7px;">
                                                        <option value="">Choose the light and/or the glass</option>
                                                        <%
                                                        Do While Not lam_rs.EOF
                                                        %>
                                                            <option value="<%=lam_rs("Titolo_en")%>"><%=lam_rs("Titolo_en")%></option>
                                                        <%
                                                        lam_rs.movenext
                                                        loop
                                                        %>
                                                        </select>
                                                    <%
                                                    end if
                                                    lam_rs.close
                                                    %>
                                                    
                                                    <a href="#" onClick="return verifica_1();" id="invia_qta_2" rel="nofollow" title="Place in the shopping basket&nbsp;<%=titolo_prodotto%>&nbsp;<%=codicearticolo%>" class="cart-link button_link_red"><span>Add to cart</span></a><span style="float:right; padding-top:7px;"><input type="text" name="quantita" id="quantita" value="0" size="2" style="width:20px; text-align:right; margin-left:5px;">&nbsp;pieces&nbsp;&nbsp;</span>
                                                </p>
                                            </form>
                                            <%end if%>	
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
													percorso_img="/public/"&zoom
                                                    'percorso_img=path_img&zoom
                                                    'percorso_img="../"&zoom
                                            else
                                                    'percorso_img=server.mappath("public/"&file_img&"")
													percorso_img="/public/"&file_img
                                                    'percorso_img=path_img&file_img
                                                    'percorso_img="../"&file_img
                                            end if
                                            'calcolo misure immagini
                                            Set objImageSize = New ImageSize
                                            With objImageSize
                                              .ImageFile = server.mappath("/public/"&file_img&"")
                                              '.ImageFile = path_img&file_img
                                              '.ImageFile = "public/"&file_img

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


                                    <div class="thumb" style="width:<%if W>H then%><%if W<=160 then%><%=W%><%else%>160<%end if%><%else%><%if W<=90 then%><%=W%><%else%>90<%end if%><%end if%>px; height:<%if H<=120 then%><%=H%><%else%>120<%end if%>px">
                                        <a href="<%=percorso_img%>" class="fancybox" rel="gallery" title="<%if titolo_img<>"" then%><%=titolo_img%><%else%><%=titolo_prodotto%><%end if%>">
					<img class="img-border" src="/public/<%=file_img%>"  alt="<%if titolo_img<>"" then%><%=titolo_img%>&nbsp;<%=titolo_cat%><%else%><%=titolo_prodotto%>&nbsp;<%=titolo_cat%><%end if%>" title="<%if titolo_img<>"" then%><%=titolo_img%>&nbsp;<%=titolo_cat%><%else%><%=titolo_prodotto%>&nbsp;<%=titolo_cat%><%end if%>" /></a>
                                    </div>
                                    <%
                                    img_rs.movenext
                                    loop
                                    end if
                                    img_rs.close
                                    %>
                                </li>
                                <hr />
                                <li class="clearfix">
                                    <a href="http://www.facebook.com/pages/Cristalensi-vendita-lampade-per-interni-ed-esterni/144109972402284" target="_blank" title="Pagina ufficiale Cristalensi"><img src="/images/facebook2.png" hspace="10" align="absmiddle" border="0" alt="Pagina Ufficiale Cristalensi" class="facebook"></a><span style="line-height:80px;">If you like this article, share it with your friends on FACEBOOK</span>&nbsp;&nbsp;<div class="fb-like" data-send="false" data-layout="button_count" data-width="300" data-show-faces="false" data-font="verdana"></div>
                                    
                                </li>
                                <hr />
                                <li class="clearfix">
                                    <strong>Contact us!</strong> Our staff is at your disposal for any clarification, information and advice on the article desired.<br /><br />
                                </li>
                                
                            </ul>
                            
                        </div>
                    </div>
                </div>
                <!--#include file="../../en/inc_sx_prodotti.asp"-->
            </div>
        </div>
         <!--#include file="../../en/inc_footer.asp"-->
    </body>
</html>
<!--#include file="../../en/inc_strClose.asp"-->