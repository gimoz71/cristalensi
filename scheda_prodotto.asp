<!--#include file="inc_strConn.asp"-->
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
    	<SCRIPT language="JavaScript">
		function verifica_1() {
				
			quantita=document.newsform2.quantita.value;
			num_colori=document.newsform2.num_colori.value;
			colore=document.newsform2.colore.value;
		
			if (quantita=="0"){
				alert("La quantita\' deve essere maggiore di 0");
				return false;
			}
			
			if (num_colori>0 && colore==""){
				alert("Deve essere scelto un colore");
				return false;
			}
			
			else
				
				document.newsform2.method = "post";
				document.newsform2.action = "../../carrello1.asp";
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
			
			if (num_colori>0 && colore==""){
				alert("Deve essere scelto un colore");
				return false;
			}
			
			else
				
				document.newsform2.method = "post";
				document.newsform2.action = "../../carrello1.asp";
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
                            <p>Le nostre offerte in vetrina a prezzi scontati. Consulta tutte le offerte nell'apposita sezione "Prodotti in offerta"<br>Non perdere l'occasione!!!</p>
                            <ul class="scheda-prodotto clearfix">
                                <li class="clearfix">
                                    <p class="area clearfix">Codice articolo <strong>[AP ANTIGUA 1R]</strong><span class="produttore">produttore: <a href="#"><strong>Illumnando</strong></a></span></p>
                                    <div class="data">
                                        <h3>Applique di vetro murano</h3> 
                                        <p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vel mauris vitae urna ornare convallis a eu elit. Nulla varius lobortis lorem a molestie. Maecenas dictum pretium tellus, quis porttitor ipsum congue sed. Nunc vel sodales arcu. Praesent at ipsum at nisi aliquam commodo.</p>
                                        <br
                                        <p> Il prodotto lo trovi nella categoria:<br>
                                            <a href="#">Faretti e binari a parete</a>
                                            
                                        <p class="cart clearfix"><span class="price">Prezzo listino: <span>155€</span></span> <span class="cristalprice">Prezzo listino: 155€</span><a href="#" class="cart-link">Inserisci nel carrello</a></p>
                                    </div>
                                    <div class="thumb">
                                        <a href="#">
                                            <img src="images/example.jpg">
                                        </a>
                                    </div>
                                    <div class="thumb">
                                        <a href="#">
                                            <img src="images/example.jpg">
                                        </a>
                                    </div>
                                    <div class="thumb">
                                        <a href="#">
                                            <img src="images/example.jpg">
                                        </a>
                                    </div>
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