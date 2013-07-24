<!--#include file="inc_strConn.asp"-->
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>>CRISTALENSI ricerca avanzata catalogo prodotti illuminazione</title>
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
        <SCRIPT language="JavaScript">

		function verifica() {
				
			titolo=document.newsform.titolo.value;
			cat=document.newsform.Cat.value;
			FkProduttore=document.newsform.FkProduttore.value;
			prezzo_da=document.newsform.prezzo_da.value;
			prezzo_a=document.newsform.prezzo_a.value;
		
			if (titolo=="" && cat=="0" && FkProduttore=="0" && prezzo_da=="" && prezzo_a==""){
				alert("Inserire o scegliere almeno un valore per effettuare la ricerca.");
				return false;
			}
		
			else
		return true
		
		}
		
		</SCRIPT>
        <!--Codice Statistiche Google Analytics Iury Mazzoni ## NON CANCELLARE!! ## -->
		<script type="text/javascript">
        
          //var _gaq = _gaq || [];
//          _gaq.push(['_setAccount', 'UA-320952-2']);
//          _gaq.push(['_trackPageview']);
//        
//          (function() {
//            var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
//            ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
//            var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
//          })();
        
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
                        	
                        	
                            <h3 style="font-size: 14px; display: inline; border: none;">Ricerca avanzata articoli illuminazione</h3>
                            <p>Grazie alla Ricerca avanzata hai la possiblit&agrave; di inserire il <strong>Nome</strong> o il <strong>Codice</strong> dell'articolo di illuminazione, oppure selezionare una <strong>Categoria</strong> o il <strong>Produttore</strong>, oppure inserire una <strong>fascia di prezzo</strong> ma puoi anche combinare le singole ricerche per arrivare ad un elenco maggiormente filtrato e <strong>su misura per i tuoi desideri</strong>.
                            </p>
                            <div class="iscrizione clearfix">                                
                                <div class="table">
                                    <form method="post" action="ricerca_avanzata_elenco.asp" name="newsform" onSubmit="return verifica();">
                                    <div class="tr">
	                                        Nome o Codice del prodotto<br />
                                            <input name="titolo" type="text" id="titolo" style="width:300px;" />
                                    </div>
                                    <div class="tr">
                                        <div class="td">

                                        Categoria<br />
                                            <%
                                                Set cs=Server.CreateObject("ADODB.Recordset")
                                                sql = "SELECT Categorie1.PkId as PkId_1, Categorie1.Titolo as Titolo_1, Categorie2.PkId as PkId_2, Categorie2.Titolo as Titolo_2 "
                                                sql = sql + "FROM Categorie1 INNER JOIN Categorie2 ON Categorie1.PkId = Categorie2.Fkcategoria1 "
                                                'sql = sql + "WHERE Categorie2.FkCategoria1 = "&cat_principale&" "
                                                sql = sql + "ORDER BY Categorie1.Titolo ASC, Categorie2.Titolo ASC"
                                                cs.Open sql, conn, 1, 1
                                            %>
                                            <select name="Cat" id="Cat" style="width:300px;">
                                                    <option value="0" >Seleziona una categoria</option>
                                                    <%
                                                    if cs.recordcount>0 then
                                                    Do While Not cs.EOF
                                                    %>
                                                    <option title="<%=cs("Titolo_2")%>" value=<%=cs("pkid_2")%> ><%=cs("Titolo_1")%> - <%=cs("Titolo_2")%></option>
                                                    <%
                                                    cs.movenext
                                                    loop
                                                    end if
                                                    %>
                                             </select>
                                             <%cs.close%>
                                    </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                            Produttore<br />
                                            <%
                                            Set cs=Server.CreateObject("ADODB.Recordset")
                                            sql = "Select * From Produttori order by titolo ASC"
                                            cs.Open sql, conn, 1, 1
                                            if cs.recordcount>0 then
                                            %>
                                            <select name="FkProduttore" id="FkProduttore" style="width:300px;">
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
                                            <%end if%>
                                            <%cs.close%>
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">

	                                        Fascia di prezzo<br />
                                            Da <input name="prezzo_da" type="text" id="prezzo_da" style="width:100px;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;A <input name="prezzo_a" type="text" id="prezzo_a" style="width:100px;" />
                                        </div>
                                    </div>
                                    
                                    <div class="tr">
                                        <div class="td">

                                            <button name="Submit" type="submit" class="button_link" value="Avvia la ricerca" align="absmiddle">Avvia la ricerca</button>
                                        </div>
                                    </div>
                                    </form>
                                </div>
                            </div>
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