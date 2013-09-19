<!--#include file="inc_strConn.asp"-->
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>CRISTALENSI Advanced search for Lamps Lights</title>
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
        <SCRIPT language="JavaScript">

		function verifica() {
				
			titolo=document.newsform.titolo.value;
			cat=document.newsform.Cat.value;
			FkProduttore=document.newsform.FkProduttore.value;
			prezzo_da=document.newsform.prezzo_da.value;
			prezzo_a=document.newsform.prezzo_a.value;
		
			if (titolo=="" && cat=="0" && FkProduttore=="0" && prezzo_da=="" && prezzo_a==""){
				alert("Enter or select at least one value to search.");
				return false;
			}
		
			else
		return true
		
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
        <div id="wrap">
            <!--#include file="inc_header.asp"-->
            <div id="main-content">
                
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                        	
                        	
                            <h3 style="font-size: 14px; display: inline; border: none;">Advanced search for Lamps & Lights</h3>
                            <p>With advanced search you have the possibility to enter the lighthing product<strong> name</strong> or <strong>code</strong>, or select a <strong>category</strong> or <strong>manufacturer</strong>, or enter a <strong>price range</strong> but you can also combine the individual searches to get a filtered list of Lamps & Lights, and more tailored to your desires.
                            </p>
                            <div class="iscrizione clearfix">                                
                                <div class="table">
                                    <form method="post" action="ricerca_avanzata_elenco.asp" name="newsform" onSubmit="return verifica();">
                                    <p>&nbsp;</p>
                                    <div class="tr text-center">
	                                        Name or Code of product<br />
                                            <input name="titolo" type="text" id="titolo" style="width:300px;" />
                                    </div>
                                    <div class="tr text-center">
                                        <div class="td">

                                        Category<br />
                                            <%
                                                Set cs=Server.CreateObject("ADODB.Recordset")
                                                sql = "SELECT Categorie1.PkId as PkId_1, Categorie1.Titolo_en as Titolo_1, Categorie2.PkId as PkId_2, Categorie2.Titolo_en as Titolo_2 "
                                                sql = sql + "FROM Categorie1 INNER JOIN Categorie2 ON Categorie1.PkId = Categorie2.Fkcategoria1 "
                                                'sql = sql + "WHERE Categorie2.FkCategoria1 = "&cat_principale&" "
                                                sql = sql + "ORDER BY Categorie1.Titolo_en ASC, Categorie2.Titolo_en ASC"
                                                cs.Open sql, conn, 1, 1
                                            %>
                                            <select name="Cat" id="Cat" style="width:300px;">
                                                    <option value="0" >Select the category</option>
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
                                    <div class="tr text-center">
                                        <div class="td">
                                            Producers<br />
                                            <%
                                            Set cs=Server.CreateObject("ADODB.Recordset")
                                            sql = "Select * From Produttori order by titolo ASC"
                                            cs.Open sql, conn, 1, 1
                                            if cs.recordcount>0 then
                                            %>
                                            <select name="FkProduttore" id="FkProduttore" style="width:300px;">
                                            <option value="0">Select the producer</option>
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
                                    <div class="tr text-center">
                                        <div class="td">

	                                        Price range<br />
                                            From <input name="prezzo_da" type="text" id="prezzo_da" style="width:100px;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;To <input name="prezzo_a" type="text" id="prezzo_a" style="width:100px;" />
                                        </div>
                                    </div>
                                    
                                    <div class="tr text-center">
                                        <div class="td">

                                            <button name="Submit" type="submit" class="button_link" value="Start the search" align="absmiddle">Start the search</button>
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