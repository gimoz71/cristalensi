<!--#include file="inc_strConn.asp"-->
<%
cat=request("pkid")				  
if cat="" then cat=0

if cat>0 then
	Set cat_rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Categorie1 WHERE PKId="&cat
	cat_rs.open sql,conn, 1, 1
	if cat_rs.recordcount>0 then
		titolo_cat=cat_rs("titolo_en")
		descrizione_cat=cat_rs("descrizione_en")
		NomePagina=cat_rs("NomePagina_en")
		
		title=cat_rs("testo1_en")
		'description=cat_rs("testo2_en")
		kw=cat_rs("testo2_en")
	end if
	cat_rs.close
	
	Call Visualizzazione("Categorie1",Cat,NomePagina)
else
	'Call Visualizzazione("",0,"prodotti.asp")
	response.Redirect("/en/prodotti.asp")
end if
%>
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title><%=title%> </title>
		<meta name="description" content="<%=TogliTAG(descrizione_cat)%>">
		<meta name="keywords" content="<%=kw%>">
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="../js/media-queries-ie.js"></script>
        <![endif]-->
        <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
        <script src="../js/jquery.blueberry.js"></script>
        <script src="../js/jquery.tipTip.js"></script>
        <link href="../css/css.css" rel="stylesheet" type="text/css">
        <link href="../css/blueberry.css" rel="stylesheet" type="text/css">
        <link href="../css/tipTip.css" rel="stylesheet" type="text/css">
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
            <link href="../css/tipTip_ie7.css" media="all" rel="stylesheet" type="text/css" />
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
                            <h1>Chosen category: <%=titolo_cat%></h1>
                            <%if descrizione_cat<>"" then%>
                            <p>
                                <i><%=NoLettAcc(descrizione_cat)%></i>
                            </p>
                            <%end if%>
                            <%
							Set prod_rs = Server.CreateObject("ADODB.Recordset")
							sql = "SELECT * FROM Categorie2 WHERE FkCategoria1="&cat&" ORDER BY Posizione"
							prod_rs.open sql,conn, 1, 1
							if prod_rs.recordcount>0 then
							%>
                            <ul class="galleria clearfix">
                                <%
								Do while not prod_rs.EOF
								
								id=prod_rs("PkId")
								'url="../../en/prodotti.asp?cat="&id
								url="/en/prodotti.asp?cat="&id
								%>
                                <li>
                                    <a href="<%=url%>" title="<%=titolo_cat%><%=" - "&prod_rs("titolo_en")%>">
                                        <%
										'file_img="../"&prod_rs("logo")
										file_img="/public/"&prod_rs("logo")
										if file_img<>"" then
										%>
                                        <img src="<%=file_img%>" width="160" height="120" style="margin-bottom: 10px" alt="<%=titolo_cat%><%=" - "&prod_rs("titolo_en")%>" title="<%=titolo_cat%><%=" - "&prod_rs("titolo_en")%>" />
                                        <%else%>
                                        <img src="/public/logo_cristalensi_piccolo.jpg" width="160" height="120" style="margin-bottom: 10px" alt="image of category <%=titolo_cat%><%=" - "&prod_rs("titolo_en")%> not available" />
                                        <%end if%>
                                        <span class="button_link"><%=prod_rs("titolo_en")%></span>
                                    </a>
                                </li>
                                <%
								prod_rs.movenext
								loop	
								%>
                            </ul>
                            <%else%>
                                <p><br /><br /><br />No subcategories to this</p>
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