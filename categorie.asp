<!--#include file="inc_strConn.asp"-->
<%
cat=request("pkid")				  
if cat="" then cat=0

if cat>0 then
	Set cat_rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Categorie1 WHERE PKId="&cat
	cat_rs.open sql,conn, 1, 1
	if cat_rs.recordcount>0 then
		titolo_cat=cat_rs("titolo")
		descrizione_cat=cat_rs("descrizione")
		NomePagina=cat_rs("NomePagina")
		
		title=cat_rs("testo1")
		description=cat_rs("testo2")
		kw=title + " " + description
	end if
	cat_rs.close
	
	Call Visualizzazione("Categorie1",Cat,NomePagina)
else
	'Call Visualizzazione("",0,"prodotti.asp")
	response.Redirect("../../prodotti.asp")
end if
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
                            <h3>Categoria scelta: <%=titolo_cat%></h3>
                            <%if descrizione_cat<>"" then%>
                            <p>
                                <i><%=NoLettAcc(descrizione_cat)%></i>.
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
								'url="../../prodotti.asp?cat="&id
								url="prodotti.asp?cat="&id
								%>
                                <li>
                                    <a href="<%=url%>" title="<%=titolo_cat%><%=" - "&prod_rs("titolo")%>">
                                        <%
										file_img=prod_rs("logo")
										if file_img<>"" then
										%>
                                        <img src="../<%=file_img%>" width="160" height="120" style="margin-bottom: 10px" alt="<%=titolo_cat%><%=" - "&prod_rs("titolo")%>" title="<%=titolo_cat%><%=" - "&prod_rs("titolo")%>" />
                                        <%else%>
                                        <img src="../logo_cristalensi_piccolo.jpg" width="160" height="120" style="margin-bottom: 10px" alt="Immagine della categoria <%=titolo_cat%><%=" - "&prod_rs("titolo")%> non disponibile" />
                                        <%end if%>
                                        <span class="button_link"><%=prod_rs("titolo")%></span>
                                    </a>
                                </li>
                                <%
								prod_rs.movenext
								loop	
								%>
                            </ul>
                            <%else%>
                                <p><br /><br /><br />Nessuna sottocategoria presente</p>
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
          <script src="js/init.js"></script>
    </body>
</html>
<!--#include file="inc_strClose.asp"-->