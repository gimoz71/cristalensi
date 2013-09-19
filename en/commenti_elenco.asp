<!--#include file="inc_strConn.asp"-->
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Comments about lighthings shop online reviews of italian lighting products</title>
        <meta name="description" content="Comments about lighthings shop online and reviews of italian lighting products.">
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
                            <h3 style="font-size: 14px; display: inline; border: none;">Commenti</h3>
                            
                            <div class="carrello clearfix">
                                <div class="data">
                                    <p><em>With a view to transparency, approach to customers and improving our services we have opened this area where customers can submit a message or comment on the functioning of the website, or a review lighting products purchased but also to the services of the staff itself. <br /> To post a comment you must be registered on the website and sent messages will be approved by the staff to prevent the publication or offensive lyrics insert advertising links to other Internet sites.</em></p>
                                    <p class="riga" style="text-align: right; padding-bottom:10px;"><a href="commenti_form.asp" class="button_link_red">Send a comment you too!</a></p>
									<%
									Set prod_rs = Server.CreateObject("ADODB.Recordset")
									'sql = "SELECT * FROM Commenti_Clienti WHERE Pubblicato=True ORDER BY PkId DESC"
									sql = "SELECT Commenti_Clienti.PkId, Commenti_Clienti.Testo, Commenti_Clienti.Risposta, Commenti_Clienti.Pubblicato, Clienti.Nome, Clienti.Citta "
									sql = sql + "FROM Commenti_Clienti INNER JOIN Clienti ON Commenti_Clienti.FkIscritto = Clienti.PkId "
									sql = sql + "WHERE (((Commenti_Clienti.Pubblicato)=True)) "
									sql = sql + "ORDER BY Commenti_Clienti.PkId DESC"

									prod_rs.open sql,conn, 1, 1
									if prod_rs.recordcount>0 then
										Do while not prod_rs.EOF
											pkid_commento=prod_rs("PkId")
											testo_commento=prod_rs("testo")
											risposta=prod_rs("risposta")
											nome=prod_rs("nome")
											citta=prod_rs("citta")
											if risposta="" then risposta=False
											if risposta=True then
												Set ris_rs = Server.CreateObject("ADODB.Recordset")
												sql = "SELECT * FROM Commenti_Risposte WHERE FkCommento="&pkid_commento&" AND Pubblicato=True"
												ris_rs.open sql,conn, 1, 1
												if ris_rs.recordcount>0 then
													testo_risposta=ris_rs("Testo")
												end if
												ris_rs.close
											end if
									%>
                                        <div class="riga">
										<p><%=NoLettAcc(testo_commento)%><br /><strong><%=Nome%>&nbsp;(<%=Citta%>)</strong></p>
                                        <%if testo_risposta<>"" and risposta=True then%>
                                        <p style="padding:0px 10px;"><strong>Replay staff Cristalensi:</strong><br /><em><%=NoLettAcc(testo_risposta)%></em></p>
                                        <%end if%>
                                        </div>
                                        
                                    <% 
										prod_rs.movenext
										loop
									end if
									prod_rs.close
									%>
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