<!--#include file="inc_strConn.asp"-->
<%
Call Visualizzazione("",0,"carrello1.asp")
	
	mode=request("mode")
	if mode="" then mode=0
	
	'se la session è già aperta sfrutto il pkid dell'ordine, altrimenti ne apro una
	IdOrdine=session("ordine_shop")
	if IdOrdine="" then IdOrdine=0
	
	id=request("id")
	if id="" then id=0
	
		if IdOrdine=0 and id<>0 then
			Set os1 = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM Ordini"
			os1.Open sql, conn, 3, 3
	
			os1.addnew
			os1("FkCliente")=idsession
			os1("stato")=0
			os1("TotaleCarrello")=0
			os1("TotaleGenerale")=0
			os1("DataOrdine")=now()
			os1("DataAggiornamento")=now()
			os1("IpOrdine")=Request.ServerVariables("REMOTE_ADDR")
			os1.update
	
			os1.close
			
			'Prendo l'id dell'ordine
			Set os2 = Server.CreateObject("ADODB.Recordset")
			sql = "Select @@Identity As pkid"
			os2.Open sql, conn, 1, 1
			IdOrdine=os2("pkid")
	
			os2.close
		
			'Creo una sessione con l'id dell'ordine
			Session("ordine_shop")=IdOrdine		
		
		end if
		
		IdOrdine=cInt(IdOrdine)
		
	'modifica del carrello: eliminazione o modifica di un articolo nel carrello	
		if mode=2 then
			cs = conn.Execute("Delete * FROM RigheOrdine Where FkOrdine="&IdOrdine)
			mode=0
		end if
		
		if mode=1 then
		
			eliminare=request("eliminare")
		'parte per eliminare il prodotto dal carrello
			if eliminare<>"" then
				arrProd = Split(eliminare, ", ")
				For iLoop = LBound(arrProd) to UBound(arrProd)
					cs = conn.Execute("Delete * FROM RigheOrdine Where PkId="&arrProd(iLoop))
	   			next
		'fine parte per eliminazione
			else
		'parte per la modifica delle quantita di un articolo nel carrello
				
			'modifica delle quantità
				Set ts = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT * FROM RigheOrdine where FkOrdine="&idordine
				ts.Open sql, conn, 3, 3
				num=0
				Do while not ts.EOF
					'aggiornamento
					PrezzoProdotto=ts("PrezzoProdotto")
					Quantita=request("quantita"&num)
					ts("Quantita")=Quantita
					ts("TotaleRiga")=(Quantita*PrezzoProdotto)
					ts.update
					num=num+1
					ts.movenext
				loop
				ts.close
			end if
		'fine della parte di modifica
			
		else
	'inserimento di un prodotto per la prima volta scelto con il carrello già aperto
			'Prendo il prezzo del prodotto
			
			
			if id<>0 then
				quantita=request("quantita")
				if quantita="" then quantita=1
				
				colore=request("colore")
				if colore="*****" then colore=""
				
				lampadina=request("lampadina")
				if lampadina="*****" then lampadina=""
				
				'prendo le caretteristriche del prodotto
				
				Set prodotto_rs = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT * FROM Prodotti where PkId="&id&""
				prodotto_rs.Open sql, conn, 1, 1

				PrezzoProdotto=prodotto_rs("PrezzoProdotto")
				CodiceArticolo=prodotto_rs("CodiceArticolo")
				TitoloProdotto=prodotto_rs("Titolo")
				
				prodotto_rs.close
				
				
				Set riga_rs = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT * FROM RigheOrdine"
				riga_rs.Open sql, conn, 3, 3
	
				riga_rs.addnew
				riga_rs("FkOrdine")=IdOrdine
				riga_rs("FkCliente")=idsession
				riga_rs("FkProdotto")=id
				riga_rs("PrezzoProdotto")=PrezzoProdotto
				riga_rs("Quantita")=Quantita
				TotaleRiga=PrezzoProdotto*Quantita
				riga_rs("TotaleRiga")=TotaleRiga
				riga_rs("colore")=Colore
				riga_rs("lampadina")=Lampadina
				riga_rs("CodiceArticolo")=CodiceArticolo
				riga_rs("Titolo")=TitoloProdotto
				riga_rs("Data")=now()
				riga_rs.update

				riga_rs.close
			end if
		end if		
				
				'Calcolo la somma per l'ordine
				Set rs2 = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT sum(TotaleRiga) as TotaleCarrello FROM RigheOrdine where FkOrdine="&IdOrdine
				rs2.Open sql, conn, 3, 3
				if rs2.recordcount>0 then	
					TotaleCarrello=rs2("TotaleCarrello")
					if TotaleCarrello="" then TotaleCarrello=0
				else
					TotaleCarrello=0
				end if
				rs2.close
				
				
				'Aggiorno la tabella dell'ordine con la somma calcolata prima
				Set ss = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT * FROM Ordini where PkId="&IdOrdine
				ss.Open sql, conn, 3, 3
				if ss.recordcount>0 then
					ss("TotaleCarrello")=TotaleCarrello
					ss("TotaleGenerale")=TotaleCarrello
					'ss("DataOrdine")=now()
					ss("DataAggiornamento")=now()
					ss("Stato")=0
					ss("FkCliente")=idsession
					ss("IpOrdine")=Request.ServerVariables("REMOTE_ADDR")
					ss.update
				end if
				ss.close
%>
<!doctype html>
<html>
    <head>
        <meta charset="iso-8859-1">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>CRISTALENSI Cart - You order</title>
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
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	'sql = "SELECT RigheOrdine.PkId, RigheOrdine.FkOrdine, RigheOrdine.PrezzoProdotto as PrezzoProdotto, RigheOrdine.FkProdotto, RigheOrdine.Quantita, RigheOrdine.TotaleRiga, Prodotti.Titolo, Prodotti.CodiceArticolo FROM Prodotti INNER JOIN RigheOrdine ON Prodotti.PkId = RigheOrdine.FkProdotto WHERE (((RigheOrdine.FkOrdine)="&idOrdine&"))"
	sql = "SELECT PkId, FkOrdine, FkProdotto, PrezzoProdotto, Quantita, TotaleRiga, Titolo, CodiceArticolo, Colore, Lampadina FROM RigheOrdine WHERE FkOrdine="&idOrdine&""
	rs.Open sql, conn, 1, 1
	num_prodotti_carrello=rs.recordcount
	
	Set ss = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where pkid="&idOrdine
	ss.Open sql, conn, 1, 1
%>                
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                            <h3 style="font-size: 14px; display: inline; border: none;">Your order:</h3> <p style="display: inline;"><%=num_prodotti_carrello%>&nbsp;<%if num_prodotti_carrello=1 then%>product added<%else%>products added<%end if%></p>  <%if idOrdine<>0 then%><a href="carrello1.asp?mode=2" style="float:right; font-size: 12px;">Empty cart <span style="display: inline-block; padding: 0 2px; font-weight: bold;border: dotted 1px #c00; font-style: italic; color: #c00;">x</span></a><%end if%>
                            <div class="carrello clearfix">
                                <p class="area clearfix"><span class="colonna articolo">[article code] product name</span><span class="colonna quantita">quantity</span><span class="colonna prezzo_unitario">unit cost</span><span class="colonna prezzo_totale">total</span><span class="colonna elimina">remove</span></p>
                                <div class="data">
                                    <%if rs.recordcount>0 then%>
                                        <form method="post" action="carrello1.asp?mode=1" style="margin=0px;">
                                        <%conta=0%>
                                        <%
                                        Do while not rs.EOF
                                        %>					
    
                                        <p class="riga"><span class="colonna articolo">[<%=rs("codicearticolo")%>]&nbsp;<strong><%=rs("titolo")%></strong><%if Len(rs("colore"))>0 or Len(rs("lampadina"))>0 then%><br /><%if Len(rs("colore"))>0 then%>&nbsp;Col.:&nbsp;<%=rs("colore")%><%end if%><%if Len(rs("lampadina"))>0 then%>&nbsp;-&nbsp;Light:&nbsp;<%=rs("lampadina")%><%end if%><%end if%></span>
                                        <%
                                        quantita=rs("quantita")
                                        if quantita="" then quantita=1
                                        %>
                                        <span class="colonna quantita">n&deg; pieces <input name="quantita<%=conta%>" value="<%=quantita%>" type="text" style="width: 20px"></span><span class="colonna prezzo_unitario"><%=FormatNumber(rs("PrezzoProdotto"),2)%>&#8364;</span><span class="colonna prezzo_totale"><%=FormatNumber(rs("TotaleRiga"),2)%>&#8364;</span><span class="colonna elimina"><input name="eliminare" value="<%=rs("pkid")%>" type="checkbox"></span></p>
                                        <%
                                        conta=conta+1
                                        rs.movenext
                                        loop
                                        %>
                                        <p class="riga" style="text-align: right"><button name="aggiorna" type="submit" class="button_link">update products</button></p>
                                        </form>
									<%else%>
                                    	<p class="riga">Cart is empty</p>
                                    <%end if%>    
                                </div>
                                <%if ss.recordcount>0 then%>
                                  <h4 class="cart clearfix"><span class="total_price">Total purchase: <%if ss("TotaleGenerale")<>0 then%>
								  <%=FormatNumber(ss("TotaleGenerale"),2)%>
                                  <%else%>
                                  0,00
                                  <%end if%>
                                  &#8364;
                                  </span></h4>
									<%if rs.recordcount>0 then%>
                                    <form method="post" action="<%if italia_log="Si" then%>carrello2.asp<%end if%><%if italia_log="No" then%>carrello2extra.asp<%end if%>">
                                    <h3 style="font-size:12px;">Any notes</h3>
                                    <p>You can use this space to enter any notes or communications in relation to the products purchased.</p>
                                    <textarea name="NoteCliente" cols="105" rows="2" id="NoteCliente"><%=ss("NoteCliente")%></textarea>
                                    <p><button type="submit" name="continua" style="float: left" class="button_link">&laquo; Click here to continue to buy</button>&nbsp;&nbsp;<button type="submit" name="continua" style="float: right" class="button_link_red">Click here to continue the order &raquo;</button></p>
                                    </form>
                                    <%end if%>
                                    <br>
                                    <h3 style="font-size:12px;">AVAIBILITY OF PRODUCTS</h3>
                                    <p>Our catalog is made up of numerous products and producers, therefore some products may not be immediately available.  In the case that you urgently need the desired product, <strong>ask our staff directly if the product is available immediately or how long it will take for the merchandise to be in stock</strong>.<br>
						Delivery will take place within a minimum of 2 days and a maximum of 30 days.<br>						<a href="contatti.asp">Contact to inquire about product availability</a></p>
								<%end if%>
								<%
                                ss.close
                                rs.close
                                %>
                            </div>
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