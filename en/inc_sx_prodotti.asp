			<div id="sidebar-alt" class="clearfix">
                    <div>
                        <%'if Instr(MM_LoginAction, "prodott")>0 or Instr(MM_LoginAction, "categorie")>0 or Instr(MM_LoginAction, "offerte")>0 or Instr(MM_LoginAction, "produttori")>0 or Instr(MM_LoginAction, "ricerca")>0 then%>
                        <%if Instr(MM_LoginAction, "carrello")>0 or Instr(MM_LoginAction, "calcolo")>0 or Instr(MM_LoginAction, "ordin")>0 or Instr(MM_LoginAction, "iscrizione")>0 or Instr(MM_LoginAction, "pagamento")>0 then%>
                        <%else%>
                            <h3>Product Gallery</h3>
                            <ul class="product-menu">
                            <%
                            Set cat_rs = Server.CreateObject("ADODB.Recordset")
                            sql = "SELECT * FROM Categorie1 ORDER BY Posizione ASC"
                            cat_rs.open sql,conn, 1, 1
                            if cat_rs.recordcount>0 then
                            Do while not cat_rs.EOF
							nomepagina_categorie = cat_rs("NomePagina_en")
							if nomepagina_categorie="" then nomepagina_categorie="#"
							if nomepagina_categorie<>"#" then nomepagina_categorie="/public/pagine/"&nomepagina_categorie
							'if nomepagina_categorie<>"#" then nomepagina_categorie="/public/pagine/categorie_en.asp?pkid="&cat_rs("PkId")
                            %>    
                                <li><a href="<%=nomepagina_categorie%>" title="<%=cat_rs("Titolo_en")%>"><span><%=cat_rs("Titolo_en")%></span></a></li>
                            <%
                            cat_rs.movenext
                            loop
                            end if
                            cat_rs.close
                            %>
                            </ul>
                            
                            <%
                            Set cs=Server.CreateObject("ADODB.Recordset")
                            sql = "Select * From Produttori order by titolo ASC"
                            cs.Open sql, conn, 1, 1
                            if cs.recordcount>0 then
                            %>
                            <h3>List of producers</h3>
                            <p>If you know the brand of the product you can select below or going to the complete list of producers</p>
                            <SCRIPT LANGUAGE=javascript>
                            <!--
                                function invia_produttore() {
                                    document.getElementById("form_produttori").submit();
                                }
                            // End -->
                            </SCRIPT>
                            <form method="post" name="form_produttori" id="form_produttori" action="/en/prodotti.asp">
                            
                                <select name="FkProduttore" id="FkProduttore" class="form" onChange="invia_produttore()">
                                    <option value="0">Select a brand</option>
                                    <%
                                    Do While Not cs.EOF
                                    %>
                                    <option value="<%=cs("pkid")%>"><%=cs("titolo")%></option>
                                    <%
                                    cs.movenext
                                    loop
                                    %>
                                </select>
                            </form>
                            <p style="font-weight: bold"><a href="/en/produttori.asp" title="Complete list of lights producers" class="button_link">Complete list of producers</a></p>
                            <%end if%>
                            <%cs.close%>
                            <p><br /><br />Discover the latest offers online</p>
                            <a href="/en/offerte.asp" class="button_link_red" title="Latest offers of lamps">&nbsp;&nbsp;LATEST OFFERS&nbsp;&nbsp;</a>
                            <p><br /><br />Do You want to search for a product code, name or in a price range? Do You want to combine a number of features?
                                Use the advanced search</p>
                            <a href="/en/ricerca_avanzata_modulo.asp" class="button_link_red" title="Advanced search lighting products">ADVANCED SEARCH</a>
						<%end if%>
                        
                        <%if idsession>0 then%>
                            <h3>Client Area</h3>
                            <p><font color="#990000"><strong>Welcome<br />&nbsp;<%=nome_log%></strong></font></p>
                            <p>&raquo;<a href="/en/iscrizione.asp">The details of your registration</a></p>
                            <p>&raquo;<a href="/en/ordini_elenco.asp">Your orders</a></p>
                            <p>&raquo;<a href="/en/commenti_form.asp">Post a comment</a></p>
                            <p>&raquo;<a href="/admin/logout.asp">Exit from your profile</a></p>
                            <p>&nbsp;</p>
                            
                        <%end if%>
                        
                        <h3>Contact us!</h3>
                        <p>Our staff will be pleased to answer any questions</p>
                        <p><a href="/en/contatti.asp" class="info" title="Lighthing Showromm in Florence">Contacts and references</a></p>
                        <p class="note"><strong>Cristalensi Snc</strong><br />
                            di Lensi Massimiliano & C.<br />
                            C.F. e Iscr. Reg. Impr. di Florence 05305820481<br />
                            R.E.A. Florence 536760<br />
                            Florence (Italy)<br />
                            50056 Montelupo F.no<br />
                            Via arti e mestieri, 1<br />
                            Tel. e Fax: 0571.911163<br />
                            E-mail: <a href="mailto:info@cristalensi.it">info@cristalensi.it</a>
                        </p>
                       
                        <h3>Follow us on</h3>
                        <a href="http://www.facebook.com/pages/Cristalensi-vendita-lampade-per-interni-ed-esterni/144109972402284" target="_blank" title="Official page of Cristalensi"><img class="fb" src="/images/facebook.png" align="absmiddle" border="0" alt="Follow us on Facebook" /></a>
                        <p>&nbsp;</p>
                        
                    </div>
                </div>
