			<div id="sidebar-alt" >
                    <div>
                        <%if Instr(MM_LoginAction, "prodott")>0 or Instr(MM_LoginAction, "categorie")>0 or Instr(MM_LoginAction, "offerte")>0 or Instr(MM_LoginAction, "produttori")>0 or Instr(MM_LoginAction, "ricerca")>0 then%>
                            <h3>Galleria prodotti</h3>
                            <ul class="product-menu">
                            <%
                            Set cat_rs = Server.CreateObject("ADODB.Recordset")
                            sql = "SELECT * FROM Categorie1 ORDER BY Posizione ASC"
                            cat_rs.open sql,conn, 1, 1
                            if cat_rs.recordcount>0 then
                            Do while not cat_rs.EOF
                            %>    
                                <li><a href="/public/pagine/<%=cat_rs("NomePagina")%>" title="<%=cat_rs("Titolo")%>">&raquo; <%=cat_rs("Titolo")%></a></li>
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
                            <h3>Elenco produttori</h3>
                            <p>Puoi consultare il catalogo dei prodotti scegliendo un produttore dalla lista seguente: selezionandone uno otterrai una lista di articoli di quell'impresa</p>
                            <SCRIPT LANGUAGE=javascript>
                            <!--
                                function invia_produttore() {
                                    document.getElementById("form_produttori").submit();
                                }
                            // End -->
                            </SCRIPT>
                            <form method="post" name="form_produttori" id="form_produttori" action="prodotti.asp">
                            
                                <select name="FkProduttore" id="FkProduttore" class="form" onChange="invia_produttore()">
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
                            </form>
                            <p style="font-weight: bold"><a href="produttori.asp" title="Elenco completo dei produttori di articoli per illuminazione">Elenco completo</a></p>
                            <%end if%>
                            <%cs.close%>
						<%end if%>
                        <p>Scopri le ultime offerte e la galleria dei prodotti</p>
                        <p style="font-weight: bold"><a href="#">[ULTIME OFFERTE]</a></p>
                        <p style="font-weight: bold"><a href="#">[GALLERIA PRODOTTI]</a></p>
                        <p>Vuoi ricercare un prodotto per codice, per nome oppure in una fascia di prezzo? Vuoi combinare una serie di caratteristiche?
                            Hai a disposizione la nuova</p>
                        <button>RICERCA AVANZATA</button>
                        <h3>Contattaci!</h3>
                        <p>Il nostro personale sarà a tua disposizione per qualsiasi chiarimento e informazione su tutto l'assortimento di articoli
                            per l'illuminazione per la casa e per il giardino.</p>
                        <p>Il nostro personale sar&agrave; a Tua disposizione per qualsiasi informazione:</p>
                        <p><strong>Cristalensi Snc</strong><br>
                            Di Lensi Massimiliano & C.<br>
                            C.F. e Iscr. Reg. Impr. di Firenze 05305820481<br>
                            R.E.A. Firenze 536760<br>

                            50056 Montelupo F.no (FI)<br>
                            Via arti e mestieri, 1<br>
                            Tel. e Fax: 0571 911163<br>
                            e-mail: <a href="mailto:info@cristalensi.it">info@cristalensi.it</a>
                        </p>
                    </div>
                </div>
