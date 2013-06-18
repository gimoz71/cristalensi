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
                            <p>Se conosci una marca puoi selezionare direttamente il produttore per avere il relativo elenco di prodotti</p>
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
                            <p style="font-weight: bold"><a href="produttori.asp" title="Elenco completo dei produttori di articoli per illuminazione" class="button_link">Elenco completo marche</a></p>
                            <%end if%>
                            <%cs.close%>
						<%end if%>
                        <p><br /><br />Scopri le ultime offerte del nostro catalogo online</p>
                        <a href="offerte.asp" class="button_link_red">PRODOTTI IN OFFERTA</a>
                        <p><br /><br />Vuoi ricercare un prodotto per codice, per nome oppure in una fascia di prezzo? Vuoi combinare una serie di caratteristiche?
                            Sfrutta la ricerca avanzata</p>
                        <a href="ricerca_avanzata_modulo.asp" class="button_link_red">RICERCA AVANZATA</a>
                        <h3>Contattaci!</h3>
                        <p>Il nostro personale sar&agrave; a Tua disposizione per qualsiasi informazione</p>
                        <p><a href="#" class="info">Contatti e riferimenti</a></p>
                        <p class="note"><strong>Cristalensi Snc</strong><br />
                            di Lensi Massimiliano & C.<br />
                            C.F. e Iscr. Reg. Impr. di Firenze 05305820481<br />
                            R.E.A. Firenze 536760<br />
                            50056 Montelupo F.no (FI)<br />
                            Via arti e mestieri, 1<br />
                            Tel. e Fax: 0571.911163<br />
                            E-mail: <a href="mailto:info@cristalensi.it">info@cristalensi.it</a>
                        </p>
                        <img src="images/telefono_cristalensi.gif" align="absmiddle" style="padding:10px 0 30px 0;" alt="Numero per chiamare lo staff del negozio Cristalensi, orario negozio dal Lunedi al Sabato: 0571.911163" />
                        <h3>Seguici anche su</h3>
                        <a href="http://www.facebook.com/pages/Cristalensi-vendita-lampade-per-interni-ed-esterni/144109972402284" target="_blank" title="Pagina ufficiale Cristalensi"><img src="images/facebook.png" align="absmiddle" border="0" alt="Collegati alla nostra pagina su Facebook" /></a>
                    </div>
                </div>
