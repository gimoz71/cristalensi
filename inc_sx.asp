<<<<<<< HEAD
<div id="sidebar-alt" class="clearfix">
    <div>
        <%if idsession>0 then%>
            <h3>Area clienti</h3>
            <p><font color="#990000"><strong>Benvenuto<br />&nbsp;<%=nome_log%></strong></font></p>
            <p>&raquo;<a href="/iscrizione.asp">I dati della tua iscrizione</a></p>
            <p>&raquo;<a href="/ordini_elenco.asp">I tuoi ordini</a></p>
            <p>&raquo;<a href="/commenti_form.asp">Inserisci un commento</a></p>
            <p>&raquo;<a href="/admin/logout.asp">Esci dall'Area clienti</a></p>
        <%else%>
            <%if Instr(MM_LoginAction, "iscrizione")>0 then%>
            <%else%>
                <h3>Area clienti</h3>
                <p>Per consultare il catalogo non &egrave; necessario iscriversi</p>
                <p><a href="#" class="info tiptip" title="<h3>Informazioni generali</h3>L'iscrizione al sito internet Cristalensi <strong>&egrave; obbligatoria solo per acquistare</strong> ma non per consultare il catalogo dei prodotti.<br />Gli iscritti, oltre che poter acquistare i prodotti inserendo solamente <strong style='color: red'>Login</strong> e <strong style='color: red'>Password</strong> senza ripetere l'iscrizione ogni volta, potranno stampare gli ordini, aggiornare i propri dati, mettere i commenti al sito internet e saranno aggiornati sulle nostre offerte.<br />Per tutte le altre informazioni relative alle condizioni di vendita consulta la pagina specifica: <a href='/condizioni_di_vendita.asp' title='Condizioni di vendita' style='color: red'>condizioni di vendita</a>.">Maggiori informazioni</a></p>
                <form method="post" name="logon" action="<%=MM_LoginAction%><%if Request.QueryString="contr=1" then%><%else%><%if Request.QueryString<>"" then%>&<%else%>?<%end if%>contr=1<%end if%>">

                   <label for="username">Login</label>
                   <input type="text" name="login">
                   <label for="password">Password</label>
                   <input type="password" name="password">
                   <%if contr=2 then%>
               		<p style="background: #900; color: #fff; font-weight: bold; padding: 5px;">Attenzione!<br />I dati inseriti non sono corretti.</p>
               	   <%end if%>
                   <button type="submit" class="button_link" style="margin-top:10px;">Entra</button>
               
               </form>
               <p><a href="/recupero_pw.asp" class="password-recover" title="Recupera la password smarrita">Recupera la password</a></p>
               <a href="/iscrizione.asp" title="Registrati per acquistare i nostri prodotti per illuminazione" class="button_link_red">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;REGISTRATI!&nbsp;&nbsp;&nbsp;&nbsp;</a>

               <h3>Condizioni</h3>
               <p><a href="/condizioni_di_vendita.asp" class="info" title="Regolamento per acquistare i prodotti Cristalensi">Condizioni di vendita</a></p>
               <p class="note">
               <strong>Costi di spedizione</strong><br />
               <span>Spedizione in tutta Italia (Isole comprese)</span><br />
               <span style="float:left; width:160px;">Ordine maggiore di 250�: &nbsp;0�</span><br />
               <span style="float:left; width:130px;">Ordine minore di 250�:</span><span style="float:right; width:15px;">10�</span><br />
               <span style="float:left; width:128px;">Ritiro in sede:</span><span style="float:right; width:16px;">&nbsp;&nbsp;0�</span><br />
               </p>
               <p class="note">
               <strong>Sistemi di pagamento</strong><br />
               <span style="float:left; width:120px;">Bonifico bancario:</span><span style="float:right; width:10px;">0�</span><br />
               <span style="float:left; width:120px;">Poste Pay:</span><span style="float:right; width:10px;">0�</span><br />
               <span style="float:left; width:120px;">Contassegno:</span><span style="float:right; width:10px;">4�</span><br />
               <span style="float:left; width:160px;">Carta di credito - PayPal: 2%</span><br />
               <span style="float:left; width:110px;">Prepagata - PayPal:</span><span style="float:right; width:13px;">2%</span><br />
               </p>
               <img class="negozio paypal" src="/images/cartedicredito.jpg" style="padding-bottom:30px;" title="Sistemi di pagamento" />
            <%end if%>
        <%end if%>

        <h3>Contattaci!</h3>
        <p>Il nostro personale sar&agrave; a Tua disposizione per qualsiasi informazione</p>
        <p><a href="/contatti.asp" class="info" title="Negozio illuminazione Firenze">Contatti e riferimenti</a></p>
        <p class="note"><strong>Cristalensi Snc</strong><br />
            di Lensi Massimiliano & C.<br />
            C.F. e Iscr. Reg. Impr. di Firenze 05305820481<br />
            R.E.A. Firenze 536760<br />
            50056 Montelupo F.no (FI)<br />
            Via arti e mestieri, 1<br />
            Tel.: 0571.911163<br />
            Fax: 0571.073327<br />
            E-mail: <a href="mailto:info@cristalensi.it">info@cristalensi.it</a>
        </p>
        <p><a href="/privacy.asp" class="info" title="Privacy policy e note legali Cristalensi">Privacy e note legali</a></p>
        <img class="telefono" src="/images/telefono_cristalensi.png" align="absmiddle" style="padding:10px 0 20px 0;" alt="Numero per chiamare lo staff del negozio Cristalensi, orario negozio dal Lunedi al Sabato: 0571.911163" />
        <p>Per <strong>aperture e chiusure</strong> straordinarie controlla gli aggiornamenti su Facebook</p>
        <h3>Seguici anche su</h3>
        <a href="http://www.facebook.com/pages/Cristalensi-vendita-lampade-per-interni-ed-esterni/144109972402284" target="_blank" title="Pagina ufficiale Cristalensi"><img class="fb" src="/images/facebook.png" align="absmiddle" border="0" alt="Collegati alla nostra pagina su Facebook" /></a>
        <p>&nbsp;</p>
    </div>
</div>
=======
<div id="sidebar-alt" class="clearfix">
    <div>
        <%if idsession>0 then%>
            <h3>Area clienti</h3>
            <p><font color="#990000"><strong>Benvenuto<br />&nbsp;<%=nome_log%></strong></font></p>
            <p>&raquo;<a href="/iscrizione.asp">I dati della tua iscrizione</a></p>
            <p>&raquo;<a href="/ordini_elenco.asp">I tuoi ordini</a></p>
            <p>&raquo;<a href="/commenti_form.asp">Inserisci un commento</a></p>
            <p>&raquo;<a href="/admin/logout.asp">Esci dall'Area clienti</a></p>
        <%else%>
            <%if Instr(MM_LoginAction, "iscrizione")>0 then%>
            <%else%>
                <h3>Area clienti</h3>
                <p>Per consultare il catalogo non &egrave; necessario iscriversi</p>
                <p><a href="#" class="info tiptip" title="<h3>Informazioni generali</h3>L'iscrizione al sito internet Cristalensi <strong>&egrave; obbligatoria solo per acquistare</strong> ma non per consultare il catalogo dei prodotti.<br />Gli iscritti, oltre che poter acquistare i prodotti inserendo solamente <strong style='color: red'>Login</strong> e <strong style='color: red'>Password</strong> senza ripetere l'iscrizione ogni volta, potranno stampare gli ordini, aggiornare i propri dati, mettere i commenti al sito internet e saranno aggiornati sulle nostre offerte.<br />Per tutte le altre informazioni relative alle condizioni di vendita consulta la pagina specifica: <a href='/condizioni_di_vendita.asp' title='Condizioni di vendita' style='color: red'>condizioni di vendita</a>.">Maggiori informazioni</a></p>
                <form method="post" name="logon" action="<%=MM_LoginAction%><%if Request.QueryString="contr=1" then%><%else%><%if Request.QueryString<>"" then%>&<%else%>?<%end if%>contr=1<%end if%>">

                   <label for="username">Login</label>
                   <input type="text" name="login">
                   <label for="password">Password</label>
                   <input type="password" name="password">
                   <%if contr=2 then%>
               		<p style="background: #900; color: #fff; font-weight: bold; padding: 5px;">Attenzione!<br />I dati inseriti non sono corretti.</p>
               	   <%end if%>
                   <button type="submit" class="button_link" style="margin-top:10px;">Entra</button>
               
               </form>
               <p><a href="/recupero_pw.asp" class="password-recover" title="Recupera la password smarrita">Recupera la password</a></p>
               <a href="/iscrizione.asp" title="Registrati per acquistare i nostri prodotti per illuminazione" class="button_link_red">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;REGISTRATI!&nbsp;&nbsp;&nbsp;&nbsp;</a>

               <h3>Condizioni</h3>
               <p><a href="/condizioni_di_vendita.asp" class="info" title="Regolamento per acquistare i prodotti Cristalensi">Condizioni di vendita</a></p>
               <p class="note">
               <strong>Costi di spedizione</strong><br />
               <span>Spedizione in tutta Italia (Isole comprese)</span><br />
               <span style="float:left; width:160px;">Ordine maggiore di 250�: &nbsp;0�</span><br />
               <span style="float:left; width:130px;">Ordine minore di 250�:</span><span style="float:right; width:15px;">10�</span><br />
               <span style="float:left; width:128px;">Ritiro in sede:</span><span style="float:right; width:16px;">&nbsp;&nbsp;0�</span><br />
               </p>
               <p class="note">
               <strong>Sistemi di pagamento</strong><br />
               <span style="float:left; width:120px;">Bonifico bancario:</span><span style="float:right; width:10px;">0�</span><br />
               <span style="float:left; width:120px;">Poste Pay:</span><span style="float:right; width:10px;">0�</span><br />
               <span style="float:left; width:120px;">Contassegno:</span><span style="float:right; width:10px;">4�</span><br />
               <span style="float:left; width:160px;">Carta di credito - PayPal: 2%</span><br />
               <span style="float:left; width:110px;">Prepagata - PayPal:</span><span style="float:right; width:13px;">2%</span><br />
               </p>
               <img class="negozio paypal" src="/images/cartedicredito.jpg" style="padding-bottom:30px;" title="Sistemi di pagamento" />
            <%end if%>
        <%end if%>

        <h3>Contattaci!</h3>
        <p>Il nostro personale sar&agrave; a Tua disposizione per qualsiasi informazione</p>
        <p><a href="/contatti.asp" class="info" title="Negozio illuminazione Firenze">Contatti e riferimenti</a></p>
        <p class="note"><strong>Cristalensi Snc</strong><br />
            di Lensi Massimiliano & C.<br />
            C.F. e Iscr. Reg. Impr. di Firenze 05305820481<br />
            R.E.A. Firenze 536760<br />
            50056 Montelupo F.no (FI)<br />
            Via arti e mestieri, 1<br />
            Tel.: 0571.911163<br />
            Fax: 0571.073327<br />
            E-mail: <a href="mailto:info@cristalensi.it">info@cristalensi.it</a>
        </p>
        <p><a href="/privacy.asp" class="info" title="Privacy policy e note legali Cristalensi">Privacy e note legali</a></p>
        <img class="telefono" src="/images/telefono_cristalensi.png" align="absmiddle" style="padding:10px 0 20px 0;" alt="Numero per chiamare lo staff del negozio Cristalensi, orario negozio dal Lunedi al Sabato: 0571.911163" />
        <p>Per <strong>aperture e chiusure</strong> straordinarie controlla gli aggiornamenti su Facebook</p>
        <h3>Seguici anche su</h3>
        <a href="http://www.facebook.com/pages/Cristalensi-vendita-lampade-per-interni-ed-esterni/144109972402284" target="_blank" title="Pagina ufficiale Cristalensi"><img class="fb" src="/images/facebook.png" align="absmiddle" border="0" alt="Collegati alla nostra pagina su Facebook" /></a>
        <p>&nbsp;</p>
    </div>
</div>
>>>>>>> 6f6a6654e3b247554e877335c9b1a368add52746
