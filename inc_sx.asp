<div id="sidebar-alt" class="clearfix">
    <div>
        <%if idsession>0 then%>
            <h3>Area clienti</h3>
            <p><font color="#990000"><strong>Benvenuto<br />&nbsp;<%=nome_log%></strong></font></p>
            <p>&raquo;<a href="iscrizione.asp">I dati della tua iscrizione</a></p>
            <p>&raquo;<a href="ordini_elenco.asp">I tuoi ordini</a></p>
            <p>&raquo;<a href="commenti_form.asp">Inserisci un commento</a></p>
            <p>&raquo;<a href="/admin/logout.asp">Esci dall'Area clienti</a></p>
        <%else%>
            <%if Instr(MM_LoginAction, "iscrizione")>0 then%>
            <%else%>
                <h3>Area clienti</h3>
                <p>Per consultare il catalogo non &egrave; necessario iscriversi</p>
                <p><a href="#" class="info tiptip" title="<h3>Informazioni generali</h3>L'iscrizione al sito internet Cristalensi <strong>&egrave; obbligatoria solo per acquistare</strong> ma non per consultare il catalogo dei prodotti.<br />Gli iscritti, oltre che poter acquistare i prodotti inserendo solamente <strong style='color: red'>Login</strong> e <strong style='color: red'>Password</strong> senza ripetere l'iscrizione ogni volta, potranno stampare gli ordini, aggiornare i propri dati, mettere i commenti al sito internet e saranno aggiornati sulle nostre offerte.<br />Per tutte le altre informazioni relative alle condizioni di vendita consulta la pagina specifica: <a href='condizioni_di_vendita.asp' style='color: red'>condizioni di vendita</a>.">Maggiori informazioni</a></p>
                <form method="post" name="logon" action="<%=MM_LoginAction%><%if Request.QueryString<>"" then%>&<%else%>?<%end if%>contr=1">

                   <label for="username">Login</label>
                   <input type="text" name="login">
                   <label for="password">Password</label>
                   <input type="password" name="password">
                   <button type="submit" style="margin-top:10px;">Entra</button>
               <%if contr=2 then%>
               <p style="background: #900; color: #fff; font-weight: bold; padding: 5px;">Attenzione! I dati inseriti<br />non sono corretti.</p>
               <%end if%>
               </form>
               <p><a href="recupero_pw.asp" class="password-recover">Recupera la password</a></p>
               <div align="center"><a href="iscrizione.asp" class="button_link_red">&nbsp;&nbsp;&nbsp;&nbsp;REGISTRATI!&nbsp;&nbsp;&nbsp;&nbsp;</a></div>

               <h3>Pagamenti</h3>
               <p><a href="condizioni_di_vendita.asp" class="info">Condizioni di vendita</a></p>
               <p class="note">Gli ordini potranno esser pagati in Contrassegno o con Bonifico Bancario oppure online grazie al sistema sicuro di PayPal con Carte di Credito e Prepagate. </p>
               <img class="negozio" src="images/cartedicredito.jpg" style="padding-bottom:30px;" />
            <%end if%>
        <%end if%>

        <h3>Contattaci!</h3>
        <p>Il nostro personale sar&agrave; a Tua disposizione per qualsiasi informazione</p>
        <p><a href="contatti.asp" class="info">Contatti e riferimenti</a></p>
        <p class="note"><strong>Cristalensi Snc</strong><br />
            di Lensi Massimiliano & C.<br />
            C.F. e Iscr. Reg. Impr. di Firenze 05305820481<br />
            R.E.A. Firenze 536760<br />
            50056 Montelupo F.no (FI)<br />
            Via arti e mestieri, 1<br />
            Tel. e Fax: 0571.911163<br />
            E-mail: <a href="mailto:info@cristalensi.it">info@cristalensi.it</a>
        </p>
        <img src="images/telefono_cristalensi.png" align="absmiddle" style="padding:10px 0 30px 0;" alt="Numero per chiamare lo staff del negozio Cristalensi, orario negozio dal Lunedi al Sabato: 0571.911163" />
        <h3>Seguici anche su</h3>
        <a href="http://www.facebook.com/pages/Cristalensi-vendita-lampade-per-interni-ed-esterni/144109972402284" target="_blank" title="Pagina ufficiale Cristalensi"><img src="images/facebook.png" align="absmiddle" border="0" alt="Collegati alla nostra pagina su Facebook" style="padding-bottom: 20px;" /></a>
    </div>
</div>
