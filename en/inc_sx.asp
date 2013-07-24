<div id="sidebar-alt" class="clearfix">
    <div>
        <%if idsession>0 then%>
            <h3>Client Area</h3>
            <p><font color="#990000"><strong>Welcome<br />&nbsp;<%=nome_log%></strong></font></p>
            <p>&raquo;<a href="iscrizione.asp">The details of your registration</a></p>
            <p>&raquo;<a href="ordini_elenco.asp">Your orders</a></p>
            <p>&raquo;<a href="commenti_form.asp">Post a comment</a></p>
            <p>&raquo;<a href="/admin/logout.asp">Exit from your profile</a></p>
        <%else%>
            <%if Instr(MM_LoginAction, "iscrizione")>0 then%>
            <%else%>
                <h3>Client Area</h3>
                <p>To see the catalogue is not necessary to register</p>
                <p><a href="#" class="info tiptip" title="<h3> General Information </ h3> Registration in the website Cristalensi <strong> is only compulsory to buy </ strong> but not to take a look at the catalogue of products. <br /> Clients, as well as be able to buy the produced by only entering <strong style='color: red'> Login </ strong> and <strong style='color: red'> Password </ strong> without re-register each time, will be able to print orders, update their data, post the comments on the website and will be updated on our offers. <br /> for all other information relating to the conditions of sale refer to the specific page <a href='condizioni_di_vendita.asp' title='conditions of sale' style='color: red'>conditions of sale</a>.">More information</a></p>
                <form method="post" name="logon" action="<%=MM_LoginAction%><%if Request.QueryString<>"" then%>&<%else%>?<%end if%>contr=1">

                   <label for="username">Login</label>
                   <input type="text" name="login">
                   <label for="password">Password</label>
                   <input type="password" name="password">
                   <button type="submit" class="button_link" style="margin-top:10px;">Submit</button>
               <%if contr=2 then%>
               <p style="background: #900; color: #fff; font-weight: bold; padding: 5px;">Warning! The data included <br />are incorrect.</p>
               <%end if%>
               </form>
               <p><a href="recupero_pw.asp" class="password-recover" title="Recover Lost Password">Recover lost Password</a></p>
               <a href="iscrizione.asp" title="Register to buy our lighting products"><button class="register button_link_red">SIGN IN!</button></a>

               <h3>Payment</h3>
               <p><a href="condizioni_di_vendita.asp" class="info" title="Conditions of sale">Conditions of sale</a></p>
               <p class="note">You can pay for your chosen products by Bank Transfer, Cash On Delivery, or directly on line with the secure PayPal system.</p>
               <img class="negozio" src="/images/cartedicredito.jpg" style="padding-bottom:30px;" title="Cards with which you can pay with PayPal" />
            <%end if%>
        <%end if%>

        <h3>Contact us!</h3>
        <p>Our staff will be pleased to answer any questions.</p>
        <p><a href="contatti.asp" class="info" title="Lighthing Showromm in Florence">Contatti e riferimenti</a></p>
        <p class="note"><strong>Cristalensi Snc</strong><br />
            di Lensi Massimiliano & C.<br />
            C.F. e Iscr. Reg. Impr. di Florence 05305820481<br />
            R.E.A. Firenze 536760<br />
            Florence (Italy)<br />
            50056 Montelupo F.no<br />
            Via arti e mestieri, 1<br />
            Tel. e Fax: 0571.911163<br />
            E-mail: <a href="mailto:info@cristalensi.it">info@cristalensi.it</a>
        </p>
        
        <h3>Follow us on Facebook</h3>
        <a href="http://www.facebook.com/pages/Cristalensi-vendita-lampade-per-interni-ed-esterni/144109972402284" target="_blank" title="Official page of Cristalensi"><img class="fb" src="/images/facebook.png" align="absmiddle" border="0" alt="Follow us on Facebook" /></a>
    </div>
</div>
